"""Inspect PowerPoint embedded EOT fonts and substitute faces in a working PPTX."""
from dataclasses import dataclass, field
from pathlib import PurePosixPath
import posixpath
import shutil
import struct
from xml.dom import minidom
from zipfile import ZipFile

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PRESENTATION = "ppt/presentation.xml"
RELS = "ppt/_rels/presentation.xml.rels"


@dataclass
class RestrictedFace:
    styles: set = field(default_factory=set)
    slides: list = field(default_factory=list)
    shared: bool = False
    related_slides: list = field(default_factory=list)


def relationships(doc):
    return {e.getAttribute("Id"): e for e in doc.documentElement.childNodes
            if e.nodeType == e.ELEMENT_NODE and e.getAttribute("TargetMode") != "External"}


def target(rel, owner=PRESENTATION):
    name = rel.getAttribute("Target")
    return posixpath.normpath(posixpath.join(posixpath.dirname(owner), name)).lstrip("/")


def inspect_fonts(package, check_cancel):
    presentation = minidom.parseString(package.read(PRESENTATION))
    entries = presentation.getElementsByTagNameNS(P, "embeddedFont")
    faces = {}
    if not entries:
        return faces
    rels = relationships(minidom.parseString(package.read(RELS)))
    for entry in entries:
        check_cancel()
        family = entry.getElementsByTagNameNS(P, "font")[0].getAttribute("typeface")
        styles = set()
        restricted = False
        for style in entry.childNodes:
            if style.nodeType != style.ELEMENT_NODE or not style.hasAttributeNS(R, "id"):
                continue
            styles.add(style.localName)
            with package.open(target(rels[style.getAttributeNS(R, "id")])) as stream:
                header = stream.read(36)
            if len(header) == 36 and header[34:36] == b"LP":
                rights = struct.unpack_from("<H", header, 32)[0] & 15
                restricted |= bool(rights & 6 and not rights & 8)
        if restricted:
            faces[family] = RestrictedFace(styles)
    if not faces:
        return faces
    slide_order = {target(rels[e.getAttributeNS(R, "id")]): i for i, e in
                   enumerate(presentation.getElementsByTagNameNS(P, "sldId"), 1)}
    shared_parts = {}
    names = set(package.namelist())
    for name in names:
        if not name.startswith("ppt/") or not name.endswith(".xml") or name == PRESENTATION:
            continue
        check_cancel()
        doc = minidom.parseString(package.read(name))
        for family in {e.getAttribute("typeface") for e in doc.getElementsByTagNameNS(A, "*")} & faces.keys():
            if name in slide_order:
                faces[family].slides.append(slide_order[name])
            else:
                faces[family].shared = True
                shared_parts.setdefault(name, set()).add(family)
        doc.unlink()
    if shared_parts:
        # A shared theme or layout is a potential use, not a direct text assignment.
        links = {}
        for slide, number in slide_order.items():
            pending = [slide]
            visited = set()
            applicable = set()
            while pending:
                check_cancel()
                owner = pending.pop()
                if owner in visited:
                    continue
                visited.add(owner)
                applicable.update(shared_parts.get(owner, ()))
                if owner not in links:
                    part = PurePosixPath(owner)
                    relation_part = str(part.parent / "_rels" / (part.name + ".rels"))
                    links[owner] = []
                    if relation_part in names:
                        rel_doc = minidom.parseString(package.read(relation_part))
                        for rel in relationships(rel_doc).values():
                            kind = rel.getAttribute("Type").rsplit("/", 1)[-1]
                            if kind == "slideLayout" and owner != slide:
                                continue  # Masters also list unused sibling layouts.
                            if kind in {"slideLayout", "slideMaster", "theme", "themeOverride",
                                        "chart", "diagramData", "diagramDrawing", "chartUserShapes"}:
                                links[owner].append(target(rel, owner))
                        rel_doc.unlink()
                pending.extend(links[owner])
            for family in applicable:
                faces[family].related_slides.append(number)
    return faces


def write_replaced_copy(package, destination, faces, replacement, check_cancel):
    if not replacement.strip() or replacement.casefold() in {name.casefold() for name in faces}:
        raise ValueError("Invalid replacement font")
    doc = minidom.parseString(package.read(PRESENTATION))
    rel_doc = minidom.parseString(package.read(RELS))
    rels = relationships(rel_doc)
    removed_ids = set()
    for entry in list(doc.getElementsByTagNameNS(P, "embeddedFont")):
        family = entry.getElementsByTagNameNS(P, "font")[0].getAttribute("typeface")
        if family in faces:
            for child in entry.childNodes:
                if child.nodeType == child.ELEMENT_NODE and child.hasAttributeNS(R, "id"):
                    removed_ids.add(child.getAttributeNS(R, "id"))
            entry.parentNode.removeChild(entry)
    live_ids = {e.getAttributeNS(R, "id") for e in doc.getElementsByTagName("*")}
    removed_parts = set()
    for rid in removed_ids - live_ids:
        rel = rels.pop(rid)
        removed_parts.add(target(rel))
        rel.parentNode.removeChild(rel)
    removed_parts -= {target(rel) for rel in rels.values()}
    types = minidom.parseString(package.read("[Content_Types].xml"))
    for override in list(types.getElementsByTagNameNS("*", "Override")):
        if override.getAttribute("PartName").lstrip("/") in removed_parts:
            override.parentNode.removeChild(override)
    changes = {PRESENTATION: doc, RELS: rel_doc, "[Content_Types].xml": types}
    with ZipFile(destination, "w") as out:
        for item in package.infolist():
            check_cancel()
            name = item.filename
            if name in removed_parts:
                continue
            xml = changes.get(name)
            if name.startswith("ppt/") and name.endswith(".xml"):
                part = xml or minidom.parseString(package.read(name))
                for el in part.getElementsByTagNameNS(A, "*"):
                    face = faces.get(el.getAttribute("typeface"))
                    if face:
                        el.setAttribute("typeface", replacement)
                        for attr in ("panose", "pitchFamily", "charset"):
                            if el.hasAttribute(attr):
                                el.removeAttribute(attr)
                        parent = el.parentNode
                        if parent.localName in ("rPr", "defRPr", "endParaRPr"):
                            if face.styles <= {"bold", "boldItalic"}:
                                parent.setAttribute("b", "1")
                            if face.styles <= {"italic", "boldItalic"}:
                                parent.setAttribute("i", "1")
                        xml = part
                if xml is None:
                    part.unlink()
            if xml is not None:
                out.writestr(item, xml.toxml(encoding="utf-8"))
                xml.unlink()
            else:
                with package.open(item) as src, out.open(item, "w") as dst:
                    shutil.copyfileobj(src, dst)
