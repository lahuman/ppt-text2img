import io
import struct
import unittest
from xml.dom import minidom
from zipfile import ZipFile

from embedded_fonts import A, P, R, PRESENTATION, RELS, inspect_fonts, write_replaced_copy


class FontPreflightTests(unittest.TestCase):
    def package(self, restriction):
        buf = io.BytesIO()
        with ZipFile(buf, "w") as z:
            z.writestr(PRESENTATION, f'<p:presentation xmlns:p="{P}" xmlns:r="{R}">'
                       '<p:sldIdLst><p:sldId r:id="s2"/><p:sldId r:id="s1"/></p:sldIdLst>'
                       '<p:embeddedFontLst><p:embeddedFont><p:font typeface="Example"/>'
                       '<p:bold r:id="f1"/></p:embeddedFont><p:embeddedFont><p:font typeface="Keep"/>'
                       '<p:regular r:id="f2"/></p:embeddedFont></p:embeddedFontLst></p:presentation>')
            z.writestr(RELS, '<Relationships>' + ''.join(f'<Relationship Id="{rid}" Target="{part}"/>'
                       for rid, part in [('s1','slides/slide1.xml'), ('s2','slides/slide2.xml'),
                                         ('f1','fonts/1.fntdata'), ('f2','fonts/2.fntdata')]) + '</Relationships>')
            for i, rights in enumerate((restriction, 8), 1):
                z.writestr(f'ppt/fonts/{i}.fntdata', bytes(32) + struct.pack('<HH', rights, 0x504c))
            z.writestr('ppt/slides/slide1.xml', '<slide/>')
            z.writestr('ppt/slides/slide2.xml', f'<a:rPr xmlns:a="{A}"><a:latin typeface="Example"/></a:rPr>')
            z.writestr('ppt/slides/_rels/slide1.xml.rels', '<Relationships><Relationship Id="layout" '
                       'Type="rel/slideLayout" Target="../slideLayouts/layout.xml"/></Relationships>')
            z.writestr('ppt/slideLayouts/layout.xml', f'<a:ea xmlns:a="{A}" typeface="Example"/>')
            z.writestr('[Content_Types].xml', '<Types><Override PartName="/ppt/fonts/1.fntdata"/></Types>')
        return buf

    def test_order_shared_usage_and_copy(self):
        src = self.package(4)
        before = src.getvalue()
        dst = io.BytesIO()
        with ZipFile(src) as z:
            faces = inspect_fonts(z, lambda: None)
            self.assertEqual(set(faces), {"Example"})
            self.assertEqual(faces["Example"].slides, [1])
            self.assertEqual(faces["Example"].related_slides, [2])
            write_replaced_copy(z, dst, faces, "Arial", lambda: None)
        self.assertEqual(src.getvalue(), before)
        with ZipFile(dst) as z:
            self.assertFalse(inspect_fonts(z, lambda: None))
            self.assertNotIn("ppt/fonts/1.fntdata", z.namelist())
            self.assertIn("ppt/fonts/2.fntdata", z.namelist())
            self.assertNotIn(b'Id="f1"', z.read(RELS))
            el = minidom.parseString(z.read("ppt/slides/slide2.xml")).documentElement
            self.assertEqual(el.getAttribute("b"), "1")
            self.assertEqual(el.firstChild.getAttribute("typeface"), "Arial")
            self.assertIn(b"Arial", z.read("ppt/slideLayouts/layout.xml"))

    def test_editable_fonts_do_not_prompt(self):
        for rights in (0, 8, 0x108):
            with self.subTest(rights=rights), ZipFile(self.package(rights)) as z:
                self.assertFalse(inspect_fonts(z, lambda: None))
