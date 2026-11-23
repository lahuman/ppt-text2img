/* Set published to true only after tag 1.4 and both release assets are public. */
const release = Object.freeze({ version: "1.4", published: true, stableVersion: "1.4" });

(() => {
  "use strict";
  const repository = "https://github.com/lahuman/ppt-text2img";
  const elements = [...document.querySelectorAll("[data-i18n]")];
  const ko = Object.fromEntries(elements.map(el => [el.dataset.i18n, el.innerHTML]));
  const en = {
    skip: "Skip to content", navRelease: "Features", navWorkflow: "How it works", navDownload: "Download",
    heroTitle: "Convert PPT text<br><em>to images</em>",
    heroLead: "Text in the copy looks the same on PCs without the original fonts. You can still edit the original file.",
    heroCta: "Download v1.4 <span aria-hidden='true'>↗</span>",
    compatibility: "Windows 10 / 11 · Microsoft PowerPoint required",
    previewCaption: "v1.2 Windows portable app. Click to view the full-size screenshot.",
    releaseTitle: "Conversion features",
    feature1Title: "File path and progress",
    feature1Body: "Check the file path before starting. The app shows progress as a percentage. In a small window, scroll to see the controls.",
    feature2Title: "How shapes are replaced",
    feature2Body: "The app creates an image from a duplicate, then replaces the shape and adjusts its rotation and stacking order. If recovery fails, the app stops before saving.",
    feature3Title: "Warnings and errors",
    feature3Body: "The result screen lists conversion, warning and error counts. If there are warnings or errors, select Save&nbsp;report to save the details as a text file.",
    reportTitle: "Find a shape from the report",
    reportBody: "The app skips tables, charts and SmartArt. Check each item on the slide listed in the report.",
    issueExample: "Slide 3 · Shape «Table 2»<br>Warning · Check how the table looks.",
    reportCaption: "Example report entry. Find the shape by name in the PowerPoint Selection Pane (Alt+F10).",
    workflowTitle: "How to convert a file", step1Title: "Choose your PPT",
    step1Body: "Use Browse&nbsp;files or Ctrl+O to select a .ppt or .pptx presentation.",
    step2Title: "Start the conversion",
    step2Body: "Close the source file in PowerPoint. In the app, select Convert&nbsp;to&nbsp;images and review the confirmation dialog.",
    step3Title: "Check the result",
    step3Body: "Select Open&nbsp;result to check the copy. If you see a warning, review the slide listed in the report.",
    limitsTitle: "Before you convert", limit1Title: "PowerPoint and font requirements",
    limit1Body: "The Windows PC you use for conversion needs PowerPoint and access to the presentation’s fonts.",
    limit2Title: "Editing and searching text",
    limit2Body: "The entire shape containing the text becomes an image. Use the original PPT to edit or search the text.",
    limit3Title: "Groups and other objects",
    limit3Body: "Conversion ungroups shapes and skips tables, charts and SmartArt. Check slide masters, notes, animations and hyperlinks before sharing. The app may not preserve them.",
    limit4Title: "Font licenses",
    limit4Body: "The font’s use and distribution terms still apply after conversion.",
    downloadTitle: "Download for Windows",
    downloadLead: "Choose the portable version to run the app without installing it.",
    allReleases: "All releases on GitHub ↗", downloadCardTitle: "PPT Text to Image",
    downloadRequirement: "Requires Microsoft PowerPoint",
    checksums: "SHA-256 checksums", portableDownload: "Download portable version",
    releaseNotes: "Read the full release notes ↗", source: "Source code ↗"
  };

  function applyLanguage(lang) {
    if (lang !== "ko" && lang !== "en") lang = "ko";
    const dict = lang === "en" ? en : ko;
    document.documentElement.lang = lang;
    document.title = lang === "en" ? "PPT / PIXEL — PPT text to images" : "PPT / PIXEL — PPT 텍스트를 이미지로";
    elements.forEach(el => {
      const value = dict[el.dataset.i18n];
      if (value !== undefined) el.innerHTML = value;
    });
    document.querySelectorAll("[data-lang]").forEach(button => {
      button.setAttribute("aria-pressed", String(button.dataset.lang === lang));
    });
    const status = `v${release.version} · ${lang === "en"
      ? (release.published ? "Available now" : "Release in preparation")
      : (release.published ? "출시" : "출시 준비")}`;
    document.querySelectorAll("[data-release-status]").forEach(el => { el.textContent = status; });
    const version = release.published ? release.version : release.stableVersion;
    const download = document.querySelector("[data-download]");
    const suffix = `_v${version}`;
    download.href = `${repository}/releases/download/${version}/ppt-txt2img${suffix}.exe`;
    download.textContent = lang === "en"
      ? `${release.published ? "Download installer" : "Previous release"} v${version} ↓`
      : `${release.published ? "설치 프로그램 다운로드" : "기존 배포"} v${version}${release.published ? "" : " 다운로드"} ↓`;
    const portable = document.querySelector("[data-portable]");
    portable.hidden = false;
    portable.href = `${repository}/releases/download/${version}/ppt-txt2img-portable${suffix}.exe`;
    document.querySelector("[data-i18n=checksums]").href = `${repository}/releases/download/${version}/SHA256SUMS.txt`;
    const primary = document.querySelector("[data-primary-link]");
    primary.href = release.published ? "#download" : "#release";
    if (release.published) {
      primary.textContent = lang === "en" ? `Download v${release.version} ↗` : `v${release.version} 다운로드 ↗`;
    }
    try { localStorage.setItem("ppt-text2img-lang", lang); } catch { /* Storage may be disabled. */ }
  }

  document.querySelectorAll("[data-lang]").forEach(button => {
    button.addEventListener("click", () => applyLanguage(button.dataset.lang));
  });
  let initialLang = navigator.language?.toLowerCase().startsWith("en") ? "en" : "ko";
  try { initialLang = localStorage.getItem("ppt-text2img-lang") || initialLang; } catch { /* Use the browser language. */ }
  applyLanguage(initialLang);
})();
