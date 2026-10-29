/* Set published to true only after tag 1.2 and both release assets are public. */
const release = Object.freeze({ version: "1.2", published: false, stableVersion: "1.1" });

(() => {
  "use strict";
  const repository = "https://github.com/lahuman/ppt-text2img";
  const elements = [...document.querySelectorAll("[data-i18n]")];
  const ko = Object.fromEntries(elements.map(el => [el.dataset.i18n, el.innerHTML]));
  const en = {
    skip: "Skip to content", navRelease: "What’s new", navWorkflow: "How it works", navDownload: "Download",
    heroTitle: "Different fonts.<br><em>The same impression.</em>",
    heroLead: "Prepare a delivery copy that travels with its design. Convert the text in your PPT into images before sharing it with another PC.",
    heroCta: "Explore v1.2 <span aria-hidden='true'>↗</span>",
    compatibility: "Windows 10 / 11 · Microsoft PowerPoint required",
    previewSide: "Keep the details.<br>Share your vision.",
    previewSideBody: "Make your presentation look right,<br>even on a PC without your fonts.",
    previewStep1: "Choose a presentation", previewStep2: "Turn text into images", previewStep3: "Save as a separate file",
    localNote: "Your files stay on your PC", previewWorkspace: "PRESENTATION CONVERTER", previewLanguage: "English ▾",
    previewDescription: "Turn text into images so your presentation travels with its design intact.",
    selectedFile: "Selected file", browse: "Choose another file", progress: "Progress",
    previewProgress: "Processing slide 3/5…", previewNotice: "Close the source file in PowerPoint before converting.<br>Converted text will no longer be directly editable.",
    convert: "Converting…", previewOutput: "The converted file will be saved next to the original file.",
    previewCaption: "v1.2 interface illustration. The filename and progress are examples, not a screenshot.",
    principle1: "Processed on your PC", principle2: "A separate delivery copy", principle3: "한국어 · English",
    releaseTitle: "A clearer workspace.<br>A clearer result.",
    releaseLead: "A new interface, safer shape replacement, and feedback when a conversion is only partially complete.",
    feature1Title: "Review the file before you start",
    feature1Body: "File selection and conversion are separate actions. Check the path, then start when you are ready. Numeric progress and a scrollable workspace keep the task visible.",
    feature2Title: "Keep the original shape until replacement",
    feature2Body: "Images are generated from duplicates. The original shape is replaced only after success. Rotation and stacking are adjusted; saving stops if temporary objects cannot be safely removed.",
    feature3Title: "Partial completion gets its own warning",
    feature3Body: "If an item cannot be converted, the app warns you. A JSON report next to the output lists conversion counts and the slides and processing stages that need review.",
    reportTitle: "Know where to take a closer look.",
    reportBody: "Tables, charts, and SmartArt are not automatically converted. They are recorded as items that need manual review.",
    reportCaption: "Illustrative report only. These are not results from a real document.",
    workflowTitle: "Choose. Review. Convert.", step1Title: "Choose your PPT",
    step1Body: "Use Browse files or Ctrl+O to select a .ppt or .pptx presentation.",
    step2Title: "Start the conversion",
    step2Body: "Close the source file in PowerPoint, select Convert to images, and review the confirmation dialog.",
    step3Title: "Check the result",
    step3Body: "Find the new PPT and report in the original folder. If you see a partial-completion warning, review the listed slides.",
    limitsTitle: "Before you convert.", limit1Title: "PowerPoint and the original fonts are required",
    limit1Body: "The app uses Microsoft PowerPoint installed on Windows. The conversion PC must have access to the original fonts, and the source PPT must be closed.",
    limit2Title: "Converted text is no longer directly editable",
    limit2Body: "The entire shape containing text becomes an image. Direct text editing and search are limited. Keep the editable original and use the converted copy for delivery.",
    limit3Title: "Not every object or group structure is preserved",
    limit3Body: "Groups are ungrouped. Tables, charts, and SmartArt are not automatically converted. Preserving masters, notes, animations, and hyperlinks is outside the supported conversion scope. Check the output before sharing.",
    limit4Title: "Font license conditions still apply",
    limit4Body: "Turning text into images does not waive the font’s use or distribution license conditions.",
    downloadTitle: "Prepare the copy<br>you’ll present.",
    downloadLead: "The new interface and reliability improvements are in the v1.2 source. Until it is published, the download below points to the previous release.",
    allReleases: "Check GitHub Releases ↗", downloadCardTitle: "PPT Text to Image",
    downloadRequirement: "For Windows PCs with Microsoft PowerPoint installed",
    portableDownload: "Download the portable executable",
    downloadNote: "The v1.2 interface and improvements shown above are not included in the v1.1 download.",
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
    download.href = `${repository}/releases/download/${version}/ppt-txt2img.exe`;
    download.textContent = lang === "en"
      ? `${release.published ? "Download" : "Previous release"} v${version} ↓`
      : `${release.published ? "다운로드" : "기존 배포"} v${version}${release.published ? "" : " 다운로드"} ↓`;
    const portable = document.querySelector("[data-portable]");
    portable.hidden = !release.published;
    portable.href = `${repository}/releases/download/${release.version}/ppt-txt2img-portable.exe`;
    const primary = document.querySelector("[data-primary-link]");
    primary.href = release.published ? "#download" : "#release";
    if (release.published) {
      primary.textContent = lang === "en" ? `Download v${release.version} ↗` : `v${release.version} 다운로드 ↗`;
      document.querySelector("[data-i18n=downloadLead]").textContent = lang === "en"
        ? "The new interface and reliability improvements are now available in v1.2."
        : "새 UI와 안정성 개선을 담은 v1.2를 사용할 수 있습니다.";
      document.querySelector("[data-i18n=downloadNote]").textContent = lang === "en"
        ? "Choose the installer or the portable executable. Both require Microsoft PowerPoint."
        : "설치 프로그램 또는 설치 없는 실행 파일을 선택하세요. 두 방식 모두 Microsoft PowerPoint가 필요합니다.";
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
