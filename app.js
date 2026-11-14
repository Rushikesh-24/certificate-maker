(() => {
  "use strict";

  const pdfjsLib = window.pdfjsLib;
  const { PDFDocument } = window.PDFLib;
  const XLSX = window.XLSX;
  const JSZip = window.JSZip;

  if (!pdfjsLib || !PDFDocument || !XLSX || !JSZip) {
    document.getElementById("statusMessage").textContent =
      "A required library could not be loaded. Check your internet connection and refresh.";
    document.getElementById("statusMessage").classList.add("error");
    return;
  }

  pdfjsLib.GlobalWorkerOptions.workerSrc =
    "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

  const elements = {
    pdfFile: document.getElementById("pdfFile"),
    pdfFileName: document.getElementById("pdfFileName"),
    sheetFile: document.getElementById("sheetFile"),
    sheetFileName: document.getElementById("sheetFileName"),
    nameColumn: document.getElementById("nameColumn"),
    rowCount: document.getElementById("rowCount"),
    fontFamily: document.getElementById("fontFamily"),
    fontHint: document.getElementById("fontHint"),
    titleCaseNames: document.getElementById("titleCaseNames"),
    fontSize: document.getElementById("fontSize"),
    textColor: document.getElementById("textColor"),
    colorValue: document.getElementById("colorValue"),
    pageGroup: document.getElementById("pageGroup"),
    pageNumber: document.getElementById("pageNumber"),
    generateButton: document.getElementById("generateButton"),
    statusMessage: document.getElementById("statusMessage"),
    emptyState: document.getElementById("emptyState"),
    previewArea: document.getElementById("previewArea"),
    pageFrame: document.getElementById("pageFrame"),
    pdfCanvas: document.getElementById("pdfCanvas"),
    nameOverlay: document.getElementById("nameOverlay"),
    previewName: document.getElementById("previewName"),
  };
  const textMeasureContext = document.createElement("canvas").getContext("2d");

  const state = {
    pdfBytes: null,
    pdfFileName: "",
    pdfDocument: null,
    pdfPage: null,
    pdfPageSize: null,
    previewScale: 1,
    headers: [],
    rows: [],
    selectedColumn: "",
    fontNames: [],
    position: { x: 0.5, y: 0.55 },
    isDragging: false,
  };

  function setStatus(message, isError = false) {
    elements.statusMessage.textContent = message;
    elements.statusMessage.classList.toggle("error", isError);
  }

  function updateGenerateButton() {
    elements.generateButton.disabled =
      !state.pdfBytes || !getSelectedNames().length || !state.selectedColumn;
  }

  function getSelectedNames() {
    return state.rows
      .map((row) => row[state.selectedColumn])
      .filter((value) => value !== null && value !== undefined && String(value).trim() !== "")
      .map((value) => {
        const name = String(value).trim();
        return elements.titleCaseNames.checked ? toTitleCase(name) : name;
      });
  }

  function toTitleCase(name) {
    return name.toLocaleLowerCase().replace(/(^|[\s'-])(\p{L})/gu, (match, boundary, letter) =>
      `${boundary}${letter.toLocaleUpperCase()}`,
    );
  }

  function updatePreviewName() {
    const names = getSelectedNames();
    elements.previewName.textContent = names[0] || "Your Name";
    elements.nameOverlay.style.left = `${state.position.x * 100}%`;
    elements.nameOverlay.style.top = `${state.position.y * 100}%`;
    updateOverlayStyle();
  }

  function updateOverlayStyle() {
    const fontSize = Math.min(300, Math.max(8, Number(elements.fontSize.value) || 48));
    const fontFamily = elements.fontFamily.value;
    let previewFontSize = fontSize * state.previewScale;
    textMeasureContext.font = `400 ${previewFontSize}px "${fontFamily}", sans-serif`;
    const measuredWidth = textMeasureContext.measureText(elements.previewName.textContent).width;
    const maxPreviewWidth = elements.pageFrame.clientWidth * 0.88;
    if (maxPreviewWidth > 0 && measuredWidth > maxPreviewWidth) {
      previewFontSize *= maxPreviewWidth / measuredWidth;
    }

    elements.nameOverlay.style.fontFamily = `"${fontFamily}", sans-serif`;
    elements.nameOverlay.style.fontSize = `${previewFontSize}px`;
    elements.nameOverlay.style.color = elements.textColor.value;
    elements.colorValue.textContent = elements.textColor.value.toUpperCase();
  }

  async function loadFontCatalog() {
    const response = await fetch("./google-fonts.json");
    if (!response.ok) throw new Error("The free Google Fonts list could not be loaded.");
    const fonts = await response.json();
    if (!Array.isArray(fonts) || !fonts.every((font) => typeof font === "string")) {
      throw new Error("The Google Fonts list has an invalid format.");
    }

    state.fontNames = fonts;
    const options = document.createDocumentFragment();
    for (const font of fonts) {
      const option = document.createElement("option");
      option.value = font;
      options.append(option);
    }
    document.getElementById("googleFontOptions").replaceChildren(options);
    elements.fontHint.textContent = `Search ${fonts.length.toLocaleString()} free Google Fonts.`;
  }

  async function loadGoogleFont(fontName) {
    const canonicalName = state.fontNames.find((font) => font.toLocaleLowerCase() === fontName.trim().toLocaleLowerCase());
    if (!canonicalName) {
      throw new Error("Choose a font from the Google Fonts list.");
    }

    elements.fontFamily.value = canonicalName;
    const existing = Array.from(document.querySelectorAll("link[data-google-font]"))
      .find((link) => link.dataset.googleFont === canonicalName);
    let stylesheet = existing;
    if (!stylesheet) {
      const link = document.createElement("link");
      link.rel = "stylesheet";
      link.dataset.googleFont = canonicalName;
      link.href = `https://fonts.googleapis.com/css2?family=${encodeURIComponent(canonicalName).replace(/%20/g, "+")}:wght@400&display=swap`;
      document.head.append(link);
      stylesheet = link;
    }

    if (!stylesheet.sheet) {
      await new Promise((resolve, reject) => {
        const timeout = window.setTimeout(() => reject(new Error("Google Fonts request timed out.")), 10000);
        stylesheet.addEventListener("load", () => {
          window.clearTimeout(timeout);
          resolve();
        }, { once: true });
        stylesheet.addEventListener("error", () => {
          window.clearTimeout(timeout);
          reject(new Error("Google Fonts stylesheet could not be loaded."));
        }, { once: true });
      });
    }

    await document.fonts.load(`400 16px "${canonicalName}"`);
    if (!document.fonts.check(`400 16px "${canonicalName}"`)) {
      throw new Error(`Google font "${canonicalName}" is not available.`);
    }
    updateOverlayStyle();
  }

  async function renderPage(pageNumber) {
    if (!state.pdfDocument) return;

    elements.emptyState.hidden = true;
    elements.previewArea.hidden = false;

    const page = await state.pdfDocument.getPage(pageNumber);
    const baseViewport = page.getViewport({ scale: 1 });
    const maxWidth = Math.max(120, elements.previewArea.clientWidth - 60);
    const maxHeight = Math.max(120, elements.previewArea.clientHeight - 60);
    const scale = Math.min(maxWidth / baseViewport.width, maxHeight / baseViewport.height);
    const viewport = page.getViewport({ scale });
    const canvas = elements.pdfCanvas;
    const context = canvas.getContext("2d", { alpha: false });
    const outputScale = window.devicePixelRatio || 1;

    canvas.width = Math.floor(viewport.width * outputScale);
    canvas.height = Math.floor(viewport.height * outputScale);
    canvas.style.width = `${viewport.width}px`;
    canvas.style.height = `${viewport.height}px`;
    elements.pageFrame.style.width = `${viewport.width}px`;
    elements.pageFrame.style.height = `${viewport.height}px`;
    elements.previewScale = scale;
    state.previewScale = viewport.width / baseViewport.width;
    state.pdfPage = page;
    state.pdfPageSize = { width: baseViewport.width, height: baseViewport.height };

    try {
      await page.render({
        canvasContext: context,
        viewport,
        transform: outputScale !== 1 ? [outputScale, 0, 0, outputScale, 0, 0] : null,
      }).promise;
    } catch (error) {
      elements.previewArea.hidden = true;
      elements.emptyState.hidden = false;
      throw error;
    }

    updatePreviewName();
  }

  async function handlePdfUpload(file) {
    if (!file) return;
    if (file.type !== "application/pdf" && !file.name.toLowerCase().endsWith(".pdf")) {
      setStatus("Please choose a PDF file.", true);
      return;
    }

    try {
      const bytes = new Uint8Array(await file.arrayBuffer());
      const pdfDocument = await pdfjsLib.getDocument({ data: bytes.slice() }).promise;
      state.pdfBytes = bytes;
      state.pdfFileName = file.name;
      state.pdfDocument = pdfDocument;
      elements.pdfFileName.textContent = file.name;
      elements.pageNumber.replaceChildren();

      for (let page = 1; page <= pdfDocument.numPages; page += 1) {
        const option = document.createElement("option");
        option.value = String(page);
        option.textContent = `Page ${page}`;
        elements.pageNumber.append(option);
      }

      elements.pageGroup.hidden = pdfDocument.numPages < 2;
      await renderPage(1);
      setStatus("");
      updateGenerateButton();
    } catch (error) {
      console.error("Could not read the PDF template.", error);
      setStatus("We couldn't open that PDF. Please try a valid, unprotected PDF.", true);
      state.pdfBytes = null;
      state.pdfDocument = null;
      updateGenerateButton();
    }
  }

  async function handleSpreadsheetUpload(file) {
    if (!file) return;
    const fileExtension = file.name.toLowerCase().split(".").pop();
    if (!["csv", "xls", "xlsx"].includes(fileExtension)) {
      setStatus("Please choose a CSV, XLS, or XLSX file.", true);
      return;
    }

    try {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", raw: false });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      if (!sheet) throw new Error("The spreadsheet does not contain a worksheet.");
      const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
      if (data.length < 2) throw new Error("Add a header row and at least one row of names.");

      const headers = data[0].map((value, index) => {
        const label = String(value).trim();
        return label || `Column ${index + 1}`;
      });
      const rows = data.slice(1).map((cells) =>
        Object.fromEntries(headers.map((header, index) => [header, cells[index] ?? ""])),
      );
      if (!rows.some((row) => headers.some((header) => String(row[header]).trim()))) {
        throw new Error("No names or other values were found below the header row.");
      }

      state.headers = headers;
      state.rows = rows;
      elements.sheetFileName.textContent = file.name;
      elements.nameColumn.replaceChildren();
      for (const header of headers) {
        const option = document.createElement("option");
        option.value = header;
        option.textContent = header;
        elements.nameColumn.append(option);
      }
      elements.nameColumn.disabled = false;
      state.selectedColumn = headers[0];
      elements.nameColumn.value = state.selectedColumn;
      const filledCount = getSelectedNames().length;
      elements.rowCount.textContent = `${filledCount} certificate${filledCount === 1 ? "" : "s"} ready from the first sheet.`;
      updatePreviewName();
      setStatus("");
      updateGenerateButton();
    } catch (error) {
      console.error("Could not read the spreadsheet.", error);
      setStatus(error instanceof Error ? error.message : "We couldn't read that spreadsheet.", true);
      state.rows = [];
      state.headers = [];
      state.selectedColumn = "";
      elements.nameColumn.disabled = true;
      elements.nameColumn.replaceChildren(new Option("Upload a spreadsheet first", ""));
      elements.rowCount.textContent = "Use a header row for your column names.";
      updatePreviewName();
      updateGenerateButton();
    }
  }

  function getSafeFileName(value, index) {
    const cleaned = value
      .normalize("NFKD")
      .replace(/[^\w.-]+/g, "_")
      .replace(/^\.|^\-+|_+$/g, "")
      .slice(0, 90);
    return cleaned || `certificate_${String(index + 1).padStart(3, "0")}`;
  }

  function hexToRgb(hex) {
    const value = hex.replace("#", "");
    return {
      r: Number.parseInt(value.slice(0, 2), 16) / 255,
      g: Number.parseInt(value.slice(2, 4), 16) / 255,
      b: Number.parseInt(value.slice(4, 6), 16) / 255,
    };
  }

  function createTextImage(name) {
    const { width, height } = state.pdfPageSize;
    const rasterScale = 3;
    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(width * rasterScale);
    canvas.height = Math.ceil(height * rasterScale);
    const context = canvas.getContext("2d");
    const fontSize = Math.min(300, Math.max(8, Number(elements.fontSize.value) || 48));
    const fontName = elements.fontFamily.value;
    const rgb = hexToRgb(elements.textColor.value);

    context.scale(rasterScale, rasterScale);
    context.fillStyle = `rgb(${Math.round(rgb.r * 255)}, ${Math.round(rgb.g * 255)}, ${Math.round(rgb.b * 255)})`;
    context.font = `400 ${fontSize}px "${fontName}", sans-serif`;
    context.textAlign = "center";
    context.textBaseline = "middle";

    const maxWidth = width * 0.88;
    let fittedSize = fontSize;
    while (context.measureText(name).width > maxWidth && fittedSize > 8) {
      fittedSize -= 1;
      context.font = `400 ${fittedSize}px "${fontName}", sans-serif`;
    }

    context.fillText(name, state.position.x * width, state.position.y * height, maxWidth);
    return canvas.toDataURL("image/png");
  }

  async function generateCertificates() {
    const names = getSelectedNames();
    if (!state.pdfBytes || !names.length || !state.selectedColumn) return;

    const buttonLabel = elements.generateButton.querySelector(".button-label");
    elements.generateButton.disabled = true;
    buttonLabel.textContent = "Preparing certificates…";
    setStatus(`Preparing ${names.length} certificate${names.length === 1 ? "" : "s"}…`);

    try {
      await loadGoogleFont(elements.fontFamily.value);
      const zip = new JSZip();
      const pageIndex = Number(elements.pageNumber.value || 1) - 1;
      const pageLabel = String(Number(elements.pageNumber.value || 1)).padStart(2, "0");

      for (let index = 0; index < names.length; index += 1) {
        const pdfDoc = await PDFDocument.load(state.pdfBytes);
        const pages = pdfDoc.getPages();
        if (!pages[pageIndex]) throw new Error("The selected PDF page no longer exists.");

        const pngDataUrl = createTextImage(names[index]);
        const imageBytes = Uint8Array.from(atob(pngDataUrl.split(",")[1]), (char) => char.charCodeAt(0));
        const image = await pdfDoc.embedPng(imageBytes);
        pages[pageIndex].drawImage(image, {
          x: 0,
          y: 0,
          width: pages[pageIndex].getWidth(),
          height: pages[pageIndex].getHeight(),
        });

        const pdfBytes = await pdfDoc.save();
        const filename = `${String(index + 1).padStart(3, "0")}_${getSafeFileName(names[index], index)}_page-${pageLabel}.pdf`;
        zip.file(filename, pdfBytes);
        setStatus(`Creating certificate ${index + 1} of ${names.length}…`);
      }

      const zipBlob = await zip.generateAsync({ type: "blob" });
      const link = document.createElement("a");
      const url = URL.createObjectURL(zipBlob);
      link.href = url;
      link.download = `${getSafeFileName(state.pdfFileName.replace(/\.pdf$/i, ""), 0)}_certificates.zip`;
      link.click();
      window.setTimeout(() => URL.revokeObjectURL(url), 1000);
      setStatus(`Done — ${names.length} certificate${names.length === 1 ? "" : "s"} downloaded.`);
    } catch (error) {
      console.error("Could not generate the certificates.", error);
      setStatus(error instanceof Error ? error.message : "We couldn't create the ZIP. Please try again.", true);
    } finally {
      buttonLabel.textContent = "Create certificates";
      updateGenerateButton();
    }
  }

  elements.pdfFile.addEventListener("change", (event) => {
    handlePdfUpload(event.target.files[0]);
  });

  elements.sheetFile.addEventListener("change", (event) => {
    handleSpreadsheetUpload(event.target.files[0]);
  });

  elements.nameColumn.addEventListener("change", () => {
    state.selectedColumn = elements.nameColumn.value;
    const count = getSelectedNames().length;
    elements.rowCount.textContent = `${count} certificate${count === 1 ? "" : "s"} ready from the first sheet.`;
    updatePreviewName();
    updateGenerateButton();
  });

  elements.pageNumber.addEventListener("change", () => {
    renderPage(Number(elements.pageNumber.value)).catch((error) => {
      console.error("Could not render the selected PDF page.", error);
      setStatus("We couldn't render that PDF page.", true);
    });
  });

  elements.fontFamily.addEventListener("change", () => {
    updateOverlayStyle();
    loadGoogleFont(elements.fontFamily.value).catch((error) => {
      console.error("Could not load the selected Google font.", error);
      setStatus(error instanceof Error ? error.message : "The selected Google font could not be loaded.", true);
    });
  });

  elements.titleCaseNames.addEventListener("change", () => {
    const count = getSelectedNames().length;
    elements.rowCount.textContent = `${count} certificate${count === 1 ? "" : "s"} ready from the first sheet.`;
    updatePreviewName();
  });

  elements.fontSize.addEventListener("input", updateOverlayStyle);
  elements.textColor.addEventListener("input", updateOverlayStyle);

  document.querySelectorAll(".stepper").forEach((button) => {
    button.addEventListener("click", () => {
      const current = Number(elements.fontSize.value) || 48;
      elements.fontSize.value = String(Math.min(300, Math.max(8, current + Number(button.dataset.step))));
      updateOverlayStyle();
    });
  });

  elements.generateButton.addEventListener("click", generateCertificates);

  elements.nameOverlay.addEventListener("pointerdown", (event) => {
    if (!elements.previewArea.hidden) {
      state.isDragging = true;
      elements.nameOverlay.setPointerCapture(event.pointerId);
      event.preventDefault();
    }
  });

  elements.nameOverlay.addEventListener("pointermove", (event) => {
    if (!state.isDragging) return;
    const bounds = elements.pageFrame.getBoundingClientRect();
    state.position.x = Math.min(0.98, Math.max(0.02, (event.clientX - bounds.left) / bounds.width));
    state.position.y = Math.min(0.98, Math.max(0.02, (event.clientY - bounds.top) / bounds.height));
    updatePreviewName();
  });

  elements.nameOverlay.addEventListener("pointerup", () => {
    state.isDragging = false;
  });

  elements.nameOverlay.addEventListener("pointercancel", () => {
    state.isDragging = false;
  });

  elements.nameOverlay.addEventListener("keydown", (event) => {
    const movement = event.shiftKey ? 0.02 : 0.005;
    const movements = {
      ArrowLeft: [-movement, 0],
      ArrowRight: [movement, 0],
      ArrowUp: [0, -movement],
      ArrowDown: [0, movement],
    };
    if (!movements[event.key]) return;
    event.preventDefault();
    state.position.x = Math.min(0.98, Math.max(0.02, state.position.x + movements[event.key][0]));
    state.position.y = Math.min(0.98, Math.max(0.02, state.position.y + movements[event.key][1]));
    updatePreviewName();
  });

  let resizeTimeout;
  window.addEventListener("resize", () => {
    if (!state.pdfDocument) return;
    window.clearTimeout(resizeTimeout);
    resizeTimeout = window.setTimeout(() => {
      renderPage(Number(elements.pageNumber.value || 1)).catch((error) => {
        console.error("Could not resize the PDF preview.", error);
        setStatus("The PDF preview couldn't be resized.", true);
      });
    }, 120);
  });

  loadFontCatalog().catch((error) => {
    console.error("Could not load the Google Fonts catalog.", error);
    elements.fontHint.textContent = "The Google Fonts list could not be loaded.";
    setStatus(error instanceof Error ? error.message : "The Google Fonts list could not be loaded.", true);
  });
})();
