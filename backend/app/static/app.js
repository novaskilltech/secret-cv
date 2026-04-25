const API = {
  merge: "/api/merge",
  split: "/api/split",
  reorder: "/api/reorder",
  rotate: "/api/rotate",
  crop: "/api/crop",
  compress: "/api/compress",
  repair: "/api/repair",
  convert: "/api/convert/image-to-pdf",
  "html-to-pdf": "/api/convert/html-to-pdf",
  "word-to-pdf": "/api/convert/word-to-pdf",
  "excel-to-pdf": "/api/convert/excel-to-pdf",
  "powerpoint-to-pdf": "/api/convert/powerpoint-to-pdf",
  delete: "/api/delete",
  extract: "/api/extract",
  ocr: "/api/ocr",
  watermark: "/api/watermark",
  "pdf-to-jpg": "/api/pdf-to-jpg",
  "pdf-to-word": "/api/pdf-to-word",
  "pdf-to-excel": "/api/pdf-to-excel",
  "pdf-to-powerpoint": "/api/pdf-to-powerpoint",
  numbering: "/api/numbering",
  protect: "/api/protect",
  unlock: "/api/unlock",
  compare: "/api/compare",
  censor: "/api/censor",
};

const PAGE_RANGE_PATTERN = /^\s*\d+\s*(?:-\s*\d+\s*)?(?:,\s*\d+\s*(?:-\s*\d+\s*)?)*\s*$/;

async function submitForm(action) {
  const form = new FormData();
  setBusy(true);
  setMessage("Traitement en cours...", "info");

  try {
    switch (action) {
      case "merge":
        appendMultipleFiles(form, "files", "mergeFiles", 2, "Selectionnez au moins deux PDF.");
        await submitRequest(API.merge, form, "merged.pdf");
        break;
      case "split":
        form.append("file", getRequiredFile("splitFile", "Selectionnez un PDF."));
        form.append("pages", getPageRanges("splitPages"));
        await submitRequest(API.split, form, "splitted.pdf");
        break;
      case "reorder":
        form.append("file", getRequiredFile("reorderFile", "Selectionnez un PDF."));
        form.append("pages", getRequiredText("reorderPages", "Indiquez l'ordre des pages."));
        await submitRequest(API.reorder, form, "reordered.pdf");
        break;
      case "rotate": {
        form.append("file", getRequiredFile("rotateFile", "Selectionnez un PDF."));
        form.append("angle", document.getElementById("rotateAngle").value);
        const pages = document.getElementById("rotatePages").value.trim();
        if (pages) {
          validatePageRanges(pages);
          form.append("pages", pages);
        }
        await submitRequest(API.rotate, form, "rotated.pdf");
        break;
      }
      case "crop":
        form.append("file", getRequiredFile("cropFile", "Selectionnez un PDF."));
        form.append("top", getNumericValue("cropTop"));
        form.append("right", getNumericValue("cropRight"));
        form.append("bottom", getNumericValue("cropBottom"));
        form.append("left", getNumericValue("cropLeft"));
        await submitRequest(API.crop, form, "cropped.pdf");
        break;
      case "compress":
        form.append("file", getRequiredFile("compressFile", "Selectionnez un PDF."));
        await submitRequest(API.compress, form, "compressed.pdf");
        break;
      case "repair":
        form.append("file", getRequiredFile("repairFile", "Selectionnez un PDF."));
        await submitRequest(API.repair, form, "repaired.pdf");
        break;
      case "convert":
        form.append("file", getRequiredFile("imageFile", "Selectionnez une image."));
        await submitRequest(API.convert, form, "converted.pdf");
        break;
      case "word-to-pdf":
        form.append("file", getRequiredFile("wordToPdfFile", "Selectionnez un DOCX."));
        await submitRequest(API["word-to-pdf"], form, "word-converted.pdf");
        break;
      case "excel-to-pdf":
        form.append("file", getRequiredFile("excelToPdfFile", "Selectionnez un XLSX."));
        await submitRequest(API["excel-to-pdf"], form, "excel-converted.pdf");
        break;
      case "powerpoint-to-pdf":
        form.append("file", getRequiredFile("powerpointToPdfFile", "Selectionnez un PPTX."));
        await submitRequest(API["powerpoint-to-pdf"], form, "powerpoint-converted.pdf");
        break;
      case "html-to-pdf":
        form.append("file", getRequiredFile("htmlToPdfFile", "Selectionnez un fichier HTML."));
        await submitRequest(API["html-to-pdf"], form, "html-converted.pdf");
        break;
      case "delete":
        form.append("file", getRequiredFile("deleteFile", "Selectionnez un PDF."));
        form.append("pages", getPageRanges("deletePages"));
        await submitRequest(API.delete, form, "deleted.pdf");
        break;
      case "extract":
        form.append("file", getRequiredFile("extractFile", "Selectionnez un PDF."));
        form.append("pages", getPageRanges("extractPages"));
        await submitRequest(API.extract, form, "extracted.pdf");
        break;
      case "ocr":
        form.append("file", getRequiredFile("ocrFile", "Selectionnez un PDF."));
        await submitRequest(API.ocr, form, "ocr.txt");
        break;
      case "watermark":
        form.append("file", getRequiredFile("watermarkFile", "Selectionnez un PDF."));
        form.append("text", getRequiredText("watermarkText", "Entrez un texte de filigrane."));
        form.append("opacity", document.getElementById("watermarkOpacity").value);
        await submitRequest(API.watermark, form, "watermarked.pdf");
        break;
      case "pdf-to-jpg":
        form.append("file", getRequiredFile("pdfToJpgFile", "Selectionnez un PDF."));
        await submitRequest(API["pdf-to-jpg"], form, "images.zip");
        break;
      case "pdf-to-word":
        form.append("file", getRequiredFile("pdfToWordFile", "Selectionnez un PDF."));
        await submitRequest(API["pdf-to-word"], form, "converted.docx");
        break;
      case "pdf-to-excel":
        form.append("file", getRequiredFile("pdfToExcelFile", "Selectionnez un PDF."));
        await submitRequest(API["pdf-to-excel"], form, "converted.xlsx");
        break;
      case "pdf-to-powerpoint":
        form.append("file", getRequiredFile("pdfToPowerPointFile", "Selectionnez un PDF."));
        await submitRequest(API["pdf-to-powerpoint"], form, "converted.pptx");
        break;
      case "numbering":
        form.append("file", getRequiredFile("numberingFile", "Selectionnez un PDF."));
        form.append("format_str", getRequiredText("numberingFormat", "Entrez un format de numerotation."));
        form.append("position", document.getElementById("numberingPosition").value);
        await submitRequest(API.numbering, form, "numbered.pdf");
        break;
      case "unlock":
        form.append("file", getRequiredFile("unlockFile", "Selectionnez un PDF."));
        form.append("password", getRequiredText("unlockPassword", "Entrez le mot de passe du PDF."));
        await submitRequest(API.unlock, form, "unlocked.pdf");
        break;
      case "protect":
        form.append("file", getRequiredFile("protectFile", "Selectionnez un PDF."));
        form.append("user_password", getRequiredText("protectUserPassword", "Entrez un mot de passe utilisateur."));
        const ownerPassword = document.getElementById("protectOwnerPassword").value.trim();
        if (ownerPassword) {
          form.append("owner_password", ownerPassword);
        }
        await submitRequest(API.protect, form, "protected.pdf");
        break;
      case "compare":
        form.append("file_a", getRequiredFile("compareFileA", "Selectionnez le premier PDF."));
        form.append("file_b", getRequiredFile("compareFileB", "Selectionnez le second PDF."));
        await submitRequest(API.compare, form, "compare-report.json");
        break;
      case "censor":
        form.append("file", getRequiredFile("censorFile", "Selectionnez un PDF."));
        form.append("terms", getRequiredText("censorTerms", "Indiquez un ou plusieurs termes a censurer."));
        form.append("case_sensitive", document.getElementById("censorCaseSensitive").checked ? "true" : "false");
        await submitRequest(API.censor, form, "censored.pdf");
        break;
      default:
        throw new Error("Action inconnue.");
    }
  } catch (error) {
    setMessage(error.message || "Erreur inconnue.", "error");
  } finally {
    setBusy(false);
  }
}

async function submitRequest(url, formData, fallbackFilename) {
  const response = await fetch(url, {
    method: "POST",
    body: formData,
  });

  if (!response.ok) {
    const error = await response.json().catch(() => null);
    throw new Error(error?.detail || "Erreur serveur.");
  }

  const blob = await response.blob();
  const filename = getDownloadFilename(response.headers.get("Content-Disposition")) || fallbackFilename;
  downloadBlob(blob, filename);
  setMessage(`Telechargement lance : ${filename}`, "success");
}

function appendMultipleFiles(form, fieldName, inputId, minCount, errorMessage) {
  const files = document.getElementById(inputId).files;
  if (!files || files.length < minCount) {
    throw new Error(errorMessage);
  }
  for (const file of files) {
    form.append(fieldName, file);
  }
}

function getRequiredFile(inputId, errorMessage) {
  const file = document.getElementById(inputId).files[0];
  if (!file) {
    throw new Error(errorMessage);
  }
  return file;
}

function getRequiredText(inputId, errorMessage) {
  const value = document.getElementById(inputId).value.trim();
  if (!value) {
    throw new Error(errorMessage);
  }
  return value;
}

function getNumericValue(inputId) {
  const value = document.getElementById(inputId).value.trim();
  return value === "" ? "0" : value;
}

function getPageRanges(inputId) {
  const value = getRequiredText(inputId, "Indiquez une ou plusieurs pages.");
  validatePageRanges(value);
  return value;
}

function validatePageRanges(value) {
  if (!PAGE_RANGE_PATTERN.test(value)) {
    throw new Error("Format de pages invalide. Exemple attendu : 1,3-5");
  }
}

function getDownloadFilename(contentDisposition) {
  if (!contentDisposition) {
    return null;
  }
  const match = contentDisposition.match(/filename="?([^"]+)"?/i);
  return match ? match[1] : null;
}

function downloadBlob(blob, filename) {
  const link = document.createElement("a");
  const objectUrl = URL.createObjectURL(blob);
  link.href = objectUrl;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(objectUrl);
}

function setMessage(text, type = "info") {
  const message = document.getElementById("message");
  message.textContent = text;
  message.className = type;
}

function setBusy(isBusy) {
  document.querySelectorAll("[data-action]").forEach((button) => {
    button.disabled = isBusy;
  });
}

function setupButtons() {
  document.querySelectorAll("[data-action]").forEach((button) => {
    button.addEventListener("click", () => submitForm(button.dataset.action));
  });
}

function setupOpacitySlider() {
  const slider = document.getElementById("watermarkOpacity");
  const label = document.getElementById("opacityLabel");
  if (!slider || !label) {
    return;
  }
  const syncLabel = () => {
    label.textContent = `Opacite: ${slider.value}`;
  };
  slider.addEventListener("input", syncLabel);
  syncLabel();
}

function setupDropzones() {
  document.querySelectorAll("[data-dropzone]").forEach((dropzone) => {
    const inputId = dropzone.dataset.dropzone;
    const input = document.getElementById(inputId);
    if (!input) {
      return;
    }

    const refreshLabel = () => {
      const label = document.querySelector(`[data-file-label="${inputId}"]`);
      if (!label) {
        return;
      }
      if (!input.files || input.files.length === 0) {
        label.textContent = "Aucun fichier selectionne";
        return;
      }
      if (input.multiple) {
        label.textContent = input.files.length === 1 ? input.files[0].name : `${input.files.length} fichiers selectionnes`;
        return;
      }
      label.textContent = input.files[0].name;
    };

    input.addEventListener("change", refreshLabel);
    refreshLabel();

    ["dragenter", "dragover"].forEach((eventName) => {
      dropzone.addEventListener(eventName, (event) => {
        event.preventDefault();
        dropzone.classList.add("dragover");
      });
    });

    ["dragleave", "dragend", "drop"].forEach((eventName) => {
      dropzone.addEventListener(eventName, (event) => {
        event.preventDefault();
        dropzone.classList.remove("dragover");
      });
    });

    dropzone.addEventListener("drop", (event) => {
      const droppedFiles = Array.from(event.dataTransfer?.files || []);
      if (!droppedFiles.length) {
        return;
      }
      const nextFiles = input.multiple ? droppedFiles : [droppedFiles[0]];
      const transfer = new DataTransfer();
      nextFiles.forEach((file) => transfer.items.add(file));
      input.files = transfer.files;
      refreshLabel();
    });
  });
}

document.addEventListener("DOMContentLoaded", () => {
  setupButtons();
  setupOpacitySlider();
  setupDropzones();
});
