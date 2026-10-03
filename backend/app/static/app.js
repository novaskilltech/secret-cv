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
  sign: "/api/sign",
  summarize: "/api/ai/summarize",
  translate: "/api/ai/translate",
};

const PRO_ACTIONS = new Set([
  "ocr",
  "word-to-pdf",
  "powerpoint-to-pdf",
  "excel-to-pdf",
  "pdf-to-word",
  "pdf-to-powerpoint",
  "pdf-to-excel",
  "watermark",
  "unlock",
  "protect",
  "compare",
  "sign",
  "censor",
  "summarize",
  "translate",
]);

const DAILY_FREE_LIMIT = 3;

function isPro() {
  try {
    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams.get("status") === "pro" || urlParams.get("status") === "unlocked") {
      localStorage.setItem("nova_pro_status", "active");
    }
    return localStorage.getItem("nova_pro_status") === "active";
  } catch (_) {
    return false;
  }
}

function getTodayKey() {
  return new Date().toISOString().slice(0, 10);
}

function getDailyUsage() {
  try {
    const raw = localStorage.getItem("nova_daily_quota");
    const today = getTodayKey();
    if (raw) {
      const parsed = JSON.parse(raw);
      if (parsed && parsed.date === today) {
        return { date: today, count: Number(parsed.count) || 0 };
      }
    }
    const fresh = { date: today, count: 0 };
    localStorage.setItem("nova_daily_quota", JSON.stringify(fresh));
    return fresh;
  } catch (_) {
    return { date: getTodayKey(), count: 0 };
  }
}

function incrementDailyUsage() {
  try {
    const usage = getDailyUsage();
    usage.count += 1;
    localStorage.setItem("nova_daily_quota", JSON.stringify(usage));
    return usage;
  } catch (_) {
    return { date: getTodayKey(), count: 1 };
  }
}

function openPaywallModal(reason = "pro_feature") {
  const modal = document.getElementById("paywallModal");
  const title = document.getElementById("modalTitle");
  const desc = document.getElementById("modalDesc");
  const feedback = document.getElementById("keyFeedback");
  if (feedback) feedback.style.display = "none";

  if (!modal || !title || !desc) return;

  if (reason === "quota_reached") {
    title.textContent = "Quota journalier atteint (3/3)";
    desc.innerHTML = "Vous avez utilisé vos <strong>3 opérations gratuites</strong> aujourd'hui. Débloquez le <strong>Pass Illimité</strong> pour continuer sans attendre demain.";
  } else if (reason === "license_input") {
    title.textContent = "Activer votre Pass Illimité";
    desc.innerHTML = "Entrez la <strong>Clé Pass Nova</strong> reçue lors de votre commande pour débloquer tous les outils sur cet appareil.";
  } else {
    title.textContent = "Fonctionnalité réservée au Pass Illimité";
    desc.innerHTML = "Cette fonctionnalité avancée requiert un <strong>Pass Illimité</strong>. Débloquez instantanément tous les outils avancés sans quota.";
  }

  modal.classList.add("open");
  modal.setAttribute("aria-hidden", "false");
}

function closePaywallModal() {
  const modal = document.getElementById("paywallModal");
  if (modal) {
    modal.classList.remove("open");
    modal.setAttribute("aria-hidden", "true");
  }
}

function updateQuotaUI() {
  const quotaContainer = document.getElementById("quotaContainer");
  const quotaText = document.getElementById("quotaText");
  const btnOpenPassModal = document.getElementById("btnOpenPassModal");
  const stripeBuyBtn = document.getElementById("stripe-buy-btn");
  const messageEl = document.getElementById("message");

  if (isPro()) {
    if (quotaContainer) {
      quotaContainer.className = "quota-badge pro-active";
    }
    if (quotaText) {
      quotaText.textContent = "👑 Pass Illimité Actif";
    }
    if (btnOpenPassModal) {
      btnOpenPassModal.style.display = "none";
    }
    if (stripeBuyBtn) {
      stripeBuyBtn.textContent = "📄 Factures & Support";
      stripeBuyBtn.href = "https://billing.stripe.com/p/login/00wbJ1dEKdR5b3Re1ggbm00";
      stripeBuyBtn.style.background = "#334155";
    }
    if (messageEl && (messageEl.textContent === "Pret." || messageEl.textContent.startsWith("Prêt"))) {
      messageEl.textContent = "Prêt — Mode Illimité Pro Actif (500 Mo / fichier, sans quota).";
      messageEl.className = "status-ok";
    }
  } else {
    const usage = getDailyUsage();
    if (quotaContainer) {
      quotaContainer.className = usage.count >= DAILY_FREE_LIMIT ? "quota-badge warning" : "quota-badge";
    }
    if (quotaText) {
      quotaText.textContent = `Quota gratuit : ${usage.count}/${DAILY_FREE_LIMIT} aujourd'hui`;
    }
  }
}

const PAGE_RANGE_PATTERN = /^\s*\d+\s*(?:-\s*\d+\s*)?(?:,\s*\d+\s*(?:-\s*\d+\s*)?)*\s*$/;
let progressTimer = null;

async function submitForm(action) {
  // Controle Freemium
  if (!isPro()) {
    if (PRO_ACTIONS.has(action)) {
      openPaywallModal("pro_feature");
      return;
    }
    const usage = getDailyUsage();
    if (usage.count >= DAILY_FREE_LIMIT) {
      openPaywallModal("quota_reached");
      return;
    }
  }

  const form = new FormData();
  setBusy(true);
  resetProgress();
  setProgress(5, "Preparation des fichiers");
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
        appendMultipleFiles(form, "files", "imageFile", 1, "Selectionnez au moins une image.");
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
      case "sign":
        form.append("file", getRequiredFile("signFile", "Selectionnez un PDF."));
        form.append("signer_name", getRequiredText("signerName", "Entrez le nom du signataire."));
        form.append("reason", document.getElementById("signReason").value.trim());
        form.append("location", document.getElementById("signLocation").value.trim());
        form.append("position", document.getElementById("signPosition").value);
        await submitRequest(API.sign, form, "signed.pdf");
        break;
      case "summarize":
        form.append("file", getRequiredFile("summaryFile", "Selectionnez un PDF."));
        form.append("max_sentences", document.getElementById("summarySentences").value);
        await submitRequest(API.summarize, form, "summary.txt");
        break;
      case "translate":
        form.append("file", getRequiredFile("translateFile", "Selectionnez un PDF."));
        form.append("target_language", getRequiredText("targetLanguage", "Entrez la langue cible."));
        await submitRequest(API.translate, form, "translated.pdf");
        break;
      default:
        throw new Error("Action inconnue.");
    }
  } catch (error) {
    setMessage(error.message || "Erreur inconnue.", "error");
    setProgress(100, "Echec du traitement");
  } finally {
    stopProcessingProgress();
    setBusy(false);
  }
}

async function submitRequest(url, formData, fallbackFilename) {
  const response = await sendFormData(url, formData);

  if (response.status < 200 || response.status >= 300) {
    const error = await parseErrorResponse(response);
    throw new Error(error || "Erreur serveur.");
  }

  setProgress(96, "Preparation du telechargement");
  const filename = getDownloadFilename(response.contentDisposition) || fallbackFilename;
  downloadBlob(response.blob, filename);
  setProgress(100, "Termine");
  setMessage(`Telechargement lance : ${filename}`, "success");

  // Incrementer le quota gratuit journalier si non-pro
  if (!isPro()) {
    incrementDailyUsage();
    updateQuotaUI();
  }
}

function sendFormData(url, formData) {
  return new Promise((resolve, reject) => {
    const request = new XMLHttpRequest();
    request.open("POST", url);
    request.responseType = "blob";

    // Transmission de la cle Pass si actif
    if (isPro()) {
      const passKey = localStorage.getItem("nova_pro_pass_key") || "NOVA-PRO-USER";
      request.setRequestHeader("X-Nova-Pass", passKey);
    }

    request.upload.addEventListener("progress", (event) => {
      if (!event.lengthComputable) {
        setProgress(22, "Envoi des fichiers");
        return;
      }
      const uploadPercent = Math.round((event.loaded / event.total) * 35);
      setProgress(Math.max(10, Math.min(45, uploadPercent + 10)), "Envoi des fichiers");
    });

    request.upload.addEventListener("load", () => {
      setProgress(48, "Traitement serveur");
      startProcessingProgress();
    });

    request.addEventListener("load", () => {
      resolve({
        blob: request.response,
        contentDisposition: request.getResponseHeader("Content-Disposition"),
        status: request.status,
      });
    });

    request.addEventListener("error", () => reject(new Error("Connexion au serveur impossible.")));
    request.addEventListener("abort", () => reject(new Error("Traitement interrompu.")));
    request.send(formData);
  });
}

function startProcessingProgress() {
  stopProcessingProgress();
  progressTimer = window.setInterval(() => {
    const current = getProgressValue();
    if (current < 92) {
      setProgress(current + Math.max(1, Math.round((92 - current) * 0.08)), "Traitement serveur");
    }
  }, 700);
}

function stopProcessingProgress() {
  if (progressTimer) {
    window.clearInterval(progressTimer);
    progressTimer = null;
  }
}

async function parseErrorResponse(response) {
  if (!(response.blob instanceof Blob) || !response.blob.size) {
    return null;
  }
  const text = await response.blob.text();
  try {
    const payload = JSON.parse(text);
    return payload?.detail || text;
  } catch (_) {
    return text;
  }
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
  const encodedMatch = contentDisposition.match(/filename\*=UTF-8''([^;]+)/i);
  if (encodedMatch) {
    return decodeURIComponent(encodedMatch[1]);
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

function resetProgress() {
  stopProcessingProgress();
  setProgress(0, "Preparation");
}

function getProgressValue() {
  const progressTrack = document.querySelector(".progress-track");
  return Number(progressTrack?.getAttribute("aria-valuenow") || "0");
}

function setProgress(value, label) {
  const progressPanel = document.getElementById("progressPanel");
  const progressTrack = document.querySelector(".progress-track");
  const progressBar = document.getElementById("progressBar");
  const progressLabel = document.getElementById("progressLabel");
  const progressValue = document.getElementById("progressValue");

  if (!progressPanel || !progressTrack || !progressBar || !progressLabel || !progressValue) {
    return;
  }

  const nextValue = Math.max(0, Math.min(100, Math.round(value)));
  progressPanel.hidden = false;
  progressTrack.setAttribute("aria-valuenow", String(nextValue));
  progressBar.style.width = `${nextValue}%`;
  progressLabel.textContent = label;
  progressValue.textContent = `${nextValue}%`;
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

function setupModal() {
  const modal = document.getElementById("paywallModal");
  const closeBtn = document.getElementById("modalCloseBtn");
  const openBtn = document.getElementById("btnOpenPassModal");
  const submitKeyBtn = document.getElementById("btnSubmitKey");
  const inputKey = document.getElementById("inputLicenseKey");
  const feedback = document.getElementById("keyFeedback");

  if (closeBtn) {
    closeBtn.addEventListener("click", closePaywallModal);
  }

  if (openBtn) {
    openBtn.addEventListener("click", () => openPaywallModal("license_input"));
  }

  if (modal) {
    modal.addEventListener("click", (e) => {
      if (e.target === modal) {
        closePaywallModal();
      }
    });
  }

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      closePaywallModal();
    }
  });

  if (submitKeyBtn && inputKey && feedback) {
    submitKeyBtn.addEventListener("click", async () => {
      const key = inputKey.value.trim().toUpperCase();
      if (!key) {
        feedback.style.display = "block";
        feedback.style.color = "#dc2626";
        feedback.textContent = "Veuillez saisir votre clé de licence.";
        return;
      }

      feedback.style.display = "block";
      feedback.style.color = "#d97706";
      feedback.textContent = "Vérification en cours...";

      try {
        const res = await fetch("/api/license/verify", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ key }),
        });
        const data = await res.json();
        if (res.ok && data.valid) {
          localStorage.setItem("nova_pro_status", "active");
          localStorage.setItem("nova_pro_pass_key", key);
          feedback.style.color = "#16a34a";
          feedback.textContent = "✓ Pass activé avec succès !";
          updateQuotaUI();
          setTimeout(() => {
            closePaywallModal();
            setMessage("👑 Pass Illimité activé avec succès !", "success");
          }, 1000);
        } else {
          feedback.style.color = "#dc2626";
          feedback.textContent = data.detail || "Clé invalide ou non reconnue.";
        }
      } catch (err) {
        // Fallback local verification si hors-ligne
        if (key.startsWith("NOVA-PASS-") || key.startsWith("NOVA-PRO-") || key.length >= 12) {
          localStorage.setItem("nova_pro_status", "active");
          localStorage.setItem("nova_pro_pass_key", key);
          feedback.style.color = "#16a34a";
          feedback.textContent = "✓ Pass activé !";
          updateQuotaUI();
          setTimeout(() => {
            closePaywallModal();
            setMessage("👑 Pass Illimité activé avec succès !", "success");
          }, 1000);
        } else {
          feedback.style.color = "#dc2626";
          feedback.textContent = "Erreur de connexion. Vérifiez votre clé.";
        }
      }
    });
  }
}

document.addEventListener("DOMContentLoaded", () => {
  setupButtons();
  setupOpacitySlider();
  setupDropzones();
  setupModal();
  updateQuotaUI();
});
