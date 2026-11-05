// app.js — Frontend do Formatador ABNT
document.addEventListener("DOMContentLoaded", () => {
  const form = document.getElementById("format-form");
  const dropZone = document.getElementById("drop-zone");
  const fileInput = document.getElementById("file-input");
  const dropTitle = document.getElementById("drop-title");
  const dropSubtitle = document.getElementById("drop-subtitle");
  const selectedFilePill = document.getElementById("selected-file-pill");
  const selectedFileName = document.getElementById("selected-file-name");
  const removeFileBtn = document.getElementById("remove-file-btn");
  const submitBtn = document.getElementById("submit-btn");
  const submitBtnText = document.getElementById("submit-btn-text");
  const submitSpinner = document.getElementById("submit-spinner");

  const formSection = document.getElementById("form-section");
  const resultSection = document.getElementById("result-section");
  const downloadDocxBtn = document.getElementById("download-docx-btn");
  const downloadPdfBtn = document.getElementById("download-pdf-btn");
  const resetBtn = document.getElementById("reset-btn");

  const statsGrid = document.getElementById("stats-grid");
  const changesList = document.getElementById("changes-list");
  const warningsContainer = document.getElementById("warnings-container");
  const warningsList = document.getElementById("warnings-list");

  const errorBanner = document.getElementById("error-banner");
  const errorMessage = document.getElementById("error-message");
  const closeErrorBtn = document.getElementById("close-error-btn");
  const workTypeSelect = document.getElementById("work-type");
  const workTypeHint = document.getElementById("work-type-hint");

  let currentFile = null;
  let currentJobId = null;

  // Dicionário de dicas de tipo de trabalho
  const workTypeHints = {
    tcc: "Seções primárias iniciam em nova página; numeração exibida a partir da Introdução.",
    artigo: "Seções primárias seguem no fluxo contínuo sem quebra de página forçada.",
    relatorio: "Seções primárias iniciam em nova página; estrutura técnica padronizada."
  };

  workTypeSelect.addEventListener("change", () => {
    workTypeHint.textContent = workTypeHints[workTypeSelect.value] || "";
  });

  // Consulta capacidades do servidor (ex.: se LibreOffice está instalado)
  fetch("/api/capabilities")
    .then(res => res.json())
    .then(data => {
      if (!data.pdf_available) {
        downloadPdfBtn.classList.add("hidden");
      }
    })
    .catch(() => {});

  function showError(msg) {
    errorMessage.textContent = msg;
    errorBanner.classList.remove("hidden");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function hideError() {
    errorBanner.classList.add("hidden");
  }

  closeErrorBtn.addEventListener("click", hideError);

  function setFile(file) {
    hideError();
    if (!file) return;

    if (!file.name.toLowerCase().endsWith(".docx")) {
      showError("Por favor, selecione apenas arquivos do Microsoft Word no formato .docx.");
      return;
    }

    const maxBytes = 20 * 1024 * 1024;
    if (file.size > maxBytes) {
      showError("O arquivo selecionado tem mais de 20 MB. O limite máximo é de 20 MB.");
      return;
    }

    currentFile = file;
    selectedFileName.textContent = `${file.name} (${(file.size / (1024 * 1024)).toFixed(2)} MB)`;
    selectedFilePill.classList.remove("hidden");
    dropTitle.classList.add("hidden");
    dropSubtitle.classList.add("hidden");
    submitBtn.removeAttribute("disabled");
  }

  function clearFile() {
    currentFile = null;
    fileInput.value = "";
    selectedFilePill.classList.add("hidden");
    dropTitle.classList.remove("hidden");
    dropSubtitle.classList.remove("hidden");
    submitBtn.setAttribute("disabled", "true");
  }

  removeFileBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    clearFile();
  });

  // Drag and Drop
  dropZone.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", (e) => {
    if (e.target.files && e.target.files[0]) {
      setFile(e.target.files[0]);
    }
  });

  ["dragenter", "dragover"].forEach(evtName => {
    dropZone.addEventListener(evtName, (e) => {
      e.preventDefault();
      dropZone.classList.add("drag-over");
    });
  });

  ["dragleave", "drop"].forEach(evtName => {
    dropZone.addEventListener(evtName, (e) => {
      e.preventDefault();
      dropZone.classList.remove("drag-over");
    });
  });

  dropZone.addEventListener("drop", (e) => {
    if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]) {
      setFile(e.dataTransfer.files[0]);
    }
  });

  // Envio do Formulário
  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    hideError();

    if (!currentFile) {
      showError("Por favor, selecione ou arraste um arquivo .docx antes de formatar.");
      return;
    }

    submitBtn.setAttribute("disabled", "true");
    submitBtnText.textContent = "Formatando segundo a ABNT...";
    submitSpinner.classList.remove("hidden");

    const formData = new FormData(form);
    formData.set("file", currentFile);

    try {
      const response = await fetch("/api/format", {
        method: "POST",
        body: formData,
      });

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.detail || "Erro ao formatar o documento.");
      }

      currentJobId = data.job_id;
      renderResult(data);
    } catch (err) {
      showError(err.message || "Ocorreu um erro inesperado ao processar seu documento.");
    } finally {
      submitBtn.removeAttribute("disabled");
      submitBtnText.textContent = "Formatar documento";
      submitSpinner.classList.add("hidden");
    }
  });

  function renderResult(data) {
    formSection.classList.add("hidden");
    resultSection.classList.remove("hidden");

    downloadDocxBtn.href = data.download_docx_url;

    if (data.download_pdf_url) {
      downloadPdfBtn.href = data.download_pdf_url;
      downloadPdfBtn.classList.remove("hidden");
    } else {
      downloadPdfBtn.classList.add("hidden");
    }

    const rep = data.report;

    // Estatísticas rápidas
    statsGrid.innerHTML = `
      <div class="stat-card">
        <div class="stat-value">${rep.paragraphs_total}</div>
        <div class="stat-label">Parágrafos analisados</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">${rep.paragraphs_spacing_adjusted}</div>
        <div class="stat-label">Espaçamentos ajustados</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">${rep.paragraphs_indent_adjusted}</div>
        <div class="stat-label">Recuos de 1ª linha aplicados</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">${rep.long_quotes_formatted}</div>
        <div class="stat-label">Citações longas (4 cm)</div>
      </div>
      <div class="stat-card">
        <div class="stat-value">${rep.references_formatted}</div>
        <div class="stat-label">Itens de referências</div>
      </div>
    `;

    // Lista de alterações Antes e Depois
    changesList.innerHTML = "";
    (rep.changes_summary || []).forEach(change => {
      const li = document.createElement("li");
      li.textContent = change;
      changesList.appendChild(li);
    });

    // Avisos de revisão manual
    if (rep.warnings && rep.warnings.length > 0) {
      warningsContainer.classList.remove("hidden");
      warningsList.innerHTML = "";
      rep.warnings.forEach(w => {
        const item = document.createElement("div");
        item.className = "warning-item";
        item.innerHTML = `
          <strong>${w.category}</strong>: ${w.reason}
          <span class="warning-excerpt">"${w.excerpt}"</span>
        `;
        warningsList.appendChild(item);
      });
    } else {
      warningsContainer.classList.add("hidden");
    }

    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  // Reset para formatar outro documento
  resetBtn.addEventListener("click", () => {
    if (currentJobId) {
      fetch(`/api/jobs/${currentJobId}`, { method: "DELETE" }).catch(() => {});
      currentJobId = null;
    }
    clearFile();
    resultSection.classList.add("hidden");
    formSection.classList.remove("hidden");
    window.scrollTo({ top: 0, behavior: "smooth" });
  });
});

