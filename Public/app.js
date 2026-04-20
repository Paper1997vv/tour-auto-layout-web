const state = {
  config: null,
  jobId: null,
  pollHandle: null,
};

const loginPanel = document.getElementById("login-panel");
const composerPanel = document.getElementById("composer-panel");
const jobPanel = document.getElementById("job-panel");
const loginForm = document.getElementById("login-form");
const jobForm = document.getElementById("job-form");
const passwordInput = document.getElementById("password-input");
const loginError = document.getElementById("login-error");
const submitError = document.getElementById("submit-error");
const templateInput = document.getElementById("template-input");
const documentsInput = document.getElementById("documents-input");
const templateName = document.getElementById("template-name");
const documentsSummary = document.getElementById("documents-summary");
const documentsList = document.getElementById("documents-list");
const submitButton = document.getElementById("submit-button");
const limitsText = document.getElementById("limits-text");
const progressFill = document.getElementById("progress-fill");
const jobStatusText = document.getElementById("job-status-text");
const jobMeta = document.getElementById("job-meta");
const results = document.getElementById("results");
const batchDownload = document.getElementById("batch-download");

templateInput.addEventListener("change", () => {
  templateName.textContent = templateInput.files?.[0]?.name || "未选择模板图";
});

documentsInput.addEventListener("change", () => {
  renderSelectedDocuments();
});

loginForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  loginError.classList.add("hidden");

  const response = await fetch("/api/auth/login", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ password: passwordInput.value }),
  });

  if (!response.ok) {
    loginError.textContent = "密码不正确，请重试。";
    loginError.classList.remove("hidden");
    return;
  }

  passwordInput.value = "";
  await bootstrap();
});

jobForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  submitError.classList.add("hidden");

  const templateFile = templateInput.files?.[0];
  const documentFiles = Array.from(documentsInput.files || []);

  if (!templateFile || documentFiles.length === 0) {
    submitError.textContent = "请先选择模板图和至少一份文档。";
    submitError.classList.remove("hidden");
    return;
  }

  submitButton.disabled = true;
  resetCurrentJobView();
  jobPanel.classList.remove("hidden");
  jobStatusText.textContent = "状态：正在提交，完成 0%";

  try {
    const formData = new FormData();
    formData.append("templateImage", templateFile);
    documentFiles.forEach((file) => formData.append("documents[]", file));

    const response = await fetch("/api/jobs", {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      const text = await response.text();
      throw new Error(text || "任务创建失败");
    }

    const payload = await response.json();
    state.jobId = payload.jobId;
    await pollJob();
    startPolling();
  } catch (error) {
    jobPanel.classList.add("hidden");
    submitError.textContent = error.message || "任务创建失败。";
    submitError.classList.remove("hidden");
  } finally {
    submitButton.disabled = false;
  }
});

async function bootstrap() {
  const response = await fetch("/api/config");
  state.config = await response.json();

  limitsText.textContent = `单次上传总大小建议不超过 ${state.config.maxUploadSizeMB}MB`;

  if (state.config.requiresPassword && !state.config.authenticated) {
    loginPanel.classList.remove("hidden");
    composerPanel.classList.add("hidden");
    jobPanel.classList.add("hidden");
    return;
  }

  loginPanel.classList.add("hidden");
  composerPanel.classList.remove("hidden");
  jobPanel.classList.add("hidden");
}

function renderSelectedDocuments() {
  documentsList.innerHTML = "";
  const files = Array.from(documentsInput.files || []);

  if (files.length === 0) {
    documentsSummary.textContent = "未选择文档";
    const item = document.createElement("li");
    item.textContent = "可一次选择多份 .doc / .docx 文档";
    item.className = "file-item empty";
    documentsList.appendChild(item);
    return;
  }

  documentsSummary.textContent = `已选择 ${files.length} 份文档`;

  files.forEach((file) => {
    const item = document.createElement("li");
    item.className = "file-item";
    item.textContent = file.name;
    documentsList.appendChild(item);
  });
}

function resetCurrentJobView() {
  stopPolling();
  progressFill.style.width = "0%";
  batchDownload.classList.add("hidden");
  batchDownload.removeAttribute("href");
  jobMeta.textContent = "正在创建任务…";
  results.innerHTML = "";
}

async function pollJob() {
  if (!state.jobId) return;

  const response = await fetch(`/api/jobs/${state.jobId}`);
  if (!response.ok) {
    jobStatusText.textContent = "任务状态读取失败。";
    return;
  }

  const payload = await response.json();
  renderJob(payload);

  if (payload.status === "completed") {
    stopPolling();
  }
}

function startPolling() {
  stopPolling();
  state.pollHandle = window.setInterval(pollJob, 1500);
}

function stopPolling() {
  if (state.pollHandle) {
    window.clearInterval(state.pollHandle);
    state.pollHandle = null;
  }
}

function renderJob(payload) {
  progressFill.style.width = `${Math.round(payload.progress * 100)}%`;
  jobStatusText.textContent = `状态：${translateOverallStatus(payload.status)}，完成 ${Math.round(payload.progress * 100)}%`;
  jobMeta.textContent = `任务 ID：${payload.jobId} · 创建时间：${payload.createdAt}`;

  if (payload.status === "completed") {
    batchDownload.href = `/api/jobs/${payload.jobId}/download.zip`;
    batchDownload.classList.remove("hidden");
  } else {
    batchDownload.classList.add("hidden");
  }

  results.innerHTML = "";
  payload.files.forEach((file) => {
    const card = document.createElement("article");
    card.className = "result-card";
    card.innerHTML = `
      <div class="result-top">
        <h3>${escapeHtml(file.filename)}</h3>
        <span class="badge ${file.status}">${translateStatus(file.status)}</span>
      </div>
      <p class="card-meta">${file.outputFilename ? `输出：${escapeHtml(file.outputFilename)}` : "等待生成结果"}</p>
      ${file.durationSeconds ? `<p class="card-meta">耗时：${file.durationSeconds.toFixed(2)} 秒</p>` : ""}
      ${file.errorMessage ? `<p class="message error">${escapeHtml(file.errorMessage)}</p>` : ""}
      ${file.warnings?.length ? `<p class="message">${file.warnings.map(escapeHtml).join("<br>")}</p>` : ""}
      ${file.downloadURL ? `<div class="card-actions"><a class="secondary-link" href="${file.downloadURL}" target="_blank" rel="noreferrer">下载文件</a></div>` : ""}
    `;
    results.appendChild(card);
  });
}

function translateStatus(status) {
  const map = {
    queued: "排队中",
    processing: "处理中",
    success: "成功",
    warning: "成功但有警告",
    failure: "失败",
  };
  return map[status] || status;
}

function translateOverallStatus(status) {
  const map = {
    queued: "等待处理",
    processing: "正在处理",
    completed: "处理完成",
  };
  return map[status] || status;
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

renderSelectedDocuments();
bootstrap();
