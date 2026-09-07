(function () {
  const config = window.CHITU_CONFIG || {};
  const service = config.service || {};
  const state = {
    chatFiles: [],
    baselineFile: null,
    taskId: null,
    pollTimer: null,
    submitting: false
  };

  const el = {
    connectionStatus: document.querySelector("#connectionStatus"),
    form: document.querySelector("#analysisForm"),
    productName: document.querySelector("#productName"),
    analysisPeriod: document.querySelector("#analysisPeriod"),
    chatFiles: document.querySelector("#chatFiles"),
    baselinePanel: document.querySelector("#baselinePanel"),
    baselineText: document.querySelector("#baselineText"),
    baselineDays: document.querySelector("#baselineDays"),
    baselinePeriod: document.querySelector("#baselinePeriod"),
    baselineTextBlock: document.querySelector("#baselineTextBlock"),
    baselineFileBlock: document.querySelector("#baselineFileBlock"),
    baselineFile: document.querySelector("#baselineFile"),
    fileList: document.querySelector("#fileList"),
    formError: document.querySelector("#formError"),
    startAnalysis: document.querySelector("#startAnalysis"),
    taskPanel: document.querySelector("#taskPanel"),
    taskState: document.querySelector("#taskState"),
    taskMessage: document.querySelector("#taskMessage"),
    taskProgressText: document.querySelector("#taskProgressText"),
    progressBar: document.querySelector("#progressBar"),
    chunkProgress: document.querySelector("#chunkProgress"),
    taskError: document.querySelector("#taskError"),
    technicalDetails: document.querySelector("#technicalDetails"),
    technicalError: document.querySelector("#technicalError"),
    taskActions: document.querySelector("#taskActions"),
    retryTask: document.querySelector("#retryTask"),
    resultPanel: document.querySelector("#resultPanel"),
    resultSummary: document.querySelector("#resultSummary"),
    downloadExcel: document.querySelector("#downloadExcel"),
    downloadMarkdown: document.querySelector("#downloadMarkdown"),
    downloadManifest: document.querySelector("#downloadManifest")
  };

  function mode() {
    return document.querySelector('input[name="analysisMode"]:checked').value;
  }

  function baselineMode() {
    return document.querySelector('input[name="baselineMode"]:checked')?.value || "text";
  }

  function formatSize(bytes) {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
  }

  function setHidden(node, hidden) {
    node.classList.toggle("hidden", hidden);
  }

  function renderFiles() {
    el.fileList.innerHTML = "";
    const items = state.chatFiles.map((file) => ({ file, role: "本次聊天" }));
    if (state.baselineFile && mode() === "baseline_compare" && baselineMode() === "file") {
      items.push({ file: state.baselineFile, role: "上次基准" });
    }
    for (const item of items) {
      const row = document.createElement("div");
      row.className = "file-row";
      row.innerHTML = `<span class="file-role">${item.role}</span><span class="file-name"></span><span class="file-size">${formatSize(item.file.size)}</span>`;
      row.querySelector(".file-name").textContent = item.file.name;
      el.fileList.appendChild(row);
    }
  }

  function showFormError(message) {
    el.formError.textContent = message || "";
    setHidden(el.formError, !message);
  }

  function validate() {
    if (!state.chatFiles.length) return "请至少选择一份本次聊天记录。";
    if (mode() === "baseline_compare" && baselineMode() === "text" && !el.baselineText.value.trim()) {
      return "请在「粘贴上期词条」里填入上期内容，或切换为上传基准 Excel。";
    }
    if (mode() === "baseline_compare" && baselineMode() === "file" && !state.baselineFile) {
      return "对比分析需要上传上次基准报告 Excel。";
    }
    return "";
  }

  function buildPayload() {
    const analysisMode = mode();
    const product = el.productName.value.trim();
    const period = el.analysisPeriod.value.trim();
    const messageParts = [];
    if (product) messageParts.push(`产品：${product}`);
    if (period) messageParts.push(`分析周期：${period}`);
    messageParts.push(analysisMode === "baseline_compare" ? "与上次基准对比" : "没有历史基准，建立本期分析");
    const files = state.chatFiles.map((file) => ({
      name: file.name,
      size: file.size,
      type: file.type,
      role: "chat_record"
    }));
    const useBaselineFile = analysisMode === "baseline_compare" && baselineMode() === "file";
    if (useBaselineFile && state.baselineFile) {
      files.push({
        name: state.baselineFile.name,
        size: state.baselineFile.size,
        type: state.baselineFile.type,
        role: "baseline"
      });
    }
    const payload = {
      user_message: messageParts.join("；") + "。",
      analysis_type: analysisMode,
      input_mode: analysisMode === "baseline_compare" ? "baseline_and_chat" : "chat_only",
      files
    };
    if (analysisMode === "baseline_compare" && baselineMode() === "text") {
      payload.baseline_text = el.baselineText.value.trim();
      const days = Number(el.baselineDays.value);
      if (Number.isFinite(days) && days > 0) payload.baseline_days = Math.round(days);
      const baselinePeriod = el.baselinePeriod.value.trim();
      if (baselinePeriod) payload.baseline_period = baselinePeriod;
    }
    return payload;
  }

  function setSubmitting(value) {
    state.submitting = value;
    el.startAnalysis.disabled = value;
    el.startAnalysis.textContent = value ? "正在提交" : "开始分析";
  }

  function stateLabel(status) {
    return {
      queued: "排队中",
      running: "分析中",
      completed: "已完成",
      failed: "未完成"
    }[status] || status;
  }

  function renderTask(job) {
    setHidden(el.taskPanel, false);
    const progress = Math.max(0, Math.min(100, Number(job.progress || 0)));
    el.taskState.textContent = stateLabel(job.status);
    el.taskState.dataset.status = job.status;
    el.taskMessage.textContent = job.message || "正在处理";
    el.taskProgressText.textContent = `${progress}%`;
    el.progressBar.style.width = `${progress}%`;
    el.chunkProgress.textContent = job.total_chunks
      ? `语义分析批次：${job.completed_chunks || 0}/${job.total_chunks}`
      : "";

    const failed = job.status === "failed";
    setHidden(el.taskError, !failed);
    setHidden(el.taskActions, !failed);
    el.taskError.textContent = failed ? job.error?.message || "本次分析没有完成。" : "";
    const detail = job.error?.detail || "";
    setHidden(el.technicalDetails, !detail);
    el.technicalError.textContent = detail;
  }

  function applyDownload(anchor, dataUrl, serverUrl, filename) {
    if (dataUrl) {
      anchor.href = dataUrl;
      if (filename) anchor.setAttribute("download", filename);
      return;
    }
    anchor.removeAttribute("download");
    anchor.href = serverUrl || "#";
  }

  function renderResult(result) {
    const files = result?.files || {};
    const summary = Array.isArray(result?.summary) ? result.summary : [];
    el.resultSummary.textContent = summary[0] || "可以下载查看完整结果。";
    applyDownload(el.downloadExcel, files.excel_data_url, files.excel_download_url, files.excel_filename);
    applyDownload(el.downloadMarkdown, files.markdown_data_url, files.markdown_download_url, files.markdown_filename);
    applyDownload(el.downloadManifest, files.manifest_data_url, files.manifest_download_url, files.manifest_filename);
    setHidden(el.resultPanel, false);
  }

  async function getJson(url, options) {
    const response = await fetch(url, options);
    const body = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(body.message || `服务返回 ${response.status}`);
    return body;
  }

  async function pollTask() {
    if (!state.taskId) return;
    try {
      const job = await getJson(`/api/tasks/${encodeURIComponent(state.taskId)}`);
      renderTask(job);
      if (job.status === "completed") {
        clearInterval(state.pollTimer);
        state.pollTimer = null;
        localStorage.removeItem("chitu_active_task");
        renderResult(job.result);
      } else if (job.status === "failed") {
        clearInterval(state.pollTimer);
        state.pollTimer = null;
      }
    } catch (error) {
      el.taskMessage.textContent = `暂时无法读取任务状态：${error.message}`;
    }
  }

  function startPolling(taskId) {
    state.taskId = taskId;
    localStorage.setItem("chitu_active_task", taskId);
    clearInterval(state.pollTimer);
    void pollTask();
    state.pollTimer = setInterval(pollTask, Number(service.pollIntervalMs || 1500));
  }

  async function submitAnalysis(event) {
    event.preventDefault();
    if (state.submitting) return;
    const validationError = validate();
    showFormError(validationError);
    if (validationError) return;

    setSubmitting(true);
    setHidden(el.resultPanel, true);
    const formData = new FormData();
    formData.append("payload", JSON.stringify(buildPayload()));
    for (const file of state.chatFiles) formData.append("files", file);
    if (state.baselineFile) formData.append("files", state.baselineFile);

    try {
      const created = await getJson(service.tasksUrl || "/api/tasks", {
        method: "POST",
        body: formData
      });
      renderTask({ status: "queued", progress: 0, message: "任务已创建" });
      startPolling(created.task_id);
    } catch (submitError) {
      showFormError(`提交失败：${submitError.message}`);
    } finally {
      setSubmitting(false);
    }
  }

  async function retryTask() {
    if (!state.taskId) return;
    el.retryTask.disabled = true;
    try {
      await getJson(`/api/tasks/${encodeURIComponent(state.taskId)}/retry`, { method: "POST" });
      setHidden(el.taskActions, true);
      startPolling(state.taskId);
    } catch (error) {
      el.taskError.textContent = `重试失败：${error.message}`;
    } finally {
      el.retryTask.disabled = false;
    }
  }

  async function checkService() {
    try {
      await getJson(service.healthUrl || "/api/health");
      el.connectionStatus.textContent = "分析服务正常";
      el.connectionStatus.dataset.status = "ready";
    } catch {
      el.connectionStatus.textContent = "分析服务不可用";
      el.connectionStatus.dataset.status = "failed";
    }
  }

  function bindEvents() {
    document.querySelectorAll('input[name="analysisMode"]').forEach((input) => {
      input.addEventListener("change", () => {
        setHidden(el.baselinePanel, mode() !== "baseline_compare");
        renderFiles();
        showFormError("");
      });
    });
    document.querySelectorAll('input[name="baselineMode"]').forEach((input) => {
      input.addEventListener("change", () => {
        const textMode = baselineMode() === "text";
        setHidden(el.baselineTextBlock, !textMode);
        setHidden(el.baselineFileBlock, textMode);
        renderFiles();
        showFormError("");
      });
    });
    el.chatFiles.addEventListener("change", (event) => {
      state.chatFiles = Array.from(event.target.files || []);
      renderFiles();
      showFormError("");
    });
    el.baselineFile.addEventListener("change", (event) => {
      state.baselineFile = event.target.files?.[0] || null;
      renderFiles();
      showFormError("");
    });
    el.baselineText.addEventListener("input", () => showFormError(""));
    el.form.addEventListener("submit", submitAnalysis);
    el.retryTask.addEventListener("click", retryTask);
  }

  function boot() {
    bindEvents();
    void checkService();
    const activeTask = localStorage.getItem("chitu_active_task");
    if (activeTask) startPolling(activeTask);
  }

  boot();
})();
