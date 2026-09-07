(function () {
  const config = window.CHITU_CONFIG || {};
  const service = config.service || {};
  const state = {
    chatFiles: [],
    baselineFile: null,
    taskId: null,
    pollTimer: null,
    submitting: false,
    cancelling: false
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
    cancelArea: document.querySelector("#cancelArea"),
    cancelTask: document.querySelector("#cancelTask"),
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
    const MAX_UPLOAD_BYTES = 25 * 1024 * 1024;
    const oversized = [...state.chatFiles, ...(state.baselineFile ? [state.baselineFile] : [])]
      .find((f) => f && f.size > MAX_UPLOAD_BYTES);
    if (oversized) {
      return `文件「${oversized.name}」约 ${Math.ceil(oversized.size / 1024 / 1024)}MB，超过单次上传上限（25MB），请删减或拆分聊天记录后重试。`;
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
      failed: "未完成",
      cancelled: "已取消"
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

    const failed = job.status === "failed" || job.status === "cancelled";
    const canRetry = job.status === "failed" || job.status === "cancelled";
    const active = job.status === "queued" || job.status === "running";
    setHidden(el.taskError, !failed);
    setHidden(el.taskActions, !canRetry);
    setHidden(el.cancelArea, !active || Boolean(state.cancelling));
    if (!active) {
      el.cancelTask.disabled = false;
      el.cancelTask.textContent = "取消分析";
    }
    el.taskError.textContent = failed ? job.error?.message || "本次分析没有完成。" : "";
    const detail = job.error?.detail || "";
    setHidden(el.technicalDetails, !detail);
    el.technicalError.textContent = detail;
  }

  async function downloadViaBlob(url, filename) {
    // 跨域签名 URL 的 download 属性会被浏览器忽略，必须 fetch + blob 落内存后触发下载
    const response = await fetch(url);
    if (!response.ok) throw new Error(`下载失败（服务返回 ${response.status}）`);
    const blob = await response.blob();
    const blobUrl = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = blobUrl;
    if (filename) link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(blobUrl), 4000);
  }

  function bindDownload(anchor, dataUrl, objectFileUrl, serverUrl, filename) {
    anchor.onclick = null;
    anchor.removeAttribute("download");
    anchor.href = serverUrl || "#";
    if (dataUrl) {
      anchor.href = dataUrl;
      if (filename) anchor.setAttribute("download", filename);
      return;
    }
    if (!objectFileUrl) return;
    anchor.onclick = (event) => {
      event.preventDefault();
      if (anchor.dataset.downloading === "1") return;
      anchor.dataset.downloading = "1";
      const originalText = anchor.textContent;
      anchor.textContent = "正在准备下载…";
      downloadViaBlob(objectFileUrl, filename)
        .catch((error) => {
          console.error("对象存储下载失败，回退服务端路径：", error);
          window.location.href = serverUrl || objectFileUrl;
        })
        .finally(() => {
          anchor.dataset.downloading = "0";
          anchor.textContent = originalText;
        });
    };
  }

  function renderResult(result) {
    const files = result?.files || {};
    const summary = Array.isArray(result?.summary) ? result.summary : [];
    el.resultSummary.textContent = summary[0] || "可以下载查看完整结果。";
    const taskId = state.taskId;
    bindDownload(el.downloadExcel, files.excel_data_url, `/api/tasks/${taskId}/file/excel`, files.excel_download_url, files.excel_filename);
    bindDownload(el.downloadMarkdown, files.markdown_data_url, `/api/tasks/${taskId}/file/markdown`, files.markdown_download_url, files.markdown_filename);
    bindDownload(el.downloadManifest, files.manifest_data_url, `/api/tasks/${taskId}/file/manifest`, files.manifest_download_url, files.manifest_filename);
    setHidden(el.resultPanel, false);
  }

  async function getJson(url, options) {
    const response = await fetch(url, options);
    const body = await response.json().catch(() => ({}));
    if (!response.ok) {
      if (response.status === 413) {
        throw new Error("文件过大，超过平台上传大小限制，请删减或拆分聊天记录文件后重试");
      }
      if (response.status === 502 || response.status === 503 || response.status === 504) {
        throw new Error("服务正在启动或暂时不可用，请等待几秒后重试");
      }
      throw new Error(body.message || `服务返回 ${response.status}`);
    }
    return body;
  }

  async function pollTask() {
    if (!state.taskId) return;
    try {
      const job = await getJson(`/api/tasks/${encodeURIComponent(state.taskId)}`);
      state.pollMisses = 0;
      renderTask(job);
      if (job.status === "completed") {
        clearInterval(state.pollTimer);
        state.pollTimer = null;
        localStorage.removeItem("chitu_active_task");
        renderResult(job.result);
      } else if (job.status === "failed" || job.status === "cancelled") {
        clearInterval(state.pollTimer);
        state.pollTimer = null;
        localStorage.removeItem("chitu_active_task");
      }
    } catch (error) {
      state.pollMisses = (state.pollMisses || 0) + 1;
      const notFound = typeof error.message === "string" && error.message.includes("找不到");
      if (notFound && state.pollMisses <= 12) {
        el.taskMessage.textContent = "任务状态同步中（服务实例切换），请稍候…";
        return;
      }
      if (notFound) {
        clearInterval(state.pollTimer);
        state.pollTimer = null;
        localStorage.removeItem("chitu_active_task");
        el.taskMessage.textContent = "任务状态已失效（服务实例已切换且分析尚未完成），请重新上传分析。";
        return;
      }
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

  async function cancelTask() {
    if (!state.taskId || state.cancelling) return;
    if (!window.confirm("确定要取消本次分析吗？取消后需要重新提交。")) return;
    state.cancelling = true;
    el.cancelTask.disabled = true;
    el.cancelTask.textContent = "正在取消…";
    try {
      await getJson(`/api/tasks/${encodeURIComponent(state.taskId)}/cancel`, { method: "POST" });
      setHidden(el.cancelArea, true);
      void pollTask();
    } catch (error) {
      el.taskMessage.textContent = `取消失败：${error.message}`;
    } finally {
      state.cancelling = false;
      el.cancelTask.disabled = false;
      el.cancelTask.textContent = "取消分析";
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
    el.cancelTask.addEventListener("click", cancelTask);
  }

  function boot() {
    bindEvents();
    void checkService();
    const activeTask = localStorage.getItem("chitu_active_task");
    if (activeTask) startPolling(activeTask);
  }

  boot();
})();
