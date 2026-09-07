import { createServer } from "node:http";
import { copyFile, readFile, mkdir, readdir, rename, stat, writeFile } from "node:fs/promises";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Readable } from "node:stream";
import { spawn } from "node:child_process";
import { randomUUID } from "node:crypto";
import { tmpdir } from "node:os";

const __filename = fileURLToPath(import.meta.url);
const projectRoot = path.dirname(__filename);
const workspaceRoot = path.dirname(projectRoot);
const webRoot = path.join(projectRoot, "web");

loadEnvFile(path.join(projectRoot, ".env.local"));
loadEnvFile(path.join(projectRoot, ".env"));

// 生产环境（veFaaS 等）代码目录 /opt/bytefaas 为只读，
// 任务快照 / LLM 缓存 / 历史归档等运行时写入必须落到可写目录（/tmp）。
// DEV 环境仍沿用项目内目录，行为不变。
const isProdEnv = process.env.COZE_PROJECT_ENV === "PROD" || process.env.NODE_ENV === "production";
const stateRoot = path.resolve(
  process.env.CHITU_STATE_ROOT ||
    (isProdEnv ? path.join(tmpdir(), "chitu-state") : path.join(projectRoot, "storage"))
);
const historyRoot = path.resolve(
  process.env.CHITU_HISTORY_ROOT ||
    (isProdEnv ? path.join(tmpdir(), "chitu-history") : path.join(workspaceRoot, "赤兔历史分析结果"))
);
const host = process.env.CHITU_HOST || "0.0.0.0";
const port = Number(process.env.DEPLOY_RUN_PORT || process.env.CHITU_PORT || 8787);
const corsOrigin = process.env.CHITU_CORS_ORIGIN || "*";
const jobsRoot = path.join(stateRoot, "jobs");
const llmCacheRoot = path.join(stateRoot, "cache", "llm");
const jobs = new Map();
const jobQueue = [];
const maxActiveJobs = Math.max(1, Number(process.env.CHITU_MAX_ACTIVE_JOBS || 12));
let activeJobs = 0;

const runtimeConfig = {
  llmApiBase: process.env.CHITU_LLM_API_BASE || process.env.OPENAI_BASE_URL || "",
  llmApiKey: process.env.CHITU_LLM_API_KEY || process.env.OPENAI_API_KEY || "",
  analysisModel: process.env.CHITU_ANALYSIS_MODEL || "gpt-5.6-sol"
};

const mimeTypes = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".txt": "text/plain; charset=utf-8",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
};

function sendJson(res, statusCode, body) {
  res.writeHead(statusCode, {
    "content-type": "application/json; charset=utf-8",
    "access-control-allow-origin": corsOrigin,
    "access-control-allow-methods": "GET,POST,OPTIONS",
    "access-control-allow-headers": "content-type"
  });
  res.end(JSON.stringify(body, null, 2));
}

function sendText(res, statusCode, text) {
  res.writeHead(statusCode, {
    "content-type": "text/plain; charset=utf-8",
    "access-control-allow-origin": corsOrigin
  });
  res.end(text);
}

function downloadUrl(filePath) {
  return `/api/download?path=${encodeURIComponent(filePath)}`;
}

function loadEnvFile(filePath) {
  if (!existsSync(filePath)) return;
  const lines = readFileSync(filePath, "utf8").split(/\r?\n/);
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const match = trimmed.match(/^([A-Za-z_][A-Za-z0-9_]*)=(.*)$/);
    if (!match) continue;
    const [, key, rawValue] = match;
    if (process.env[key] !== undefined) continue;
    process.env[key] = rawValue.replace(/^["']|["']$/g, "");
  }
}

function safeSegment(value) {
  const safe = String(value || "未命名任务")
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+/g, "_")
    .slice(0, 120);
  return safe === "." || safe === ".." ? "未命名任务" : safe;
}

function nowParts() {
  const now = new Date();
  const pad = (value) => String(value).padStart(2, "0");
  return {
    year: String(now.getFullYear()),
    month: pad(now.getMonth() + 1),
    day: pad(now.getDate()),
    text: `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`
  };
}

function inferProductName(message, files) {
  const marked = message.match(/【([^】]{2,30})】/);
  if (marked?.[1]) return marked[1];
  if (/丹橘.*(?:三款.*卷发|卷发.*三款)|三款.*丹橘.*卷发/.test(message)) return "丹橘卷发棒三款产品";
  if (/三款.*卷发/.test(message)) return "三款卷发棒产品";
  const product = message.match(/(?:产品|商品|这是|分析|关于)(?:名|是|：|:)?\s*([A-Za-z0-9\u4e00-\u9fa5_-]{2,30})/);
  if (product?.[1]) return product[1].replace(/[，。,.\s].*$/, "");
  const chatFiles = files.filter((file) => /\.(?:log|txt|csv)$/i.test(file.name));
  if (chatFiles.length > 1) {
    const stems = chatFiles.map((file) => path.parse(file.name).name);
    let common = stems[0];
    for (const stem of stems.slice(1)) {
      while (common && !stem.startsWith(common)) common = common.slice(0, -1);
    }
    common = common.replace(/[-_\s]+$/, "");
    const countName = ({ 2: "双", 3: "三", 4: "四", 5: "五", 6: "六" })[chatFiles.length] || String(chatFiles.length);
    return `${common.length >= 2 ? common : "多商品"}${countName}SKU`.slice(0, 30);
  }
  if (chatFiles.length) return path.parse(chatFiles[0].name).name.slice(0, 30);
  if (files.length) return path.parse(files[0].name).name.slice(0, 30);
  return "未命名产品";
}

function inferFileRole(file) {
  const name = String(file?.name || "");
  const lowerName = name.toLowerCase();
  const ext = path.extname(name).toLowerCase().replace(".", "");
  const isWorkbook = ext === "xlsx" || ext === "xls";
  const looksLikeBaseline =
    /基准|baseline|上期|上一期|历史/.test(name) ||
    (isWorkbook && /报告|分析结果|管理结论|result_manifest/.test(name));
  const looksLikeChat =
    /聊天|chat|客服|raw|fulltext|记录/.test(lowerName) ||
    ["csv", "log", "txt"].includes(ext);

  if (looksLikeBaseline) return "baseline";
  if (looksLikeChat) return "chat_record";
  return "reference";
}

function roleFromPayload(payload, file) {
  const descriptor = Array.isArray(payload.files)
    ? payload.files.find((item) => item.name === file.name && item.size === file.size)
    : null;
  if (["baseline", "chat_record", "reference"].includes(descriptor?.role)) {
    return descriptor.role;
  }
  return inferFileRole(file);
}

function splitFilesByRole(files) {
  return {
    baselineFiles: files.filter((file) => file.role === "baseline"),
    chatFiles: files.filter((file) => file.role === "chat_record"),
    referenceFiles: files.filter((file) => file.role === "reference")
  };
}

function isCurlingIronTask(message, chatFiles) {
  return Boolean(buildCurlingInputSpec(chatFiles));
}

function isXiaoqipao2449Task(message, chatFiles, baselineFiles) {
  const names = [...chatFiles, ...baselineFiles].map((file) => file.name).join(" ");
  const text = `${message} ${names}`;
  return /2449|小气泡|997158824546/.test(text);
}

function curlingProductFromName(name) {
  if (/五合一/.test(name)) return "五合一卷发棒";
  if (/366/.test(name)) return "366卷发棒";
  if (/856/.test(name)) return "856卷发棒";
  return null;
}

function buildCurlingInputSpec(chatFiles) {
  const byProduct = new Map();
  for (const file of chatFiles) {
    const product = curlingProductFromName(file.name);
    if (product && !byProduct.has(product)) {
      byProduct.set(product, { product, path: file.path });
    }
  }
  const required = ["五合一卷发棒", "366卷发棒", "856卷发棒"];
  if (!required.every((product) => byProduct.has(product))) return null;
  return required.map((product) => byProduct.get(product));
}

function find2449ChatFile(chatFiles) {
  return (
    chatFiles.find((file) => /997158824546|2449|小气泡/.test(file.name)) ||
    chatFiles.find((file) => /\.(csv|txt|log)$/i.test(file.name)) ||
    null
  );
}

function findBaselineWorkbook(baselineFiles) {
  return baselineFiles.find((file) => /\.(xlsx|xls)$/i.test(file.name)) || null;
}

function inferAnalysisType(message, fileGroups, requestedType) {
  if (requestedType === "baseline_compare" && fileGroups.baselineFiles.length && fileGroups.chatFiles.length) {
    return "baseline_compare";
  }
  if (requestedType === "first_baseline") return "first_baseline";
  if (fileGroups.baselineFiles.length && fileGroups.chatFiles.length) return "baseline_compare";
  if (/首次|第一次|没有历史基准|无历史基准|V1|新品/.test(message)) return "first_baseline";
  if (/基准|对比|上一期|上期|历史/.test(message)) {
    return "baseline_compare";
  }
  return "first_baseline";
}

function buildSummary(productName, analysisType, savedFiles, fileGroups) {
  const hasBaseline = analysisType === "baseline_compare";
  return [
    `已收到${productName}的任务说明和 ${savedFiles.length} 个文件，其中本次聊天记录 ${fileGroups.chatFiles.length} 个、上次基准 ${fileGroups.baselineFiles.length} 个，原始输入已归档到《赤兔历史分析结果》。`,
    hasBaseline
      ? "本次按“有历史基准”任务处理：正式分析会优先读取并沿用上期词条、定义和去重口径。"
      : "本次按“首次/普通分析”任务处理：正式分析会从聊天记录中建立或更新需求词条。",
    "模型会读取本次聊天并重新建立/校准需求词条，逐条完成售前售后、风险、明细和同人去重分析。",
    "Excel 只复用正式报告母版的表格格式，不复用母版中的旧分析内容。"
  ];
}

function runPython(scriptPath, args, extraEnv = {}, onProgress = () => {}) {
  return new Promise((resolve, reject) => {
    const child = spawn("python", [scriptPath, ...args], {
      cwd: projectRoot,
      env: {
        ...process.env,
        PYTHONUTF8: "1",
        PYTHONIOENCODING: "utf-8",
        CHITU_LLM_TIMEOUT: "240",
        CHITU_LLM_RETRIES: "3",
        CHITU_LLM_MAX_ATTEMPTS: "0",
        CHITU_LLM_WORKERS: process.env.CHITU_LLM_WORKERS || "12",
        ...extraEnv
      },
      windowsHide: true
    });
    let stdout = "";
    let stderr = "";
    let stdoutBuffer = "";
    child.stdout.on("data", (chunk) => {
      const text = chunk.toString("utf8");
      stdout += text;
      stdoutBuffer += text;
      const lines = stdoutBuffer.split(/\r?\n/);
      stdoutBuffer = lines.pop() || "";
      for (const line of lines) {
        if (!line.startsWith("CHITU_PROGRESS ")) continue;
        try {
          onProgress(JSON.parse(line.slice("CHITU_PROGRESS ".length)));
        } catch {
          // Ignore malformed progress messages; the final process result remains authoritative.
        }
      }
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk.toString("utf8");
    });
    child.on("error", reject);
    child.on("close", (code) => {
      if (code === 0) resolve(stdout);
      else reject(new Error(stderr || `Python exited with code ${code}`));
    });
  });
}

async function newestManifestIn(dir, keyword) {
  if (!existsSync(dir)) return null;
  const entries = await readdir(dir, { withFileTypes: true });
  const manifests = [];
  for (const entry of entries) {
    if (!entry.isDirectory()) continue;
    if (keyword && !entry.name.includes(keyword)) continue;
    const manifestPath = path.join(dir, entry.name, "result_manifest.json");
    if (!existsSync(manifestPath)) continue;
    const info = await stat(manifestPath);
    manifests.push({ manifestPath, mtimeMs: info.mtimeMs });
  }
  manifests.sort((left, right) => right.mtimeMs - left.mtimeMs);
  return manifests[0]?.manifestPath || null;
}

async function copySpecializedOutputs(manifestPath, excelPath, markdownPath, label) {
  const manifest = JSON.parse(await readFile(manifestPath, "utf8"));
  const sourceExcel = manifest.output_files?.excel;
  const sourceMarkdown = manifest.output_files?.markdown;
  if (!sourceExcel || !existsSync(sourceExcel)) {
    throw new Error(`${label}正式分析脚本未生成 Excel`);
  }
  await copyFile(sourceExcel, excelPath);
  if (sourceMarkdown && existsSync(sourceMarkdown)) {
    await copyFile(sourceMarkdown, markdownPath);
  }
  return manifest;
}

async function runCurlingIronCodexReport(chatFiles, baselineFiles, excelPath, markdownPath, onProgress) {
  const inputSpec = buildCurlingInputSpec(chatFiles);
  if (!inputSpec) {
    throw new Error("三款卷发棒正式分析需要同时上传五合一、366、856三份聊天记录。");
  }
  const specializedOutputRoot = path.dirname(excelPath);
  const env = {
    CHITU_CURLING_INPUTS_JSON: JSON.stringify(inputSpec),
    CHITU_OUTPUT_ROOT: specializedOutputRoot,
    CHITU_LLM_CACHE_DIR: llmCacheRoot,
    CHITU_LLM_TIMEOUT: "240",
    CHITU_LLM_RETRIES: "3",
    CHITU_LLM_MAX_ATTEMPTS: "0",
    CHITU_LLM_WORKERS: process.env.CHITU_LLM_WORKERS || "12"
  };
  const baselineFile = findBaselineWorkbook(baselineFiles);
  if (baselineFile) {
    env.CHITU_CURLING_TAXONOMY_XLSX = baselineFile.path;
  }
  await runPython(
    path.join(projectRoot, "scripts", "analyze_curling_irons.py"),
    [],
    env,
    onProgress
  );
  const manifestPath = await newestManifestIn(
    specializedOutputRoot,
    "丹橘三款卷发棒首次基准"
  );
  if (!manifestPath) {
    throw new Error("三款卷发棒正式分析脚本未生成 result_manifest.json");
  }
  return copySpecializedOutputs(manifestPath, excelPath, markdownPath, "三款卷发棒");
}

async function runXiaoqipao2449CodexReport(chatFiles, baselineFiles, excelPath, markdownPath, onProgress) {
  const chatFile = find2449ChatFile(chatFiles);
  if (!chatFile) {
    throw new Error("2449小气泡正式分析需要上传本期聊天记录 CSV/log/txt。");
  }
  const baselineFile = findBaselineWorkbook(baselineFiles);
  const specializedOutputRoot = path.dirname(excelPath);
  const env = {
    CHITU_INPUT_CSV: chatFile.path,
    CHITU_OUTPUT_ROOT: specializedOutputRoot,
    CHITU_LLM_CACHE_DIR: llmCacheRoot
  };
  if (baselineFile) {
    env.CHITU_BASELINE_XLSX = baselineFile.path;
  }
  await runPython(path.join(projectRoot, "scripts", "analyze_2449_new_period.py"), [], env, onProgress);
  const manifestPath = await newestManifestIn(
    specializedOutputRoot,
    "2449小气泡_20260819-0825"
  );
  if (!manifestPath) {
    throw new Error("2449小气泡正式分析脚本未生成 result_manifest.json");
  }
  return copySpecializedOutputs(manifestPath, excelPath, markdownPath, "2449小气泡");
}

async function parseFormData(req) {
  const request = new Request("http://127.0.0.1/api/chitu-analyze", {
    method: req.method,
    headers: req.headers,
    body: Readable.toWeb(req),
    duplex: "half"
  });
  const formData = await request.formData();
  const payload = JSON.parse(String(formData.get("payload") || "{}"));
  const files = formData.getAll("files").filter((item) => typeof item !== "string");
  return { payload, files };
}

async function executeAnalysis(payload, files, onProgress = () => {}) {
  const now = nowParts();
  const taskId = safeSegment(payload.task_id || `${now.year}${now.month}${now.day}_${Date.now()}`);
  const userMessage = String(payload.user_message || "");
  const receivedFiles = files.map((file) => ({
    file,
    name: safeSegment(file.name || "uploaded_file"),
    role: roleFromPayload(payload, file)
  }));
  const receivedGroups = splitFilesByRole(receivedFiles);
  const manualBaselineText = String(payload.baseline_text || "").trim();
  const manualBaseline = manualBaselineText
    ? {
        text: manualBaselineText,
        days: Number(payload.baseline_days) > 0 ? Math.round(Number(payload.baseline_days)) : null,
        period: String(payload.baseline_period || "").trim()
      }
    : null;
  const productName = inferProductName(userMessage, receivedFiles);
  const analysisType = manualBaselineText
    ? "baseline_compare"
    : inferAnalysisType(userMessage, receivedGroups, payload.analysis_type);

  if (!receivedGroups.chatFiles.length) {
    throw new Error("请至少上传本次聊天记录文件。若要做基准对照，请同时上传上次分析基准。");
  }

  const uploadDir = path.join(historyRoot, "01_用户上传", now.year, now.month, taskId);
  const processDir = path.join(historyRoot, "02_过程文件", now.year, now.month, taskId);
  const resultDir = path.join(historyRoot, "03_最终报告", now.year, now.month, taskId);
  await mkdir(uploadDir, { recursive: true });
  await mkdir(processDir, { recursive: true });
  await mkdir(resultDir, { recursive: true });
  onProgress({ stage: "saving_files", progress: 5, message: "正在保存上传文件" });

  const savedFiles = [];
  for (const item of receivedFiles) {
    const file = item.file;
    const filename = item.name;
    const target = path.join(uploadDir, filename);
    const buffer = Buffer.from(await file.arrayBuffer());
    await writeFile(target, buffer);
    savedFiles.push({ name: filename, size: buffer.length, role: item.role, path: target });
  }
  const savedFileGroups = splitFilesByRole(savedFiles);

  await writeFile(path.join(uploadDir, "用户文字说明.md"), userMessage || "用户未填写文字说明。", "utf8");
  await writeFile(path.join(uploadDir, "上传文件清单.json"), JSON.stringify(savedFiles, null, 2), "utf8");
  if (manualBaseline) {
    await writeFile(path.join(uploadDir, "上期基准_手工录入.txt"), manualBaselineText, "utf8");
  }

  const summary = buildSummary(productName, analysisType, savedFiles, savedFileGroups);
  await writeFile(
    path.join(processDir, "口径检查.md"),
    [
      "# 口径检查",
      "",
      `- 任务 ID：${taskId}`,
      `- 产品名：${productName}`,
      `- 分析类型：${analysisType}`,
      `- 输入模式：${analysisType === "baseline_compare" ? "上次基准 + 本次聊天记录" : "仅本次聊天记录"}`,
      `- 模型：${payload.model || runtimeConfig.analysisModel}`,
      `- API Base：${runtimeConfig.llmApiBase || "未配置"}`,
      `- API Key：${runtimeConfig.llmApiKey ? "已配置" : "未配置"}`,
      `- 文件数量：${savedFiles.length}`,
      `- 本次聊天记录：${savedFileGroups.chatFiles.length}`,
      `- 上次基准：${savedFileGroups.baselineFiles.length}${manualBaseline ? "（另含手工录入基准）" : ""}`,
      "- 去重口径：同一买家 + 同一阶段 + 同一需求词条 + 当前周期",
      "- P0/质量/安全：需要二次证据复核",
      "- 状态：已按 Codex 正式报告母版生成"
    ].join("\n"),
    "utf8"
  );

  const safeProduct = safeSegment(productName);
  const markdownPath = path.join(resultDir, `${safeProduct}_${now.year}${now.month}${now.day}_管理结论.md`);
  const excelPath = path.join(resultDir, `${safeProduct}_${now.year}${now.month}${now.day}_客服聊天需求分析报告.xlsx`);
  const manifestPath = path.join(resultDir, "result_manifest.json");

  await writeFile(
    markdownPath,
    [`# ${productName} 客服聊天分析结果`, "", "## 核心结论", "", ...summary.map((item, index) => `${index + 1}. ${item}`)].join("\n"),
    "utf8"
  );

  const reportDataPath = path.join(processDir, "report_data.json");
  const reportData = {
    task_id: taskId,
    product_name: productName,
    analysis_type: analysisType,
    model: payload.model || runtimeConfig.analysisModel,
    llm_api_base: runtimeConfig.llmApiBase,
    llm_api_key_configured: Boolean(runtimeConfig.llmApiKey),
    created_at: now.text,
    user_message: userMessage,
    summary,
    saved_files: savedFiles,
    manual_baseline: manualBaseline,
    baseline_files: savedFileGroups.baselineFiles,
    chat_files: savedFileGroups.chatFiles,
    reference_files: savedFileGroups.referenceFiles
  };
  await writeFile(reportDataPath, JSON.stringify(reportData, null, 2), "utf8");
  onProgress({ stage: "analyzing", progress: 10, message: "已读取文件，准备建立需求词条" });

  let finalSummary = summary;
  let specializedManifest = null;
  let analyzerName = "scripts/generate_report.py";
  let reportLevel = "formal";
  let responseMessage = "AI分析已完成";
  if (isCurlingIronTask(userMessage, savedFileGroups.chatFiles)) {
    specializedManifest = await runCurlingIronCodexReport(
      savedFileGroups.chatFiles,
      savedFileGroups.baselineFiles,
      excelPath,
      markdownPath,
      onProgress
    );
    analyzerName = "scripts/analyze_curling_irons.py";
    finalSummary = Array.isArray(specializedManifest.summary) ? specializedManifest.summary : summary;
  } else {
    await runPython(
      path.join(projectRoot, "scripts", "generate_report.py"),
      [reportDataPath, excelPath],
      { CHITU_LLM_CACHE_DIR: llmCacheRoot },
      onProgress
    );
    const analyzedReportData = JSON.parse(await readFile(reportDataPath, "utf8"));
    finalSummary = Array.isArray(analyzedReportData.summary) ? analyzedReportData.summary : summary;
    reportData.analysis_engine = analyzedReportData.analysis_engine;
    reportData.llm_usage = analyzedReportData.llm_usage;
    reportData.taxonomy = analyzedReportData.taxonomy;
    reportData.products = analyzedReportData.products || [];
    reportData.multi_product_sheets = analyzedReportData.multi_product_sheets || [];
    reportData.supplemental_sheets = analyzedReportData.supplemental_sheets || [];
    await writeFile(
      markdownPath,
      [`# ${productName} 客服聊天 AI 分析`, "", "## 核心结论", "", ...finalSummary.map((item, index) => `${index + 1}. ${item}`)].join("\n"),
      "utf8"
    );
  }
  if (specializedManifest) {
    reportData.analysis_engine = specializedManifest.analysis_engine;
    reportData.llm_usage = specializedManifest.llm_usage;
    reportData.taxonomy = specializedManifest.taxonomy;
    if (specializedManifest.status !== "completed" || specializedManifest.llm_usage?.analysis_complete === false || specializedManifest.llm_usage?.degraded_messages) {
      throw new Error("完整性门禁失败：正式分析仍有未完成消息，已禁止发布 Excel，请保持任务运行并重试。");
    }
  }
  if (reportData.llm_usage && (
    reportData.llm_usage.analysis_complete !== true
    || reportData.llm_usage.expected_messages !== reportData.llm_usage.analyzed_messages
    || reportData.llm_usage.degraded_messages
  )) {
    throw new Error("完整性门禁失败：仍有消息未完成模型分析，已禁止发布 Excel。");
  }
  reportData.summary = finalSummary;
  reportData.analyzer = analyzerName;
  reportData.report_level = reportLevel;
  const finalStatus = reportLevel === "formal" ? "completed" : "draft_completed";
  reportData.status = finalStatus;
  await writeFile(reportDataPath, JSON.stringify(reportData, null, 2), "utf8");
  onProgress({ stage: "publishing", progress: 96, message: "正在发布标准 Excel 报告" });

  const manifest = {
    task_id: taskId,
    product_name: productName,
    analysis_type: analysisType,
    analyzer: analyzerName,
    analysis_engine: reportData.analysis_engine || "llm_api",
    report_level: reportLevel,
    model: payload.model || runtimeConfig.analysisModel,
    llm_api_base: runtimeConfig.llmApiBase,
    llm_api_key_configured: Boolean(runtimeConfig.llmApiKey),
    llm_usage: reportData.llm_usage || null,
    taxonomy: reportData.taxonomy || [],
    products: reportData.products || [],
    multi_product_sheets: reportData.multi_product_sheets || [],
    supplemental_sheets: reportData.supplemental_sheets || [],
    status: finalStatus,
    input_files: savedFiles,
    baseline_files: savedFileGroups.baselineFiles,
    chat_files: savedFileGroups.chatFiles,
    reference_files: savedFileGroups.referenceFiles,
    output_files: {
      excel: excelPath,
      markdown: markdownPath,
      manifest: manifestPath
    },
    summary: finalSummary,
    created_at: now.text
  };
  await writeFile(manifestPath, JSON.stringify(manifest, null, 2), "utf8");

  const indexDir = path.join(historyRoot, "00_索引");
  await mkdir(indexDir, { recursive: true });
  await writeFile(path.join(indexDir, `${taskId}.json`), JSON.stringify(manifest, null, 2), "utf8");

  return {
    status: "completed",
    report_level: reportLevel,
    message: responseMessage,
    task_id: taskId,
    summary: finalSummary,
    files: {
      excel: excelPath,
      markdown: markdownPath,
      manifest: manifestPath,
      excel_download_url: downloadUrl(excelPath),
      markdown_download_url: downloadUrl(markdownPath),
      manifest_download_url: downloadUrl(manifestPath)
    }
  };
}

function publicJob(job) {
  return {
    task_id: job.task_id,
    status: job.status,
    stage: job.stage,
    progress: job.progress,
    completed_chunks: job.completed_chunks ?? null,
    total_chunks: job.total_chunks ?? null,
    message: job.message,
    created_at: job.created_at,
    updated_at: job.updated_at,
    error: job.error || null,
    result: job.result || null
  };
}

async function persistJob(job) {
  await mkdir(jobsRoot, { recursive: true });
  const target = path.join(jobsRoot, `${safeSegment(job.task_id)}.json`);
  const snapshot = JSON.stringify(publicJob(job), null, 2);
  job._persistChain = (job._persistChain || Promise.resolve())
    .catch(() => {})
    .then(async () => {
      const temporary = `${target}.${randomUUID()}.tmp`;
      await writeFile(temporary, snapshot, "utf8");
      await rename(temporary, target);
    });
  return job._persistChain;
}

async function updateJob(job, patch) {
  Object.assign(job, patch, { updated_at: new Date().toISOString() });
  await persistJob(job);
}

function normalizeJobError(error) {
  const raw = error instanceof Error ? error.message : String(error);
  const redacted = raw
    .replaceAll(projectRoot, "<project>")
    .replaceAll(workspaceRoot, "<workspace>");
  const isTimeout = /timed out|timeout|超时/i.test(raw);
  const isRateLimit = /\b429\b|rate.?limit|限流/i.test(raw);
  return {
    code: isTimeout ? "UPSTREAM_TIMEOUT" : isRateLimit ? "UPSTREAM_RATE_LIMIT" : "ANALYSIS_FAILED",
    message: isTimeout
      ? "模型分析服务响应超时，已完成的缓存会在重试时继续使用。"
      : isRateLimit
        ? "模型分析服务当前请求过多，请稍后重试。"
        : redacted.split(/\r?\n/).filter(Boolean).slice(-1)[0] || "分析未完成。",
    retryable: isTimeout || isRateLimit || /\b5\d\d\b/.test(raw),
    detail: redacted.slice(0, 2000)
  };
}

function drainJobQueue() {
  while (activeJobs < maxActiveJobs && jobQueue.length) {
    const job = jobQueue.shift();
    activeJobs += 1;
    void runJob(job);
  }
}

async function runJob(job) {
  await updateJob(job, {
    status: "running",
    stage: "starting",
    progress: 1,
    message: "任务已开始"
  });
  try {
    const result = await executeAnalysis(job._payload, job._files, (progress) => {
      void updateJob(job, {
        status: "running",
        stage: progress.stage || job.stage,
        progress: Math.max(job.progress || 0, Number(progress.progress || 0)),
        message: progress.message || job.message,
        completed_chunks: progress.completed_chunks,
        total_chunks: progress.total_chunks
      }).catch((error) => {
        console.error(`任务 ${job.task_id} 进度保存失败：`, error);
      });
    });
    await updateJob(job, {
      status: "completed",
      stage: "completed",
      progress: 100,
      message: "Excel 报告已生成",
      result,
      error: null
    });
  } catch (error) {
    const normalized = normalizeJobError(error);
    console.error(`任务 ${job.task_id} 分析失败：`, normalized);
    await updateJob(job, {
      status: "failed",
      stage: job.stage || "analysis",
      message: "本次分析没有完成",
      error: normalized
    });
  } finally {
    activeJobs -= 1;
    setImmediate(() => void drainJobQueue());
  }
}

async function handleCreateJob(req, res) {
  const contentLength = Number(req.headers["content-length"] || 0);
  if (contentLength > 30 * 1024 * 1024) {
    sendJson(res, 413, { status: "failed", message: "单次上传不能超过 30 MB。" });
    return;
  }
  const { payload, files } = await parseFormData(req);
  if (files.length > 10 || files.some((file) => file.size > 20 * 1024 * 1024)) {
    sendJson(res, 413, { status: "failed", message: "单次最多上传 10 个文件，单个文件不能超过 20 MB。" });
    return;
  }
  const taskId = randomUUID();
  payload.task_id = taskId;
  const job = {
    task_id: taskId,
    status: "queued",
    stage: "queued",
    progress: 0,
    message: activeJobs >= maxActiveJobs ? "任务已排队" : "任务已创建",
    created_at: new Date().toISOString(),
    updated_at: new Date().toISOString(),
    error: null,
    result: null,
    _payload: payload,
    _files: files
  };
  jobs.set(taskId, job);
  jobQueue.push(job);
  await persistJob(job);
  setImmediate(() => void drainJobQueue());
  sendJson(res, 202, {
    task_id: taskId,
    status: "queued",
    status_url: `/api/tasks/${taskId}`
  });
}

function findJob(taskId) {
  return jobs.get(String(taskId || ""));
}

async function handleTaskStatus(req, res) {
  const url = new URL(req.url || "/", `http://127.0.0.1:${port}`);
  const taskId = decodeURIComponent(url.pathname.split("/").filter(Boolean).pop() || "");
  const job = findJob(taskId);
  if (!job) {
    sendJson(res, 404, { status: "failed", message: "找不到该分析任务。" });
    return;
  }
  sendJson(res, 200, publicJob(job));
}

async function handleRetryJob(req, res) {
  const url = new URL(req.url || "/", `http://127.0.0.1:${port}`);
  const parts = url.pathname.split("/").filter(Boolean);
  const taskId = decodeURIComponent(parts[parts.length - 2] || "");
  const job = findJob(taskId);
  if (!job || job.status !== "failed" || !job._files?.length) {
    sendJson(res, 409, { status: "failed", message: "该任务当前不能重试。" });
    return;
  }
  await updateJob(job, {
    status: "queued",
    stage: "queued",
    progress: 0,
    message: "任务已重新排队",
    error: null,
    result: null
  });
  jobQueue.push(job);
  setImmediate(() => void drainJobQueue());
  sendJson(res, 202, publicJob(job));
}

async function loadPersistedJobs() {
  await mkdir(jobsRoot, { recursive: true });
  const entries = await readdir(jobsRoot, { withFileTypes: true });
  for (const entry of entries) {
    if (!entry.isFile() || !entry.name.endsWith(".json")) continue;
    try {
      const saved = JSON.parse(await readFile(path.join(jobsRoot, entry.name), "utf8"));
      if (!saved.task_id) continue;
      if (saved.status === "running" || saved.status === "queued") {
        saved.status = "failed";
        saved.message = "服务重启后任务已中止";
        saved.error = {
          code: "SERVICE_RESTARTED",
          message: "服务重启中断了任务，请重新上传文件后分析。",
          retryable: false,
          detail: ""
        };
      }
      jobs.set(saved.task_id, { ...saved, _payload: null, _files: null });
    } catch (error) {
      console.error(`无法恢复任务状态 ${entry.name}：`, error);
    }
  }
}

async function handleDownload(req, res) {
  const url = new URL(req.url || "/", `http://127.0.0.1:${port}`);
  const requested = url.searchParams.get("path");
  if (!requested) {
    sendText(res, 400, "Missing path");
    return;
  }

  const target = path.normalize(requested);
  const allowedRoot = path.normalize(historyRoot);
  const relative = path.relative(allowedRoot, target);
  if (!relative || relative.startsWith("..") || path.isAbsolute(relative) || !existsSync(target)) {
    sendText(res, 404, "File not found");
    return;
  }

  const ext = path.extname(target).toLowerCase();
  const filename = path.basename(target);
  const body = await readFile(target);
  res.writeHead(200, {
    "content-type": mimeTypes[ext] || "application/octet-stream",
    "content-disposition": `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`,
    "access-control-allow-origin": corsOrigin
  });
  res.end(body);
}

async function serveStatic(req, res) {
  const url = new URL(req.url || "/", `http://127.0.0.1:${port}`);
  const requestedPath = decodeURIComponent(url.pathname === "/" ? "/index.html" : url.pathname);
  const target = path.normalize(path.join(webRoot, requestedPath));
  const relative = path.relative(webRoot, target);
  if (!relative || relative.startsWith("..") || path.isAbsolute(relative)) {
    sendText(res, 403, "Forbidden");
    return;
  }
  if (!existsSync(target)) {
    sendText(res, 404, "Not found");
    return;
  }
  const ext = path.extname(target).toLowerCase();
  const body = await readFile(target);
  res.writeHead(200, {
    "content-type": mimeTypes[ext] || "application/octet-stream",
    "access-control-allow-origin": corsOrigin
  });
  res.end(body);
}

await loadPersistedJobs();

createServer(async (req, res) => {
  try {
    const requestUrl = new URL(req.url || "/", `http://127.0.0.1:${port}`);
    const pathname = requestUrl.pathname;
    if (req.method === "OPTIONS") {
      sendJson(res, 204, {});
      return;
    }
    if (req.method === "GET" && pathname === "/api/health") {
      sendJson(res, 200, {
        status: "ok",
        active_jobs: activeJobs,
        queued_jobs: jobQueue.length,
        max_active_jobs: maxActiveJobs
      });
      return;
    }
    if (req.method === "POST" && (pathname === "/api/chitu-analyze" || pathname === "/api/tasks")) {
      await handleCreateJob(req, res);
      return;
    }
    if (req.method === "POST" && /^\/api\/tasks\/[^/]+\/retry$/.test(pathname)) {
      await handleRetryJob(req, res);
      return;
    }
    if (req.method === "GET" && /^\/api\/tasks\/[^/]+$/.test(pathname)) {
      await handleTaskStatus(req, res);
      return;
    }
    if (req.method === "GET" && pathname === "/api/download") {
      await handleDownload(req, res);
      return;
    }
    if (req.method === "GET") {
      await serveStatic(req, res);
      return;
    }
    sendText(res, 405, "Method not allowed");
  } catch (error) {
    sendJson(res, 500, {
      status: "failed",
      message: error instanceof Error ? error.message : String(error)
    });
  }
}).listen(port, host, () => {
  console.log(`赤兔服务已启动：http://${host}:${port}/`);
});
