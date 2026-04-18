const crypto = require("crypto");
const fs = require("fs");
const fsp = require("fs/promises");
const path = require("path");
const { spawn } = require("child_process");
const express = require("express");
const multer = require("multer");
const JSZip = require("jszip");
const XLSX = require("xlsx");
const pdfParse = require("pdf-parse");
const { installWorkflowRoutes } = require("./workflow-service");

const app = express();
const PORT = process.env.PORT || 3000;
const ROOT_DIR = __dirname;
const PUBLIC_DIR = path.join(ROOT_DIR, "public");
const GENERATED_DIR = path.join(ROOT_DIR, "generated-images");
const DATA_DIR = path.join(ROOT_DIR, "data");
const LIBRARY_DOC_PATH = path.join(DATA_DIR, "studio-library.json");
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    files: 20,
    fileSize: 50 * 1024 * 1024,
  },
});

const REGION_MAP = {
  beijing: "https://dashscope.aliyuncs.com",
};
const GEMINI_IMAGE_MODELS = new Set(["gemini-3-pro-image-preview"]);
const SUPPORTED_GEMINI_ASPECTS = new Set(["1:1", "4:3", "16:9", "3:4", "9:16"]);

fs.mkdirSync(GENERATED_DIR, { recursive: true });
fs.mkdirSync(DATA_DIR, { recursive: true });

app.use(express.json({ limit: "50mb" }));
app.use(express.static(PUBLIC_DIR));
app.use("/generated-images", express.static(GENERATED_DIR));

function resolveRegion(region) {
  return REGION_MAP[region] || REGION_MAP.beijing;
}

function isGeminiImageModel(model) {
  return GEMINI_IMAGE_MODELS.has(String(model || "").trim());
}

function extensionToMimeType(extension) {
  const normalized = String(extension || "").toLowerCase();
  if (normalized === ".jpg" || normalized === ".jpeg") return "image/jpeg";
  if (normalized === ".webp") return "image/webp";
  if (normalized === ".bmp") return "image/bmp";
  return "image/png";
}

function pickExtensionFromMimeType(mimeType) {
  const mime = String(mimeType || "").toLowerCase();
  if (mime.includes("jpeg") || mime.includes("jpg")) return ".jpg";
  if (mime.includes("webp")) return ".webp";
  if (mime.includes("bmp")) return ".bmp";
  return ".png";
}

function inferAspectRatioFromSize(size) {
  const match = String(size || "").trim().match(/^(\d+)\*(\d+)$/);
  if (!match) return "";
  const width = Number(match[1]);
  const height = Number(match[2]);
  if (!width || !height) return "";
  const ratio = width / height;
  if (Math.abs(ratio - 1) < 0.02) return "1:1";
  if (Math.abs(ratio - (4 / 3)) < 0.02) return "4:3";
  if (Math.abs(ratio - (3 / 4)) < 0.02) return "3:4";
  if (Math.abs(ratio - (16 / 9)) < 0.02) return "16:9";
  if (Math.abs(ratio - (9 / 16)) < 0.02) return "9:16";
  return "";
}

function normalizeGeminiAspectRatio(aspectRatio, size) {
  const candidate = String(aspectRatio || inferAspectRatioFromSize(size) || "16:9").trim();
  return SUPPORTED_GEMINI_ASPECTS.has(candidate) ? candidate : "16:9";
}

function normalizeGeminiImageSize(size) {
  const value = String(size || "").trim().toUpperCase();
  if (["1K", "2K", "4K"].includes(value)) return value;
  const match = value.match(/^(\d+)\*(\d+)$/);
  if (!match) return "2K";
  const width = Number(match[1]);
  const height = Number(match[2]);
  const maxEdge = Math.max(width, height);
  if (maxEdge <= 1536) return "1K";
  if (maxEdge <= 3072) return "2K";
  return "4K";
}

function parseDataUrl(dataUrl) {
  const match = String(dataUrl || "").trim().match(/^data:([^;,]+)?;base64,(.+)$/i);
  if (!match) return null;
  const mimeType = match[1] || "image/png";
  return {
    mimeType,
    extension: pickExtensionFromMimeType(mimeType),
    buffer: Buffer.from(match[2], "base64"),
  };
}

async function saveGeneratedBufferToFile(buffer, { prefix = "result", requestId = "result", index = 1, extension = ".png" } = {}) {
  await fsp.mkdir(GENERATED_DIR, { recursive: true });
  const fileName = [
    sanitizeSegment(prefix, "result"),
    buildTimestamp(),
    sanitizeSegment(requestId, "result"),
    sanitizeSegment(index, "1"),
    crypto.randomUUID().slice(0, 8),
  ].join("_") + extension;
  const targetPath = path.join(GENERATED_DIR, fileName);
  await fsp.writeFile(targetPath, buffer);
  return {
    fileName,
    savedPath: targetPath,
    localUrl: `/generated-images/${encodeURIComponent(fileName)}`,
  };
}

async function loadGeminiImageSource(source) {
  const value = String(source || "").trim();
  if (!value) {
    throw new Error("Gemini 输入图片为空。");
  }

  const parsedDataUrl = parseDataUrl(value);
  if (parsedDataUrl) {
    return parsedDataUrl;
  }

  if (value.startsWith("/generated-images/")) {
    const fileName = decodeURIComponent(value.slice("/generated-images/".length));
    const targetPath = path.join(GENERATED_DIR, fileName);
    const buffer = await fsp.readFile(targetPath);
    const extension = path.extname(targetPath) || ".png";
    return {
      buffer,
      extension,
      mimeType: extensionToMimeType(extension),
    };
  }

  if (/^https?:\/\//i.test(value)) {
    const response = await fetch(value);
    if (!response.ok) {
      throw new Error(`读取 Gemini 输入图片失败，状态码 ${response.status}。`);
    }
    const contentType = response.headers.get("content-type") || "";
    const extension = pickFileExtension(value, contentType);
    return {
      buffer: Buffer.from(await response.arrayBuffer()),
      extension,
      mimeType: contentType || extensionToMimeType(extension),
    };
  }

  throw new Error("Gemini 目前只支持 data URL、本地缓存图或 http(s) 图片链接作为输入。");
}

async function convertMessagesToGeminiContents(messages) {
  const normalizedMessages = Array.isArray(messages) ? messages : [];
  const contents = [];

  for (const message of normalizedMessages) {
    const parts = [];
    const contentItems = Array.isArray(message?.content) ? message.content : [];

    for (const item of contentItems) {
      if (typeof item?.text === "string" && item.text.trim()) {
        parts.push({ text: item.text.trim() });
        continue;
      }

      if (typeof item?.image === "string" && item.image.trim()) {
        const image = await loadGeminiImageSource(item.image.trim());
        parts.push({
          inlineData: {
            mimeType: image.mimeType || "image/png",
            data: image.buffer.toString("base64"),
          },
        });
      }
    }

    if (!parts.length) continue;
    contents.push({
      role: message?.role === "assistant" || message?.role === "model" ? "model" : "user",
      parts,
    });
  }

  if (!contents.length) {
    throw new Error("Gemini 请求里没有可用的文本或图片输入。");
  }

  return contents;
}

async function buildGeminiGenerationBody({ payload, slideAspect }) {
  const size = payload?.parameters?.size;
  return {
    contents: await convertMessagesToGeminiContents(payload?.input?.messages),
    generationConfig: {
      responseModalities: ["TEXT", "IMAGE"],
      imageConfig: {
        aspectRatio: normalizeGeminiAspectRatio(slideAspect, size),
        imageSize: normalizeGeminiImageSize(size),
      },
    },
  };
}

function extractGeminiErrorMessage(data) {
  return String(
    data?.error?.message
    || data?.message
    || "",
  ).trim();
}

function buildGeminiModelUrl(model) {
  return `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}`;
}

async function requestJsonViaFetch({ url, method = "POST", headers = {}, body }) {
  const response = await fetch(url, {
    method,
    headers,
    body,
  });
  return parseJsonResponse(response);
}

async function requestJsonViaPowerShell({ url, method = "POST", headers = {}, body = "" }) {
  await fsp.mkdir(DATA_DIR, { recursive: true });
  const requestFile = path.join(DATA_DIR, `gemini-request-${crypto.randomUUID()}.json`);
  const scriptFile = path.join(DATA_DIR, `gemini-request-${crypto.randomUUID()}.ps1`);

  const script = [
    "param([string]$RequestFile)",
    "$ErrorActionPreference = 'Stop'",
    "[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)",
    "$request = Get-Content -Raw -Encoding UTF8 $RequestFile | ConvertFrom-Json",
    "$headers = @{}",
    "if ($request.headers) {",
    "  $request.headers.PSObject.Properties | ForEach-Object {",
    "    $headers[$_.Name] = [string]$_.Value",
    "  }",
    "}",
    "try {",
    "  if ($request.body) {",
    "    $response = Invoke-RestMethod -Method $request.method -Uri $request.url -Headers $headers -ContentType 'application/json; charset=utf-8' -Body ([string]$request.body)",
    "  } else {",
    "    $response = Invoke-RestMethod -Method $request.method -Uri $request.url -Headers $headers",
    "  }",
    "  @{ ok = $true; status = 200; data = $response } | ConvertTo-Json -Depth 100",
    "} catch {",
    "  $status = 500",
    "  $content = ''",
    "  if ($_.Exception.Response) {",
    "    try { $status = [int]$_.Exception.Response.StatusCode } catch {}",
    "    try {",
    "      $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())",
    "      $content = $reader.ReadToEnd()",
    "      $reader.Close()",
    "    } catch {}",
    "  }",
    "  if (-not $content) { $content = $_.Exception.Message }",
    "  $parsed = $null",
    "  try { $parsed = $content | ConvertFrom-Json } catch {}",
    "  if ($parsed) {",
    "    @{ ok = $false; status = $status; data = $parsed } | ConvertTo-Json -Depth 100",
    "  } else {",
    "    @{ ok = $false; status = $status; data = @{ message = [string]$content } } | ConvertTo-Json -Depth 100",
    "  }",
    "}",
  ].join("\r\n");

  await fsp.writeFile(requestFile, JSON.stringify({ url, method, headers, body }), "utf8");
  await fsp.writeFile(scriptFile, script, "utf8");

  try {
    const result = await new Promise((resolve, reject) => {
      const child = spawn("powershell.exe", ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", scriptFile, requestFile], {
        windowsHide: true,
      });

      let stdout = "";
      let stderr = "";

      child.stdout.on("data", (chunk) => {
        stdout += chunk.toString("utf8");
      });
      child.stderr.on("data", (chunk) => {
        stderr += chunk.toString("utf8");
      });
      child.on("error", reject);
      child.on("close", (code) => {
        if (code !== 0 && !stdout.trim()) {
          reject(new Error(stderr.trim() || `PowerShell request failed with code ${code}.`));
          return;
        }

        try {
          resolve(JSON.parse(stdout.trim() || "{}"));
        } catch (error) {
          reject(new Error(stderr.trim() || stdout.trim() || error.message));
        }
      });
    });
    return result;
  } finally {
    await Promise.allSettled([
      fsp.unlink(requestFile),
      fsp.unlink(scriptFile),
    ]);
  }
}

async function requestGeminiJson(request) {
  if (process.platform === "win32") {
    try {
      return await requestJsonViaPowerShell(request);
    } catch (error) {
      return requestJsonViaFetch(request);
    }
  }
  return requestJsonViaFetch(request);
}

async function normalizeGeminiGenerateResponse(data, model) {
  const candidate = Array.isArray(data?.candidates) ? data.candidates[0] : null;
  const parts = Array.isArray(candidate?.content?.parts) ? candidate.content.parts : [];
  const requestId = String(data?.responseId || data?.response_id || crypto.randomUUID()).trim();
  const content = [];
  let imageIndex = 0;

  for (const part of parts) {
    if (typeof part?.text === "string" && part.text.trim()) {
      content.push({ type: "text", text: part.text.trim() });
      continue;
    }

    const inlineData = part?.inlineData || part?.inline_data;
    if (!inlineData?.data) continue;
    imageIndex += 1;
    const mimeType = inlineData.mimeType || inlineData.mime_type || "image/png";
    const saved = await saveGeneratedBufferToFile(Buffer.from(inlineData.data, "base64"), {
      prefix: "gemini",
      requestId,
      index: imageIndex,
      extension: pickExtensionFromMimeType(mimeType),
    });
    content.push({ type: "image", image: saved.localUrl });
  }

  return {
    request_id: requestId,
    provider: "gemini",
    model,
    usage: data?.usageMetadata || null,
    output: {
      choices: [
        {
          message: {
            content,
          },
        },
      ],
    },
  };
}

function getDefaultLibraryDoc() {
  return {
    version: 1,
    savedAt: new Date().toISOString(),
    settings: {},
    themes: [],
    workflowPlans: [],
  };
}

async function readLibraryDoc() {
  try {
    const text = await fsp.readFile(LIBRARY_DOC_PATH, "utf8");
    const parsed = JSON.parse(text);
    return {
      ...getDefaultLibraryDoc(),
      ...parsed,
      settings: parsed?.settings && typeof parsed.settings === "object" ? parsed.settings : {},
      themes: Array.isArray(parsed?.themes) ? parsed.themes : [],
      workflowPlans: Array.isArray(parsed?.workflowPlans) ? parsed.workflowPlans : [],
    };
  } catch (error) {
    if (error.code === "ENOENT") {
      return getDefaultLibraryDoc();
    }
    throw error;
  }
}

async function writeLibraryDoc(doc) {
  const normalized = {
    ...getDefaultLibraryDoc(),
    ...doc,
    savedAt: new Date().toISOString(),
    settings: doc?.settings && typeof doc.settings === "object" ? doc.settings : {},
    themes: Array.isArray(doc?.themes) ? doc.themes : [],
    workflowPlans: Array.isArray(doc?.workflowPlans) ? doc.workflowPlans : [],
  };
  await fsp.mkdir(DATA_DIR, { recursive: true });
  await fsp.writeFile(LIBRARY_DOC_PATH, JSON.stringify(normalized, null, 2), "utf8");
  return normalized;
}

function buildGenerationUrl(region, asyncMode) {
  const host = resolveRegion(region);
  const endpoint = asyncMode
    ? "/api/v1/services/aigc/image-generation/generation"
    : "/api/v1/services/aigc/multimodal-generation/generation";
  return `${host}${endpoint}`;
}

function buildMultimodalUrl(region) {
  return `${resolveRegion(region)}/api/v1/services/aigc/multimodal-generation/generation`;
}

function buildTaskUrl(region, taskId) {
  return `${resolveRegion(region)}/api/v1/tasks/${taskId}`;
}

function buildResponsesUrl(region) {
  return `${resolveRegion(region)}/compatible-mode/v1/responses`;
}

function buildChatCompletionsUrl(region) {
  return `${resolveRegion(region)}/compatible-mode/v1/chat/completions`;
}

function sanitizeSegment(value, fallback) {
  const normalized = String(value || "")
    .replace(/[^a-zA-Z0-9_-]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "");
  return normalized || fallback;
}

function buildTimestamp() {
  const now = new Date();
  const parts = [
    now.getFullYear(),
    String(now.getMonth() + 1).padStart(2, "0"),
    String(now.getDate()).padStart(2, "0"),
    String(now.getHours()).padStart(2, "0"),
    String(now.getMinutes()).padStart(2, "0"),
    String(now.getSeconds()).padStart(2, "0"),
  ];
  return `${parts[0]}${parts[1]}${parts[2]}_${parts[3]}${parts[4]}${parts[5]}`;
}

function decodeXmlEntities(value) {
  return String(value || "")
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(parseInt(dec, 10)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, "&");
}

function normalizeExtractedText(value) {
  return String(value || "")
    .replace(/\r/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function detectExtension(fileName = "") {
  return path.extname(String(fileName || "")).toLowerCase();
}

function detectFileCategory(fileName = "", mimeType = "") {
  const extension = detectExtension(fileName);
  const mime = String(mimeType || "").toLowerCase();
  if ([".png", ".jpg", ".jpeg", ".bmp", ".gif"].includes(extension) || mime.startsWith("image/")) return "image";
  if ([".mp4", ".mkv", ".avi", ".mov", ".wmv", ".webm", ".flv"].includes(extension) || mime.startsWith("video/")) return "video";
  if ([".aac", ".amr", ".flac", ".m4a", ".mp3", ".mpeg", ".ogg", ".opus", ".wav", ".wma"].includes(extension) || mime.startsWith("audio/")) return "audio";
  return "document";
}

async function extractDocxText(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const entry = zip.file("word/document.xml");
  if (!entry) return "";
  const xml = await entry.async("string");
  const matches = [...xml.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)];
  return normalizeExtractedText(matches.map((item) => decodeXmlEntities(item[1])).join("\n"));
}

async function extractPptxText(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const slideFiles = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
    .sort((left, right) => {
      const leftIndex = Number(left.match(/slide(\d+)\.xml/i)?.[1] || 0);
      const rightIndex = Number(right.match(/slide(\d+)\.xml/i)?.[1] || 0);
      return leftIndex - rightIndex;
    });

  const slides = [];
  for (const slideName of slideFiles) {
    const xml = await zip.file(slideName).async("string");
    const texts = [...xml.matchAll(/<a:t[^>]*>([\s\S]*?)<\/a:t>/g)]
      .map((item) => decodeXmlEntities(item[1]))
      .filter(Boolean);
    if (texts.length) {
      const slideIndex = Number(slideName.match(/slide(\d+)\.xml/i)?.[1] || slides.length + 1);
      slides.push(`Slide ${slideIndex}\n${texts.join("\n")}`);
    }
  }
  return normalizeExtractedText(slides.join("\n\n"));
}

function extractSpreadsheetText(buffer) {
  const workbook = XLSX.read(buffer, { type: "buffer" });
  return normalizeExtractedText(
    workbook.SheetNames.slice(0, 6).map((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      const csv = XLSX.utils.sheet_to_csv(sheet, { blankrows: false });
      return `Sheet: ${sheetName}\n${csv}`;
    }).join("\n\n"),
  );
}

async function extractFileText(file) {
  const extension = detectExtension(file.originalname);
  const buffer = file.buffer;

  if ([".txt", ".md", ".csv", ".json", ".jsonl", ".html", ".htm", ".xml", ".yaml", ".yml"].includes(extension)) {
    return {
      extractedText: normalizeExtractedText(buffer.toString("utf8")),
      parseStatus: "parsed",
      parseNote: "已提取纯文本内容。",
    };
  }

  if (extension === ".docx") {
    return {
      extractedText: await extractDocxText(buffer),
      parseStatus: "parsed",
      parseNote: "已提取 DOCX 正文文本。",
    };
  }

  if (extension === ".pptx") {
    return {
      extractedText: await extractPptxText(buffer),
      parseStatus: "parsed",
      parseNote: "已提取 PPTX 各页文本。",
    };
  }

  if (extension === ".pdf") {
    const parsed = await pdfParse(buffer);
    return {
      extractedText: normalizeExtractedText(parsed.text || ""),
      parseStatus: "parsed",
      parseNote: "已提取 PDF 文本。",
    };
  }

  if ([".xls", ".xlsx"].includes(extension)) {
    return {
      extractedText: extractSpreadsheetText(buffer),
      parseStatus: "parsed",
      parseNote: "已提取表格内容。",
    };
  }

  if ([".png", ".jpg", ".jpeg", ".bmp", ".gif"].includes(extension)) {
    return {
      extractedText: "",
      parseStatus: "image",
      parseNote: "图片文件会作为视觉参考保留。",
    };
  }

  if ([".mp4", ".mkv", ".avi", ".mov", ".wmv", ".webm", ".flv"].includes(extension)) {
    return {
      extractedText: "",
      parseStatus: "metadata_only",
      parseNote: "已接收视频文件，当前仅保留文件元信息供人工参考。",
    };
  }

  if ([".aac", ".amr", ".flac", ".m4a", ".mp3", ".mpeg", ".ogg", ".opus", ".wav", ".wma"].includes(extension)) {
    return {
      extractedText: "",
      parseStatus: "metadata_only",
      parseNote: "已接收音频文件，当前仅保留文件元信息供人工参考。",
    };
  }

  if ([".doc", ".ppt", ".wps"].includes(extension)) {
    return {
      extractedText: "",
      parseStatus: "metadata_only",
      parseNote: "已接收旧版 Office/WPS 文件，当前仅保留文件元信息，请优先上传 docx/pptx 版本以便自动抽取文本。",
    };
  }

  return {
    extractedText: "",
    parseStatus: "metadata_only",
    parseNote: "文件已接收，当前未自动提取正文。",
  };
}

function pickFileExtension(imageUrl, contentType) {
  const pathname = (() => {
    try {
      return new URL(imageUrl).pathname.toLowerCase();
    } catch {
      return "";
    }
  })();

  if (pathname.endsWith(".jpg") || pathname.endsWith(".jpeg")) {
    return ".jpg";
  }
  if (pathname.endsWith(".webp")) {
    return ".webp";
  }
  if (pathname.endsWith(".bmp")) {
    return ".bmp";
  }
  if (pathname.endsWith(".png")) {
    return ".png";
  }

  const mime = String(contentType || "").toLowerCase();
  if (mime.includes("jpeg") || mime.includes("jpg")) {
    return ".jpg";
  }
  if (mime.includes("webp")) {
    return ".webp";
  }
  if (mime.includes("bmp")) {
    return ".bmp";
  }

  return ".png";
}

async function parseJsonResponse(response) {
  const text = await response.text();

  try {
    return {
      ok: response.ok,
      status: response.status,
      data: JSON.parse(text),
    };
  } catch {
    return {
      ok: response.ok,
      status: response.status,
      data: {
        code: "InvalidJSON",
        message: text || "DashScope returned a non-JSON response.",
      },
    };
  }
}

function extractResponsesOutputText(data) {
  if (typeof data?.output_text === "string" && data.output_text.trim()) {
    return data.output_text.trim();
  }

  const output = Array.isArray(data?.output) ? data.output : [];
  const parts = [];

  output.forEach((item) => {
    if (item?.type !== "message" || !Array.isArray(item.content)) return;
    item.content.forEach((contentItem) => {
      if (typeof contentItem?.text === "string" && contentItem.text.trim()) {
        parts.push(contentItem.text.trim());
      }
    });
  });

  return parts.join("\n\n").trim();
}

function extractChatCompletionText(data) {
  const content = data?.choices?.[0]?.message?.content;
  return typeof content === "string" ? content.trim() : "";
}

function stripMarkdownCodeFence(text) {
  const value = String(text || "").trim();
  const fenced = value.match(/^```(?:json)?\s*([\s\S]*?)\s*```$/i);
  return fenced ? fenced[1].trim() : value;
}

function safeParseJsonObject(text) {
  const normalized = stripMarkdownCodeFence(text);
  try {
    return JSON.parse(normalized);
  } catch {
    const match = normalized.match(/\{[\s\S]*\}/);
    if (!match) return null;
    try {
      return JSON.parse(match[0]);
    } catch {
      return null;
    }
  }
}

function normalizeResearchCandidates(candidates) {
  if (!Array.isArray(candidates)) return [];
  return candidates
    .map((item) => {
      const sources = Array.isArray(item?.sources)
        ? item.sources
          .map((source) => ({
            title: String(source?.title || "").trim(),
            url: String(source?.url || "").trim(),
          }))
          .filter((source) => source.title && source.url)
        : [];
      return {
        text: String(item?.text || "").trim(),
        why: String(item?.why || item?.reason || "").trim(),
        sources,
      };
    })
    .filter((item) => item.text && item.sources.length);
}

function extractResponsesError(data) {
  if (!data || typeof data !== "object") return "";
  if (typeof data?.error?.message === "string" && data.error.message.trim()) {
    return data.error.message.trim();
  }
  if (typeof data?.message === "string" && data.message.trim()) {
    return data.message.trim();
  }
  if (typeof data?.status === "string" && data.status === "failed") {
    return "联网补充请求失败。";
  }
  return "";
}

async function repairResearchOutputAsJson({ apiKey, region, rawText }) {
  const response = await fetch(buildChatCompletionsUrl(region), {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: "qwen3.5-plus",
      messages: [
        {
          role: "system",
          content: [
            "You are a JSON repair assistant.",
            "Return valid JSON only.",
            "Do not add markdown fences.",
            "Preserve meaning from the input and normalize it into the target schema.",
          ].join(" "),
        },
        {
          role: "user",
          content: [
            "Convert the following research output into JSON.",
            "JSON schema:",
            "{\"summary\":\"string\",\"candidates\":[{\"text\":\"string\",\"why\":\"string\",\"sources\":[{\"title\":\"string\",\"url\":\"https://...\"}]}]}",
            "If no candidates are valid, return an empty candidates array.",
            "Input:",
            String(rawText || ""),
          ].join("\n"),
        },
      ],
      response_format: {
        type: "json_schema",
        json_schema: {
          name: "research_supplements",
          strict: true,
          schema: {
            type: "object",
            additionalProperties: false,
            properties: {
              summary: { type: "string" },
              candidates: {
                type: "array",
                items: {
                  type: "object",
                  additionalProperties: false,
                  properties: {
                    text: { type: "string" },
                    why: { type: "string" },
                    sources: {
                      type: "array",
                      items: {
                        type: "object",
                        additionalProperties: false,
                        properties: {
                          title: { type: "string" },
                          url: { type: "string" },
                        },
                        required: ["title", "url"],
                      },
                    },
                  },
                  required: ["text", "why", "sources"],
                },
              },
            },
            required: ["summary", "candidates"],
          },
        },
      },
      stream: false,
    }),
  });

  const parsed = await parseJsonResponse(response);
  if (!parsed.ok) {
    return {
      ok: false,
      status: parsed.status,
      data: parsed.data,
    };
  }

  const repairedText = extractChatCompletionText(parsed.data);
  const repairedJson = safeParseJsonObject(repairedText);
  if (!repairedJson || typeof repairedJson !== "object") {
    return {
      ok: false,
      status: 502,
      data: {
        code: "InvalidRepairedResearchOutput",
        message: "联网补充结构化修复后仍无法解析为 JSON。",
        raw: repairedText || parsed.data,
      },
    };
  }

  return {
    ok: true,
    status: 200,
    data: repairedJson,
    raw: repairedText,
  };
}

async function generateResearchSearchQuery({ apiKey, region, pageType, pageTitle, pageContent, themeLabel, visibleTextBlock }) {
  const response = await fetch(buildChatCompletionsUrl(region), {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: "qwen3.5-plus",
      messages: [
        {
          role: "system",
          content: [
            "You create safe, focused web search queries for PPT research.",
            "Return JSON only.",
            "The query must stay on the page topic and avoid drifting into unrelated politics, entertainment, or social news.",
            "Prefer English or bilingual technical keywords when the page is about technology.",
          ].join(" "),
        },
        {
          role: "user",
          content: [
            "Generate one concise web search query for this PPT page.",
            "Goal: find factual supplements with reliable sources.",
            "Keep it under 16 words when possible.",
            "JSON schema: {\"query\":\"...\"}",
            `page_type: ${pageType || "content"}`,
            themeLabel ? `theme: ${themeLabel}` : "",
            pageTitle ? `page_title: ${pageTitle}` : "",
            pageContent ? `page_content: ${pageContent}` : "",
            visibleTextBlock ? `confirmed_visible_text: ${visibleTextBlock}` : "",
          ].filter(Boolean).join("\n"),
        },
      ],
      response_format: {
        type: "json_schema",
        json_schema: {
          name: "research_query",
          strict: true,
          schema: {
            type: "object",
            additionalProperties: false,
            properties: {
              query: { type: "string" },
            },
            required: ["query"],
          },
        },
      },
      stream: false,
    }),
  });

  const parsed = await parseJsonResponse(response);
  if (!parsed.ok) {
    return "";
  }
  const queryJson = safeParseJsonObject(extractChatCompletionText(parsed.data));
  return String(queryJson?.query || "").trim();
}

function buildHeuristicResearchQuery(pageTitle, pageContent) {
  const text = `${String(pageTitle || "")} ${String(pageContent || "")}`;
  const parts = [];
  const push = (value) => {
    if (value && !parts.includes(value)) parts.push(value);
  };

  if (/智能窗|智能玻璃|调光玻璃|Smart Window|Smart Glass/i.test(text)) {
    push("smart window");
    push("smart glass");
  }
  if (/电致变色|electrochromic/i.test(text)) push("electrochromic");
  if (/热致变色|thermochromic/i.test(text)) push("thermochromic");
  if (/光致变色|photochromic/i.test(text)) push("photochromic");
  if (/\bPDLC\b|液晶/i.test(text)) push("PDLC");
  if (/\bSPD\b/i.test(text)) push("SPD");
  if (/发展历史|演进|历程|概述|定义|分类|history|overview/i.test(text)) push("technology history");
  if (/建筑|节能|幕墙|采光|energy|building/i.test(text)) push("building energy saving");
  if (/应用|场景|市场|未来|趋势|application|market|future/i.test(text)) push("applications market");

  return parts.join(" ").trim();
}

app.get("/api/health", (_req, res) => {
  res.json({
    ok: true,
    service: "nanobanana-ppt-studio",
    port: PORT,
    generatedDir: GENERATED_DIR,
    libraryDocPath: LIBRARY_DOC_PATH,
    supportedRegions: Object.keys(REGION_MAP),
  });
});

app.get("/api/library", async (_req, res) => {
  try {
    const doc = await readLibraryDoc();
    return res.json({
      ok: true,
      doc,
    });
  } catch (error) {
    return res.status(500).json({
      code: "LibraryReadFailed",
      message: error.message || "读取项目资料库失败。",
    });
  }
});

app.post("/api/library", async (req, res) => {
  const { doc } = req.body || {};

  if (!doc || typeof doc !== "object") {
    return res.status(400).json({
      code: "InvalidLibraryDoc",
      message: "保存资料库时需要传入 doc 对象。",
    });
  }

  try {
    const savedDoc = await writeLibraryDoc(doc);
    return res.json({
      ok: true,
      doc: savedDoc,
    });
  } catch (error) {
    return res.status(500).json({
      code: "LibraryWriteFailed",
      message: error.message || "保存项目资料库失败。",
    });
  }
});

installWorkflowRoutes(app, {
  resolveRegion,
  parseJsonResponse,
  requestGeminiJson,
  normalizeGeminiGenerateResponse,
  buildGeminiModelUrl,
  normalizeGeminiAspectRatio,
  normalizeGeminiImageSize,
  parseDataUrl,
});

app.post("/api/test-image-key", async (req, res) => {
  const { apiKey, googleApiKey, region, model } = req.body || {};

  if (!model) {
    return res.status(400).json({
      code: "MissingModel",
      message: "测试 Key 需要传入当前图片模型。",
    });
  }

  if (isGeminiImageModel(model)) {
    if (!googleApiKey) {
      return res.status(400).json({
        code: "MissingGoogleApiKey",
        message: "请先填写 Google API Key。",
      });
    }

    try {
      const parsed = await requestGeminiJson({
        method: "GET",
        url: `${buildGeminiModelUrl(model)}?key=${encodeURIComponent(googleApiKey)}`,
      });

      if (!parsed.ok) {
        return res.status(parsed.status).json({
          code: "GeminiKeyTestFailed",
          message: extractGeminiErrorMessage(parsed.data) || "Google API Key 测试失败。",
          details: parsed.data,
        });
      }

      return res.json({
        ok: true,
        provider: "gemini",
        message: `Google API Key 可用，可访问 ${model}。`,
      });
    } catch (error) {
      return res.status(500).json({
        code: "GeminiKeyTestFailed",
        message: error.message || "Google API Key 测试失败。",
      });
    }
  }

  if (!apiKey) {
    return res.status(400).json({
      code: "MissingApiKey",
      message: "请先填写 DashScope / Qwen API Key。",
    });
  }

  try {
    const response = await fetch(buildChatCompletionsUrl(region), {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model: "qwen3.5-plus",
        messages: [
          { role: "user", content: "请只回复 OK" },
        ],
        stream: false,
      }),
    });

    const parsed = await parseJsonResponse(response);
    if (!parsed.ok) {
      return res.status(parsed.status).json(parsed.data);
    }

    return res.json({
      ok: true,
      provider: "dashscope",
      message: "DashScope / Qwen API Key 可用，可正常调用 Qwen。",
    });
  } catch (error) {
    return res.status(500).json({
      code: "DashScopeKeyTestFailed",
      message: error.message || "DashScope / Qwen API Key 测试失败。",
    });
  }
});

app.post("/api/generate", async (req, res, next) => {
  const { googleApiKey, slideAspect, payload } = req.body || {};

  if (!isGeminiImageModel(payload?.model)) {
    return next();
  }

  if (!googleApiKey) {
    return res.status(400).json({
      code: "MissingGoogleApiKey",
      message: "请先填写 Google API Key，再调用 Nano Banana。",
    });
  }

  try {
    const geminiBody = await buildGeminiGenerationBody({ payload, slideAspect });
    const parsed = await requestGeminiJson({
      method: "POST",
      url: `${buildGeminiModelUrl(payload.model)}:generateContent`,
      headers: {
        "x-goog-api-key": googleApiKey,
      },
      body: JSON.stringify(geminiBody),
    });

    if (!parsed.ok) {
      return res.status(parsed.status).json({
        code: "GeminiRequestFailed",
        message: extractGeminiErrorMessage(parsed.data) || "调用 Nano Banana 失败。",
        details: parsed.data,
      });
    }

    const normalized = await normalizeGeminiGenerateResponse(parsed.data, payload.model);
    return res.json(normalized);
  } catch (error) {
    return res.status(500).json({
      code: "GeminiProxyRequestFailed",
      message: error.message || "调用 Nano Banana 失败。",
    });
  }
});

app.post("/api/generate", async (req, res) => {
  const { apiKey, region, asyncMode, payload } = req.body || {};

  if (!apiKey) {
    return res.status(400).json({
      code: "MissingApiKey",
      message: "请先在页面中填写 API Key。",
    });
  }

  const missingPayloadFields = [];
  if (!payload || typeof payload !== "object") {
    missingPayloadFields.push("payload");
  } else {
    if (!payload.model) missingPayloadFields.push("model");
    if (!Array.isArray(payload.input?.messages) || !payload.input.messages.length) {
      missingPayloadFields.push("input.messages");
    }
  }

  if (missingPayloadFields.length) {
    return res.status(400).json({
      code: "InvalidPayload",
      message: `请求体不完整：缺少 ${missingPayloadFields.join("、")}。请检查模型、消息和参数配置。`,
      details: {
        missing: missingPayloadFields,
        receivedModel: payload?.model || "",
        messageCount: Array.isArray(payload?.input?.messages) ? payload.input.messages.length : 0,
      },
    });
  }

  if (!payload || !payload.model || !payload.input?.messages?.length) {
    return res.status(400).json({
      code: "InvalidPayload",
      message: "请求体不完整，请检查模型、消息和参数配置。",
    });
  }

  try {
    const headers = {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    };

    if (asyncMode) {
      headers["X-DashScope-Async"] = "enable";
    }

    const response = await fetch(buildGenerationUrl(region, asyncMode), {
      method: "POST",
      headers,
      body: JSON.stringify(payload),
    });

    const parsed = await parseJsonResponse(response);
    return res.status(parsed.status).json(parsed.data);
  } catch (error) {
    return res.status(500).json({
      code: "ProxyRequestFailed",
      message: error.message || "调用 DashScope / Qwen 失败。",
    });
  }
});

app.post("/api/tasks/fetch", async (req, res) => {
  const { apiKey, region, taskId } = req.body || {};

  if (!apiKey || !taskId) {
    return res.status(400).json({
      code: "MissingParams",
      message: "查询任务需要 API Key 和 task_id。",
    });
  }

  try {
    const response = await fetch(buildTaskUrl(region, taskId), {
      method: "GET",
      headers: {
        Authorization: `Bearer ${apiKey}`,
      },
    });

    const parsed = await parseJsonResponse(response);
    return res.status(parsed.status).json(parsed.data);
  } catch (error) {
    return res.status(500).json({
      code: "ProxyRequestFailed",
      message: error.message || "查询任务失败。",
    });
  }
});

app.post("/api/assistant", async (req, res) => {
  const { apiKey, region, payload } = req.body || {};

  if (!apiKey) {
    return res.status(400).json({
      code: "MissingApiKey",
      message: "调用提示词助手前请先填写 API Key。",
    });
  }

  if (!payload || !payload.model || !payload.input?.messages?.length) {
    return res.status(400).json({
      code: "InvalidPayload",
      message: "助手请求体不完整，请检查模型、消息和参数。",
    });
  }

  try {
    const response = await fetch(buildMultimodalUrl(region), {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify(payload),
    });

    const parsed = await parseJsonResponse(response);
    return res.status(parsed.status).json(parsed.data);
  } catch (error) {
    return res.status(500).json({
      code: "AssistantRequestFailed",
      message: error.message || "调用 Qwen 助手失败。",
    });
  }
});

app.post("/api/research-supplements", async (req, res) => {
  const { apiKey, region, page, themeLabel, visibleText, searchQuery: requestedSearchQuery } = req.body || {};

  if (!apiKey) {
    return res.status(400).json({
      code: "MissingApiKey",
      message: "调用联网补充前请先填写 API Key。",
    });
  }

  if (!page || typeof page !== "object") {
    return res.status(400).json({
      code: "MissingPage",
      message: "联网补充需要当前页面信息。",
    });
  }

  const pageTitle = String(page.pageTitle || "").trim();
  const pageContent = String(page.pageContent || "").trim();
  const pageType = String(page.pageType || "content").trim();
  const visibleTextBlock = String(visibleText || "").trim();

  if (!pageTitle && !pageContent) {
    return res.status(400).json({
      code: "MissingPageContent",
      message: "当前页面缺少可研究的标题或正文。",
    });
  }

  const heuristicQuery = buildHeuristicResearchQuery(pageTitle, pageContent);
  const generatedQuery = requestedSearchQuery
    ? ""
    : await generateResearchSearchQuery({
      apiKey,
      region,
      pageType,
      pageTitle,
      pageContent,
      themeLabel,
      visibleTextBlock,
    });
  const searchQuery = String(requestedSearchQuery || "").trim() || heuristicQuery || generatedQuery;

  const prompt = [
    "You are a lightweight web research assistant for PPT slide writing.",
    "Use web search only for the exact topic below and do not drift to unrelated subjects.",
    "Search the web and propose 0 to 4 small factual supplements for one slide.",
    "The supplements must be directly supported by sources and suitable for slide use after human review.",
    "Do not rewrite existing user text.",
    "Do not invent organizations, authors, dates, numbers, conclusions, or background stories.",
    "Prefer short milestone facts, concise term clarifications, representative applications, or one verified industry datapoint.",
    "Candidate text must be in Simplified Chinese and should stay short.",
    "If no reliable supplement is needed, return an empty candidates array.",
    "Return pure JSON only. No markdown. No prose.",
    "JSON schema:",
    "{\"summary\":\"...\",\"candidates\":[{\"text\":\"...\",\"why\":\"...\",\"sources\":[{\"title\":\"...\",\"url\":\"https://...\"}]}]}",
    "",
    searchQuery ? `search_focus_query: ${searchQuery}` : "",
  ].filter(Boolean).join("\n");

  try {
    const response = await fetch(buildResponsesUrl(region), {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model: "qwen3.5-plus",
        input: prompt,
        tools: [
          { type: "web_search" },
        ],
        enable_thinking: false,
      }),
    });

    const parsed = await parseJsonResponse(response);
    if (!parsed.ok) {
      return res.status(parsed.status).json(parsed.data);
    }

    const upstreamError = extractResponsesError(parsed.data);
    if (upstreamError) {
      return res.status(502).json({
        code: "ResearchToolError",
        message: upstreamError,
        raw: parsed.data,
      });
    }

    const outputText = extractResponsesOutputText(parsed.data);
    let parsedJson = safeParseJsonObject(outputText);
    let normalizedRaw = outputText || parsed.data;

    if (!parsedJson || typeof parsedJson !== "object") {
      const repaired = await repairResearchOutputAsJson({
        apiKey,
        region,
        rawText: outputText || JSON.stringify(parsed.data),
      });
      if (!repaired.ok) {
        return res.status(repaired.status).json(repaired.data);
      }
      parsedJson = repaired.data;
      normalizedRaw = repaired.raw || outputText || parsed.data;
    }

    if (!parsedJson || typeof parsedJson !== "object") {
      return res.status(502).json({
        code: "InvalidResearchOutput",
        message: "联网补充返回结果无法解析为 JSON。",
        raw: normalizedRaw,
      });
    }

    return res.json({
      ok: true,
      searchQuery,
      summary: String(parsedJson.summary || "").trim(),
      candidates: normalizeResearchCandidates(parsedJson.candidates),
      raw: normalizedRaw,
    });
  } catch (error) {
    return res.status(500).json({
      code: "ResearchSupplementFailed",
      message: error.message || "联网补充请求失败。",
    });
  }
});

app.post("/api/files/parse", upload.array("files", 20), async (req, res) => {
  const files = Array.isArray(req.files) ? req.files : [];

  if (!files.length) {
    return res.status(400).json({
      code: "MissingFiles",
      message: "请至少上传一个文件。",
    });
  }

  try {
    const parsedFiles = [];
    for (const file of files) {
      try {
        const parsed = await extractFileText(file);
        const extractedText = normalizeExtractedText(parsed.extractedText || "");
        parsedFiles.push({
          id: crypto.randomUUID(),
          name: file.originalname,
          extension: detectExtension(file.originalname),
          mimeType: file.mimetype || "",
          size: file.size || 0,
          category: detectFileCategory(file.originalname, file.mimetype),
          extractedText,
          previewText: extractedText.slice(0, 2400),
          parseStatus: parsed.parseStatus,
          parseNote: parsed.parseNote,
        });
      } catch (error) {
        parsedFiles.push({
          id: crypto.randomUUID(),
          name: file.originalname,
          extension: detectExtension(file.originalname),
          mimeType: file.mimetype || "",
          size: file.size || 0,
          category: detectFileCategory(file.originalname, file.mimetype),
          extractedText: "",
          previewText: "",
          parseStatus: "error",
          parseNote: error.message || "文件解析失败。",
        });
      }
    }

    return res.json({
      ok: true,
      files: parsedFiles,
    });
  } catch (error) {
    return res.status(500).json({
      code: "FileParseFailed",
      message: error.message || "文件解析失败。",
    });
  }
});

app.post("/api/download", async (req, res, next) => {
  const { imageUrl } = req.body || {};

  if (typeof imageUrl === "string" && imageUrl.startsWith("/generated-images/")) {
    const fileName = decodeURIComponent(imageUrl.slice("/generated-images/".length));
    const targetPath = path.join(GENERATED_DIR, fileName);

    try {
      await fsp.access(targetPath, fs.constants.F_OK);
      return res.json({
        ok: true,
        fileName,
        savedPath: targetPath,
        localUrl: `/generated-images/${encodeURIComponent(fileName)}`,
      });
    } catch (error) {
      return res.status(404).json({
        code: "LocalImageNotFound",
        message: error.message || "本地缓存图片不存在。",
      });
    }
  }

  const parsedDataUrl = parseDataUrl(imageUrl);
  if (parsedDataUrl) {
    try {
      const saved = await saveGeneratedBufferToFile(parsedDataUrl.buffer, {
        prefix: "inline",
        requestId: "inline-image",
        index: 1,
        extension: parsedDataUrl.extension,
      });
      return res.json({
        ok: true,
        fileName: saved.fileName,
        savedPath: saved.savedPath,
        localUrl: saved.localUrl,
      });
    } catch (error) {
      return res.status(500).json({
        code: "InlineImageSaveFailed",
        message: error.message || "保存内联图片失败。",
      });
    }
  }

  return next();
});

app.post("/api/download", async (req, res) => {
  const { imageUrl, requestId, index } = req.body || {};

  if (!imageUrl) {
    return res.status(400).json({
      code: "MissingImageUrl",
      message: "下载到本地需要 imageUrl。",
    });
  }

  try {
    await fsp.mkdir(GENERATED_DIR, { recursive: true });

    const response = await fetch(imageUrl);
    if (!response.ok) {
      return res.status(response.status).json({
        code: "DownloadFailed",
        message: `拉取图片失败，状态码 ${response.status}。`,
      });
    }

    const extension = pickFileExtension(
      imageUrl,
      response.headers.get("content-type"),
    );
    const fileName = [
      "wan27",
      buildTimestamp(),
      sanitizeSegment(requestId, "result"),
      sanitizeSegment(index, "1"),
      crypto.randomUUID().slice(0, 8),
    ].join("_") + extension;

    const targetPath = path.join(GENERATED_DIR, fileName);
    const buffer = Buffer.from(await response.arrayBuffer());
    await fsp.writeFile(targetPath, buffer);

    return res.json({
      ok: true,
      fileName,
      savedPath: targetPath,
      localUrl: `/generated-images/${encodeURIComponent(fileName)}`,
    });
  } catch (error) {
    return res.status(500).json({
      code: "DownloadFailed",
      message: error.message || "保存图片到本地失败。",
    });
  }
});

function startServer(port) {
  const server = app.listen(port, () => {
    console.log(`Nano Banana PPT Studio is running at http://localhost:${port}`);
    console.log(`Generated images will be saved to ${GENERATED_DIR}`);
  });

  server.on("error", (error) => {
    if (error?.code === "EADDRINUSE") {
      const nextPort = Number(port) + 1;
      console.warn(`Port ${port} is in use. Retrying on ${nextPort}...`);
      startServer(nextPort);
      return;
    }
    throw error;
  });
}

startServer(Number(PORT) || 3000);
