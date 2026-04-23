const crypto = require("crypto");
const fs = require("fs");
const fsp = require("fs/promises");
const path = require("path");
const express = require("express");
const multer = require("multer");
const JSZip = require("jszip");
const PptxGenJS = require("pptxgenjs");
const XLSX = require("xlsx");
const pdfParse = require("pdf-parse");
const { installWorkflowRoutes } = require("./workflow-service");

const app = express();
const PORT = process.env.PORT || 3000;
const ROOT_DIR = __dirname;
const PUBLIC_DIR = path.join(ROOT_DIR, "public");
const GENERATED_DIR = path.join(ROOT_DIR, "generated-images");
const EXPORTS_DIR = path.join(ROOT_DIR, "exports");
const DATA_DIR = path.join(ROOT_DIR, "data");
const REFERENCE_ASSETS_DIR = path.join(DATA_DIR, "reference-assets");
const LIBRARY_DOC_PATH = path.join(DATA_DIR, "studio-library.json");

loadLocalEnv();

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
const GRSAI_HOSTS = {
  domestic: "https://grsai.dakka.com.cn",
  overseas: "https://grsaiapi.com",
};
const OPENAI_IMAGE_GENERATIONS_PATH = "/v1/images/generations";
const GEMINI_IMAGE_MODELS = new Set([
  "gemini-3.1-flash-image-preview",
  "gemini-3-pro-image-preview",
  "gemini-2.5-flash-image",
]);
const GRSAI_IMAGE_MODELS = new Set([
  "nano-banana-2",
  "nano-banana-pro",
  "gemini-3.1-pro",
]);
const OPENAI_IMAGE_MODELS = new Set([
  "gpt-image-2",
]);
const SUPPORTED_GEMINI_ASPECTS = new Set(["1:1", "4:3", "16:9", "3:4", "9:16"]);

fs.mkdirSync(GENERATED_DIR, { recursive: true });
fs.mkdirSync(EXPORTS_DIR, { recursive: true });
fs.mkdirSync(DATA_DIR, { recursive: true });
fs.mkdirSync(REFERENCE_ASSETS_DIR, { recursive: true });

app.use(express.json({ limit: "50mb" }));
app.get("/", (_req, res) => {
  res.redirect(302, "/v2/index.html");
});
app.use(express.static(PUBLIC_DIR));
app.use("/generated-images", express.static(GENERATED_DIR));
app.use("/exports", express.static(EXPORTS_DIR));
app.use("/reference-assets", express.static(REFERENCE_ASSETS_DIR));

function resolveRegion(region) {
  return REGION_MAP[region] || REGION_MAP.beijing;
}

function loadLocalEnv() {
  [".env.local", ".env"].forEach((fileName) => {
    const filePath = path.join(ROOT_DIR, fileName);
    if (!fs.existsSync(filePath)) return;
    const content = fs.readFileSync(filePath, "utf8");
    content.split(/\r?\n/).forEach((line) => {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith("#")) return;
      const match = trimmed.match(/^([A-Za-z_][A-Za-z0-9_]*)=(.*)$/);
      if (!match) return;
      const key = match[1];
      if (Object.prototype.hasOwnProperty.call(process.env, key)) return;
      process.env[key] = match[2].replace(/^["']|["']$/g, "");
    });
  });
}

function resolveDashScopeApiKey(apiKey) {
  return String(apiKey || process.env.DASHSCOPE_API_KEY || process.env.QWEN_API_KEY || "").trim();
}

function resolveHostedImageApiKey(apiKey) {
  return String(apiKey || process.env.OPENAI_IMAGE_API_KEY || process.env.OPENAI_API_KEY || process.env.GRSAI_API_KEY || process.env.GEMINI_API_KEY || process.env.GOOGLE_API_KEY || "").trim();
}

function resolveGrsaiApiKey(apiKey) {
  return String(apiKey || process.env.GRSAI_API_KEY || "").trim();
}

function resolveGeminiApiKey(apiKey) {
  return String(apiKey || process.env.GEMINI_API_KEY || process.env.GOOGLE_API_KEY || "").trim();
}

function resolveOpenAiImageApiKey(apiKey) {
  return String(apiKey || process.env.OPENAI_IMAGE_API_KEY || process.env.OPENAI_API_KEY || "").trim();
}

function isGeminiImageModel(model) {
  return GEMINI_IMAGE_MODELS.has(String(model || "").trim());
}

function isGrsaiImageModel(model) {
  return GRSAI_IMAGE_MODELS.has(String(model || "").trim());
}

function isOpenAiImageModel(model) {
  return OPENAI_IMAGE_MODELS.has(String(model || "").trim());
}

function isHostedImageModel(model) {
  return isGeminiImageModel(model) || isGrsaiImageModel(model) || isOpenAiImageModel(model);
}

function resolveGrsaiHost(host) {
  const value = String(host || process.env.GRSAI_HOST || "domestic").trim();
  if (/^https?:\/\//i.test(value)) return value.replace(/\/+$/, "");
  return (GRSAI_HOSTS[value] || GRSAI_HOSTS.domestic).replace(/\/+$/, "");
}

function resolveOpenAiImageEndpoint(endpointOrHost) {
  const value = String(
    endpointOrHost
    || process.env.OPENAI_IMAGE_GENERATIONS_URL
    || process.env.OPENAI_IMAGE_BASE_URL
    || process.env.OPENAI_BASE_URL
    || "https://api.openai.com",
  ).trim().replace(/\/+$/, "");
  if (new RegExp(`${OPENAI_IMAGE_GENERATIONS_PATH.replace(/\//g, "\\/")}$`, "i").test(value)) {
    return value;
  }
  return `${value}${OPENAI_IMAGE_GENERATIONS_PATH}`;
}

function extensionToMimeType(extension) {
  const normalized = String(extension || "").toLowerCase();
  if (normalized === ".jpg" || normalized === ".jpeg") return "image/jpeg";
  if (normalized === ".webp") return "image/webp";
  if (normalized === ".bmp") return "image/bmp";
  if (normalized === ".gif") return "image/gif";
  return "image/png";
}

function pickExtensionFromMimeType(mimeType) {
  const mime = String(mimeType || "").toLowerCase();
  if (mime.includes("jpeg") || mime.includes("jpg")) return ".jpg";
  if (mime.includes("webp")) return ".webp";
  if (mime.includes("bmp")) return ".bmp";
  if (mime.includes("gif")) return ".gif";
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

function normalizeOpenAiImageSize(size, slideAspect) {
  const raw = String(size || "").trim().toLowerCase().replace("*", "x");
  if (raw === "auto") return "auto";
  if (/^\d+x\d+$/.test(raw)) return raw;

  const quality = String(size || "").trim().toUpperCase();
  const aspect = String(slideAspect || "").trim();
  if (aspect === "1:1") {
    return quality === "4K" ? "2048x2048" : "1024x1024";
  }
  if (aspect === "4:3") {
    if (quality === "4K") return "2048x1536";
    if (quality === "1K") return "1024x768";
    return "1536x1152";
  }
  if (aspect === "9:16") {
    if (quality === "4K") return "2160x3840";
    if (quality === "1K") return "864x1536";
    return "1024x1536";
  }
  if (quality === "4K") return "3840x2160";
  if (quality === "1K") return "1536x864";
  return "2048x1152";
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

function toDataUrl(buffer, mimeType) {
  return `data:${mimeType || "application/octet-stream"};base64,${buffer.toString("base64")}`;
}

function getPptSlideSize(slideAspect) {
  switch (String(slideAspect || "").trim()) {
    case "4:3":
      return { width: 1024, height: 768 };
    case "1:1":
      return { width: 1080, height: 1080 };
    default:
      return { width: 1280, height: 720 };
  }
}

function getPptLayout(slideAspect) {
  switch (String(slideAspect || "").trim()) {
    case "4:3":
      return { name: "PPTGEN_4_3", width: 10, height: 7.5 };
    case "1:1":
      return { name: "PPTGEN_1_1", width: 7.5, height: 7.5 };
    default:
      return { name: "PPTGEN_16_9", width: 13.333, height: 7.5 };
  }
}

function toPptPosition(position, slideSize, layout) {
  return {
    x: (position.left / slideSize.width) * layout.width,
    y: (position.top / slideSize.height) * layout.height,
    w: (position.width / slideSize.width) * layout.width,
    h: (position.height / slideSize.height) * layout.height,
  };
}

async function saveReferenceImage(file) {
  const extension = pickExtensionFromMimeType(file.mimetype || extensionToMimeType(path.extname(file.originalname)));
  const assetId = crypto.randomUUID();
  const fileName = `${assetId}${extension}`;
  const targetPath = path.join(REFERENCE_ASSETS_DIR, fileName);
  await fsp.mkdir(REFERENCE_ASSETS_DIR, { recursive: true });
  await fsp.writeFile(targetPath, file.buffer);
  return {
    assetId,
    fileName,
    savedPath: targetPath,
    previewUrl: `/reference-assets/${encodeURIComponent(fileName)}`,
  };
}

async function loadReferenceAssetAsDataUrl(referenceFile) {
  const fileName = String(referenceFile?.assetFileName || "").trim()
    || (String(referenceFile?.previewUrl || "").startsWith("/reference-assets/")
      ? decodeURIComponent(String(referenceFile.previewUrl).slice("/reference-assets/".length))
      : "");
  if (!fileName || fileName.includes("/") || fileName.includes("\\")) return "";
  const targetPath = path.join(REFERENCE_ASSETS_DIR, fileName);
  const buffer = await fsp.readFile(targetPath);
  return toDataUrl(buffer, referenceFile?.mimeType || extensionToMimeType(path.extname(fileName)));
}

function normalizeExportText(value) {
  return String(value || "")
    .replace(/\r/g, "")
    .replace(/\u00a0/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function buildExportSlides(pages) {
  return (Array.isArray(pages) ? pages : []).map((page, index) => {
    const fallbackTitle = `第 ${index + 1} 页`;
    const title = normalizeExportText(page?.onscreenTitle || page?.pageTitle || fallbackTitle).split("\n")[0].trim();
    const body = normalizeExportText(page?.onscreenBody || page?.onscreenContent || page?.pageContent || "");
    return {
      pageNumber: Number(page?.pageNumber || index + 1),
      title: title || fallbackTitle,
      body,
      imageUrl: String(page?.imageUrl || page?.baseImage || "").trim(),
    };
  });
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

function buildGrsaiModelUrl(model, host) {
  return `${resolveGrsaiHost(host)}/v1beta/models/${encodeURIComponent(model)}`;
}

function buildGrsaiDrawUrl(host) {
  return `${resolveGrsaiHost(host)}/v1/draw/nano-banana`;
}

function buildGrsaiResultUrl(host) {
  return `${resolveGrsaiHost(host)}/v1/draw/result`;
}

function buildOpenAiImageGenerationsUrl(baseUrl) {
  return resolveOpenAiImageEndpoint(baseUrl);
}

async function requestJsonViaFetch({ url, method = "POST", headers = {}, body }) {
  const response = await fetch(url, {
    method,
    headers,
    body,
  });
  return parseJsonResponse(response);
}

async function requestGeminiJson(request) {
  return requestJsonViaFetch(request);
}

function buildOpenAiImageGenerationBody({ payload, slideAspect }) {
  const extracted = extractPromptAndImagesFromPayload(payload);
  if (!extracted.prompt) {
    const error = new Error("OpenAI 图片请求缺少提示词。");
    error.status = 400;
    throw error;
  }
  return {
    model: String(payload?.model || "gpt-image-2").trim() || "gpt-image-2",
    prompt: extracted.prompt,
    size: normalizeOpenAiImageSize(payload?.parameters?.size, slideAspect),
    response_format: "b64_json",
    ...(extracted.urls.length ? { image: extracted.urls } : {}),
  };
}

async function normalizeOpenAiImageGenerationResponse(data, model) {
  const requestId = String(data?.id || data?.request_id || crypto.randomUUID()).trim();
  const content = [];
  const items = Array.isArray(data?.data) ? data.data : [];

  let imageIndex = 0;
  for (const item of items) {
    if (typeof item?.revised_prompt === "string" && item.revised_prompt.trim()) {
      content.push({ type: "text", text: item.revised_prompt.trim() });
    }
    if (typeof item?.b64_json === "string" && item.b64_json.trim()) {
      imageIndex += 1;
      const saved = await saveGeneratedBufferToFile(Buffer.from(item.b64_json, "base64"), {
        prefix: "gpt_image_2",
        requestId,
        index: imageIndex,
        extension: ".png",
      });
      content.push({ type: "image", image: saved.localUrl });
    } else if (typeof item?.url === "string" && item.url.trim()) {
      content.push({ type: "image", image: item.url.trim() });
    }
  }

  return {
    request_id: requestId,
    provider: "openai-image",
    model,
    usage: data?.usage || null,
    output: {
      choices: [
        {
          message: {
            content,
          },
        },
      ],
    },
    raw: data,
  };
}

async function requestOpenAiImageGenerate({ apiKey, payload, slideAspect, baseUrl }) {
  const body = buildOpenAiImageGenerationBody({ payload, slideAspect });
  const parsed = await requestJsonViaFetch({
    method: "POST",
    url: buildOpenAiImageGenerationsUrl(baseUrl),
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify(body),
  });
  if (!parsed.ok) {
    const error = new Error(parsed.data?.error?.message || parsed.data?.message || "OpenAI 图片生成请求失败。");
    error.status = parsed.status || 500;
    error.details = parsed.data;
    throw error;
  }
  return normalizeOpenAiImageGenerationResponse(parsed.data, body.model);
}

async function normalizeGeminiGenerateResponse(data, model) {
  if (Array.isArray(data?.results)) {
    return normalizeGrsaiDrawResponse(data, model);
  }
  if (Array.isArray(data?.data?.results)) {
    return normalizeGrsaiDrawResponse(data.data, model);
  }
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
    if (inlineData?.data) {
      imageIndex += 1;
      const mimeType = inlineData.mimeType || inlineData.mime_type || "image/png";
      const saved = await saveGeneratedBufferToFile(Buffer.from(inlineData.data, "base64"), {
        prefix: "gemini",
        requestId,
        index: imageIndex,
        extension: pickExtensionFromMimeType(mimeType),
      });
      content.push({ type: "image", image: saved.localUrl });
      continue;
    }

    const fileUri = part?.fileData?.fileUri || part?.file_data?.file_uri;
    if (typeof fileUri === "string" && fileUri.trim()) {
      content.push({ type: "image", image: fileUri.trim() });
    }
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

function normalizeGrsaiDrawResponse(data, model) {
  const results = Array.isArray(data?.results) ? data.results : [];
  const content = [];
  results.forEach((item) => {
    if (typeof item?.content === "string" && item.content.trim()) {
      content.push({ type: "text", text: item.content.trim() });
    }
    if (typeof item?.url === "string" && item.url.trim()) {
      content.push({ type: "image", image: item.url.trim() });
    }
  });
  return {
    request_id: String(data?.id || crypto.randomUUID()),
    provider: "grsai",
    model,
    usage: null,
    output: {
      choices: [
        {
          message: {
            content,
          },
        },
      ],
    },
    raw: data,
  };
}

function extractPromptAndImagesFromPayload(payload) {
  const texts = [];
  const urls = [];
  const messages = Array.isArray(payload?.input?.messages) ? payload.input.messages : [];
  messages.forEach((message) => {
    const content = Array.isArray(message?.content) ? message.content : [];
    content.forEach((item) => {
      if (typeof item?.text === "string" && item.text.trim()) texts.push(item.text.trim());
      if (typeof item?.image === "string" && item.image.trim()) urls.push(item.image.trim());
    });
  });
  return {
    prompt: texts.join("\n\n").trim(),
    urls,
  };
}

async function requestGrsaiGenerate({ apiKey, model, payload, slideAspect, size, host }) {
  const geminiBody = await buildGeminiGenerationBody({ payload, slideAspect });
  const compatible = await requestJsonViaFetch({
    method: "POST",
    url: `${buildGrsaiModelUrl(model, host)}:generateContent`,
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
      "x-goog-api-key": apiKey,
    },
    body: JSON.stringify(geminiBody),
  });
  const compatibleHasResult = Array.isArray(compatible.data?.candidates)
    || Array.isArray(compatible.data?.results)
    || Array.isArray(compatible.data?.data?.results);
  if (compatible.ok && compatibleHasResult) {
    return normalizeGeminiGenerateResponse(compatible.data, model);
  }

  const fallback = extractPromptAndImagesFromPayload(payload);
  if (!fallback.prompt) {
    const error = new Error("Grsai 请求缺少提示词。");
    error.status = compatible.status || 500;
    error.details = compatible.data;
    throw error;
  }

  const draw = await requestJsonViaFetch({
    method: "POST",
    url: buildGrsaiDrawUrl(host),
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model,
      prompt: fallback.prompt,
      aspectRatio: normalizeGeminiAspectRatio(slideAspect, size),
      imageSize: normalizeGeminiImageSize(size || payload?.parameters?.size),
      urls: fallback.urls,
      webHook: "-1",
      shutProgress: false,
    }),
  });
  if (!draw.ok || draw.data?.code) {
    const error = new Error(draw.data?.msg || draw.data?.message || compatible.data?.message || "Grsai 生图请求失败。");
    error.status = draw.status || compatible.status || 500;
    error.details = { compatible: compatible.data, draw: draw.data };
    throw error;
  }

  const taskId = draw.data?.data?.id || draw.data?.id;
  if (!taskId) {
    return normalizeGrsaiDrawResponse(draw.data?.data || draw.data, model);
  }

  for (let attempt = 0; attempt < 60; attempt += 1) {
    await new Promise((resolve) => setTimeout(resolve, attempt < 3 ? 1000 : 2000));
    const result = await requestJsonViaFetch({
      method: "POST",
      url: buildGrsaiResultUrl(host),
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify({ id: taskId }),
    });
    if (!result.ok || result.data?.code) {
      continue;
    }
    const data = result.data?.data || result.data;
    if (data?.status === "succeeded") return normalizeGrsaiDrawResponse(data, model);
    if (data?.status === "failed") {
      const error = new Error(data.error || data.failure_reason || "Grsai 生图失败。");
      error.details = data;
      throw error;
    }
  }

  const timeout = new Error("Grsai 生图任务仍在运行，请稍后重试。");
  timeout.details = { taskId };
  throw timeout;
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

  if ([".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp"].includes(extension)) {
    const saved = await saveReferenceImage(file);
    return {
      extractedText: "",
      parseStatus: "image",
      parseNote: "图片文件会作为视觉参考保留。",
      assetId: saved.assetId,
      assetFileName: saved.fileName,
      previewUrl: saved.previewUrl,
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

  if (/鏅鸿兘绐梶鏅鸿兘鐜荤拑|璋冨厜鐜荤拑|Smart Window|Smart Glass/i.test(text)) {
    push("smart window");
    push("smart glass");
  }
  if (/鐢佃嚧鍙樿壊|electrochromic/i.test(text)) push("electrochromic");
  if (/鐑嚧鍙樿壊|thermochromic/i.test(text)) push("thermochromic");
  if (/鍏夎嚧鍙樿壊|photochromic/i.test(text)) push("photochromic");
  if (/\bPDLC\b|娑叉櫠/i.test(text)) push("PDLC");
  if (/\bSPD\b/i.test(text)) push("SPD");
  if (/鍙戝睍鍘嗗彶|婕旇繘|鍘嗙▼|姒傝堪|瀹氫箟|鍒嗙被|history|overview/i.test(text)) push("technology history");
  if (/寤虹瓚|鑺傝兘|骞曞|閲囧厜|energy|building/i.test(text)) push("building energy saving");
  if (/搴旂敤|鍦烘櫙|甯傚満|鏈潵|瓒嬪娍|application|market|future/i.test(text)) push("applications market");

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
    configuredKeys: {
      dashscope: Boolean(resolveDashScopeApiKey("")),
      hostedImage: Boolean(resolveHostedImageApiKey("")),
      openAiImage: Boolean(resolveOpenAiImageApiKey("")),
    },
    workflowModels: {
      assistant: process.env.WORKFLOW_ASSISTANT_MODEL || "qwen3.6-plus",
      style: process.env.WORKFLOW_STYLE_MODEL || process.env.QWEN_LIGHTWEIGHT_MODEL || "qwen-turbo-latest",
      jit: process.env.WORKFLOW_JIT_MODEL || process.env.QWEN_LIGHTWEIGHT_MODEL || "qwen-turbo-latest",
    },
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
  resolveGrsaiHost,
  resolveDashScopeApiKey,
  resolveHostedImageApiKey,
  resolveGrsaiApiKey,
  resolveGeminiApiKey,
  resolveOpenAiImageApiKey,
  parseJsonResponse,
  requestGeminiJson,
  requestGrsaiGenerate,
  requestOpenAiImageGenerate,
  normalizeGeminiGenerateResponse,
  buildGeminiModelUrl,
  buildGrsaiModelUrl,
  normalizeGeminiAspectRatio,
  normalizeGeminiImageSize,
  parseDataUrl,
  loadReferenceAssetAsDataUrl,
});

app.post("/api/test-image-key", async (req, res) => {
  const { apiKey, googleApiKey, grsaiApiKey, openAiImageApiKey, openAiImageBaseUrl, region, model, grsaiHost } = req.body || {};
  const effectiveApiKey = resolveDashScopeApiKey(apiKey);

  if (!model) {
    return res.status(400).json({
      code: "MissingModel",
      message: "测试 Key 需要传入当前图片模型。",
    });
  }

  if (isHostedImageModel(model)) {
    const effectiveGoogleApiKey = isOpenAiImageModel(model)
      ? resolveOpenAiImageApiKey(openAiImageApiKey || googleApiKey)
      : isGrsaiImageModel(model)
        ? resolveGrsaiApiKey(grsaiApiKey || googleApiKey)
        : resolveGeminiApiKey(googleApiKey);
    if (!effectiveGoogleApiKey) {
      return res.status(400).json({
        code: "MissingGoogleApiKey",
        message: "请先填写生图 API Key。",
      });
    }

    if (isOpenAiImageModel(model)) {
      return res.json({
        ok: true,
        provider: "openai-image",
        message: `OpenAI 图片接口 Key 已填写，模型 ${model} 将使用 ${resolveOpenAiImageEndpoint(openAiImageBaseUrl)}。`,
      });
    }

    if (isGrsaiImageModel(model)) {
      try {
        const parsed = await requestJsonViaFetch({
          method: "POST",
          url: buildGrsaiResultUrl(grsaiHost),
          headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${effectiveGoogleApiKey}`,
          },
          body: JSON.stringify({ id: "codex-health-check" }),
        });
        const serviceCode = Number(parsed.data?.code);
        if (!parsed.ok || (Number.isFinite(serviceCode) && serviceCode !== 0 && serviceCode !== -22)) {
          return res.status(parsed.ok ? 400 : parsed.status).json({
            code: "GrsaiKeyTestFailed",
            message: parsed.data?.msg || parsed.data?.message || "Grsai API Key 测试失败。",
            details: parsed.data,
          });
        }
        return res.json({
          ok: true,
          provider: "grsai",
          message: `Grsai API 可访问，模型 ${model} 将使用 ${resolveGrsaiHost(grsaiHost)}。`,
          details: parsed.data,
        });
      } catch (error) {
        return res.status(500).json({
          code: "GrsaiKeyTestFailed",
          message: error.message || "Grsai API Key 测试失败。",
        });
      }
    }

    try {
      const parsed = await requestGeminiJson({
        method: "GET",
        url: `${buildGeminiModelUrl(model)}?key=${encodeURIComponent(effectiveGoogleApiKey)}`,
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
        message: `Google API Key 可用，可以访问 ${model}。`,
      });
    } catch (error) {
      return res.status(500).json({
        code: "GeminiKeyTestFailed",
        message: error.message || "Google API Key 测试失败。",
      });
    }
  }

  if (!effectiveApiKey) {
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
        Authorization: `Bearer ${effectiveApiKey}`,
      },
      body: JSON.stringify({
        model: "qwen3.5-plus",
        messages: [
          { role: "user", content: "璇峰彧鍥炲 OK" },
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
      message: "DashScope / Qwen API Key 可用，可以正常调用 Qwen。",
    });
  } catch (error) {
    return res.status(500).json({
      code: "DashScopeKeyTestFailed",
      message: error.message || "DashScope / Qwen API Key 测试失败。",
    });
  }
});

app.post("/api/generate", async (req, res, next) => {
  const { googleApiKey, grsaiApiKey, openAiImageApiKey, openAiImageBaseUrl, slideAspect, payload, grsaiHost } = req.body || {};

  if (!isHostedImageModel(payload?.model)) {
    return next();
  }

  const effectiveGoogleApiKey = isOpenAiImageModel(payload.model)
    ? resolveOpenAiImageApiKey(openAiImageApiKey || googleApiKey)
    : isGrsaiImageModel(payload.model)
      ? resolveGrsaiApiKey(grsaiApiKey || googleApiKey)
      : resolveGeminiApiKey(googleApiKey);

  if (!effectiveGoogleApiKey) {
    return res.status(400).json({
      code: "MissingGoogleApiKey",
      message: "请先填写生图 API Key，再调用生图模型。",
    });
  }

  try {
    if (isOpenAiImageModel(payload.model)) {
      const normalized = await requestOpenAiImageGenerate({
        apiKey: effectiveGoogleApiKey,
        payload,
        slideAspect,
        baseUrl: openAiImageBaseUrl,
      });
      return res.json(normalized);
    }

    if (isGrsaiImageModel(payload.model)) {
      const normalized = await requestGrsaiGenerate({
        apiKey: effectiveGoogleApiKey,
        model: payload.model,
        payload,
        slideAspect,
        size: payload?.parameters?.size,
        host: grsaiHost,
      });
      return res.json(normalized);
    }

    const geminiBody = await buildGeminiGenerationBody({ payload, slideAspect });
    const parsed = await requestGeminiJson({
      method: "POST",
      url: `${buildGeminiModelUrl(payload.model)}:generateContent`,
      headers: {
        "x-goog-api-key": effectiveGoogleApiKey,
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
  const effectiveApiKey = resolveDashScopeApiKey(apiKey);

  if (!effectiveApiKey) {
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
    Authorization: `Bearer ${effectiveApiKey}`,
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
  const effectiveApiKey = resolveDashScopeApiKey(apiKey);

  if (!effectiveApiKey || !taskId) {
    return res.status(400).json({
      code: "MissingParams",
      message: "查询任务需要 API Key 和 task_id。",
    });
  }

  try {
    const response = await fetch(buildTaskUrl(region, taskId), {
      method: "GET",
      headers: {
        Authorization: `Bearer ${effectiveApiKey}`,
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
  const effectiveApiKey = resolveDashScopeApiKey(apiKey);

  if (!effectiveApiKey) {
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
        Authorization: `Bearer ${effectiveApiKey}`,
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
  const effectiveApiKey = resolveDashScopeApiKey(apiKey);

  if (!effectiveApiKey) {
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
      apiKey: effectiveApiKey,
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
        Authorization: `Bearer ${effectiveApiKey}`,
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
        apiKey: effectiveApiKey,
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
          assetId: parsed.assetId || "",
          assetFileName: parsed.assetFileName || "",
          previewUrl: parsed.previewUrl || "",
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

app.post("/api/export-workflow-ppt", async (req, res) => {
  const { projectTitle, slideAspect, pages } = req.body || {};
  const exportSlides = buildExportSlides(pages);

  if (!exportSlides.length) {
    return res.status(400).json({
      code: "MissingSlides",
      message: "导出 PPT 需要至少一页内容。",
    });
  }

  try {
    const slideSize = getPptSlideSize(slideAspect);
    const layout = getPptLayout(slideAspect);
    const pptx = new PptxGenJS();
    const deckTitle = normalizeExportText(projectTitle) || "智能生成导出";
    pptx.author = "Nano Banana PPT Studio";
    pptx.company = "PPTgen";
    pptx.subject = "Workflow PPT export";
    pptx.title = deckTitle;
    pptx.lang = "zh-CN";
    pptx.defineLayout(layout);
    pptx.layout = layout.name;

    for (const slideData of exportSlides) {
      const slide = pptx.addSlide();
      const margin = Math.round(slideSize.width * 0.05);
      const innerWidth = slideSize.width - margin * 2;
      const titleTop = Math.round(slideSize.height * 0.065);
      const titleHeight = Math.round(slideSize.height * 0.09);
      const pageBadgeWidth = Math.round(slideSize.width * 0.12);
      const imageTop = Math.round(slideSize.height * 0.19);
      const imageHeight = slideData.imageUrl ? Math.round(slideSize.height * 0.50) : 0;
      const bodyTop = slideData.imageUrl ? imageTop + imageHeight + Math.round(slideSize.height * 0.04) : imageTop;
      const bodyHeight = slideSize.height - bodyTop - margin;

      slide.background = { color: "F5F7FB" };
      slide.addShape(pptx.ShapeType.roundRect, {
        ...toPptPosition({
          left: Math.round(slideSize.width * 0.022),
          top: Math.round(slideSize.height * 0.03),
          width: Math.round(slideSize.width * 0.956),
          height: Math.round(slideSize.height * 0.94),
        }, slideSize, layout),
        fill: { color: "FFFFFF" },
        line: { color: "D9E2EC", width: 1 },
      });
      slide.addShape(pptx.ShapeType.roundRect, {
        ...toPptPosition({
          left: margin,
          top: Math.round(slideSize.height * 0.042),
          width: Math.round(slideSize.width * 0.11),
          height: Math.round(slideSize.height * 0.01),
        }, slideSize, layout),
        fill: { color: "2563EB" },
        line: { color: "2563EB", transparency: 100 },
      });
      slide.addText(slideData.title, {
        ...toPptPosition({
          left: margin,
          top: titleTop,
          width: innerWidth - pageBadgeWidth - 20,
          height: titleHeight,
        }, slideSize, layout),
        fontSize: Math.round(slideSize.height * 0.045),
        bold: true,
        color: "0F172A",
        fontFace: "Microsoft YaHei",
        valign: "mid",
        fit: "shrink",
      });
      slide.addText(`第 ${slideData.pageNumber} 页`, {
        ...toPptPosition({
          left: slideSize.width - margin - pageBadgeWidth,
          top: titleTop + 4,
          width: pageBadgeWidth,
          height: Math.round(slideSize.height * 0.05),
        }, slideSize, layout),
        shape: pptx.ShapeType.roundRect,
        fill: { color: "EEF4FF" },
        line: { color: "EEF4FF", transparency: 100 },
        fontSize: Math.round(slideSize.height * 0.022),
        color: "2563EB",
        bold: true,
        fontFace: "Microsoft YaHei",
        align: "center",
        valign: "mid",
      });

      if (slideData.imageUrl) {
        const image = await loadGeminiImageSource(slideData.imageUrl);
        slide.addImage({
          data: toDataUrl(image.buffer, image.mimeType),
          altText: slideData.title,
          ...toPptPosition({
            left: margin,
            top: imageTop,
            width: innerWidth,
            height: imageHeight,
          }, slideSize, layout),
        });
      }

      slide.addShape(pptx.ShapeType.roundRect, {
        ...toPptPosition({
          left: margin,
          top: bodyTop,
          width: innerWidth,
          height: Math.max(bodyHeight, Math.round(slideSize.height * 0.18)),
        }, slideSize, layout),
        fill: { color: "F8FAFC" },
        line: { color: "E2E8F0", width: 1 },
      });
      slide.addText(slideData.body || " ", {
        ...toPptPosition({
          left: margin + Math.round(slideSize.width * 0.018),
          top: bodyTop + Math.round(slideSize.height * 0.018),
          width: innerWidth - Math.round(slideSize.width * 0.036),
          height: Math.max(bodyHeight - Math.round(slideSize.height * 0.036), Math.round(slideSize.height * 0.14)),
        }, slideSize, layout),
        fontSize: slideData.imageUrl ? Math.round(slideSize.height * 0.024) : Math.round(slideSize.height * 0.03),
        color: "334155",
        fontFace: "Microsoft YaHei",
        margin: 0,
        fit: "shrink",
        breakLine: false,
      });
    }

    const fileName = `${sanitizeSegment(deckTitle, "workflow-export")}_${buildTimestamp()}_${crypto.randomUUID().slice(0, 6)}.pptx`;
    const targetPath = path.join(EXPORTS_DIR, fileName);
    await pptx.writeFile({ fileName: targetPath, compression: true });

    return res.json({
      ok: true,
      fileName,
      savedPath: targetPath,
      downloadUrl: `/exports/${encodeURIComponent(fileName)}`,
    });
  } catch (error) {
    return res.status(500).json({
      code: "WorkflowPptExportFailed",
      message: error.message || "导出 PPT 失败。",
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


