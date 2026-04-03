const crypto = require("crypto");
const fs = require("fs");
const fsp = require("fs/promises");
const path = require("path");
const express = require("express");

const app = express();
const PORT = process.env.PORT || 3000;
const ROOT_DIR = __dirname;
const PUBLIC_DIR = path.join(ROOT_DIR, "public");
const GENERATED_DIR = path.join(ROOT_DIR, "generated-images");

const REGION_MAP = {
  singapore: "https://dashscope-intl.aliyuncs.com",
  beijing: "https://dashscope.aliyuncs.com",
};

fs.mkdirSync(GENERATED_DIR, { recursive: true });

app.use(express.json({ limit: "50mb" }));
app.use(express.static(PUBLIC_DIR));
app.use("/generated-images", express.static(GENERATED_DIR));

function resolveRegion(region) {
  return REGION_MAP[region] || REGION_MAP.singapore;
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

app.get("/api/health", (_req, res) => {
  res.json({
    ok: true,
    service: "ppt-image-studio",
    port: PORT,
    generatedDir: GENERATED_DIR,
    supportedRegions: Object.keys(REGION_MAP),
  });
});

app.post("/api/generate", async (req, res) => {
  const { apiKey, region, asyncMode, payload } = req.body || {};

  if (!apiKey) {
    return res.status(400).json({
      code: "MissingApiKey",
      message: "请先在页面中填写 API Key。",
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
      message: error.message || "调用 DashScope 失败。",
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

app.listen(PORT, () => {
  console.log(`Wan 2.7 Image Studio is running at http://localhost:${PORT}`);
  console.log(`Generated images will be saved to ${GENERATED_DIR}`);
});
