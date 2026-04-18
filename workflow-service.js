const crypto = require("crypto");

const WORKFLOW_ASSISTANT_MODEL = "qwen3.6-plus";
const WORKFLOW_IMAGE_MODEL = "gemini-3-pro-image-preview";
const DEFAULT_REGION = "beijing";
const DEFAULT_DECORATION_LEVEL = "medium";

const DEFAULT_PREFERENCES = {
  styleMode: "business",
  layoutVariety: "balanced",
  detailLevel: "polished",
  visualDensity: "balanced",
  compositionFocus: "balanced",
  dataNarrative: "balanced",
  pageMood: "modern",
};

const PREFERENCE_LABELS = {
  styleMode: { business: "商务", academic: "学术", creative: "创意" },
  layoutVariety: { uniform: "统一稳定", balanced: "平衡变化", diverse: "尽量多样" },
  detailLevel: { minimal: "偏简约", polished: "精致平衡", rich: "偏精细" },
  visualDensity: { airy: "留白更多", balanced: "均衡信息量", dense: "信息更满" },
  compositionFocus: { imageLead: "视觉主导", balanced: "图文平衡", textLead: "内容主导" },
  dataNarrative: { clean: "清晰克制", balanced: "适度信息图", expressive: "更强视觉化" },
  pageMood: { steady: "稳重统一", modern: "现代清爽", dramatic: "更有冲击" },
};

const DECORATION_LEVELS = {
  plain: "朴素",
  medium: "中等",
  complex: "复杂",
};

const workflowJobs = new Map();

function stringifyStructuredField(value) {
  if (typeof value === "string") return value.trim();
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  if (Array.isArray(value)) {
    return value.map((item) => stringifyStructuredField(item)).filter(Boolean).join("\n").trim();
  }
  if (value && typeof value === "object") {
    return Object.entries(value)
      .map(([key, item]) => {
        const text = stringifyStructuredField(item);
        return text ? `${key}: ${text}` : "";
      })
      .filter(Boolean)
      .join("\n")
      .trim();
  }
  return "";
}

function normalizeOnscreenContent(value) {
  const fieldPrefixPattern = /^\s*(title|subtitle|metainfo|metaInfo|abstract|summary|body|content|keypoints?|bullets?|visualelements?|datapoints?|metric|value|highlight|note|label|type)\s*[:：]\s*/i;
  const hiddenKeys = new Set(["type", "highlight"]);
  const listLikeKeys = new Set(["keypoints", "bullets", "datapoints", "visualelements", "items", "points"]);

  const stripFieldPrefix = (text) => String(text || "")
    .split(/\r?\n+/)
    .map((line) => line.replace(fieldPrefixPattern, "").trim())
    .filter((line) => !/^highlight\s*[:：]\s*(true|false)$/i.test(line))
    .filter(Boolean)
    .join("\n")
    .trim();

  const metaLabelMap = {
    presenter: "汇报人",
    reporter: "汇报人",
    author: "作者",
    studentid: "学号",
    studentId: "学号",
    class: "班级",
    time: "时间",
    date: "时间",
    unit: "单位",
    department: "院系",
    school: "学校",
  };

  const toText = (input, keyHint = "") => {
    const normalizedHint = String(keyHint || "").trim().toLowerCase();
    if (typeof input === "string" || typeof input === "number" || typeof input === "boolean") {
      return stripFieldPrefix(input);
    }
    if (Array.isArray(input)) {
      const lines = input.map((item) => toText(item, keyHint)).filter(Boolean);
      if (listLikeKeys.has(normalizedHint)) {
        return lines.map((line) => `• ${line.replace(/^[•-]\s*/, "")}`).join("\n");
      }
      return lines.join("\n");
    }
    if (input && typeof input === "object") {
      const entries = Object.entries(input);
      const order = ["title", "subtitle", "metaInfo", "metainfo", "summary", "abstract", "body", "content", "keyPoints", "bullets", "notes"];
      entries.sort((a, b) => {
        const ai = order.indexOf(a[0]);
        const bi = order.indexOf(b[0]);
        return (ai === -1 ? 999 : ai) - (bi === -1 ? 999 : bi);
      });
      return entries.map(([key, item]) => {
        const normalizedKey = String(key || "").trim();
        const lowerKey = normalizedKey.toLowerCase();
        if (hiddenKeys.has(lowerKey)) {
          return "";
        }
        if (["metainfo", "metainfo"].includes(lowerKey) && item && typeof item === "object" && !Array.isArray(item)) {
          return Object.entries(item)
            .map(([subKey, subValue]) => {
              const text = toText(subValue, subKey);
              if (!text) return "";
              const label = metaLabelMap[subKey] || metaLabelMap[String(subKey).toLowerCase()] || "";
              return label ? `${label}：${text}` : text;
            })
            .filter(Boolean)
            .join("  ");
        }
        if (item && typeof item === "object" && !Array.isArray(item)) {
          const labelText = toText(item.label, "label");
          const metricText = toText(item.metric, "metric");
          const valueText = toText(item.value, "value");
          const noteText = toText(item.note, "note");
          const dataPointsText = toText(item.dataPoints, "dataPoints");
          const headline = [labelText, metricText].filter(Boolean).join(" ");
          if (headline || valueText || noteText || dataPointsText) {
            let line = headline;
            if (valueText) {
              line = line ? `${line}：${valueText}` : valueText;
            }
            if (noteText) {
              line = line ? `${line}（${noteText}）` : noteText;
            }
            return [line, dataPointsText].filter(Boolean).join("\n");
          }
        }
        return toText(item, normalizedKey);
      }).filter(Boolean).join("\n");
    }
    return "";
  };

  return toText(value).trim();
}

function normalizeOnscreenContent(value) {
  const normalizeKey = (input) => String(input || "").replace(/[\s_-]+/g, "").toLowerCase();
  const wrapperKeys = new Set(["blocks", "items", "points", "entries", "sections", "visualelements", "datapoints"]);
  const hiddenKeys = new Set(["type", "highlight", "index", "order", "sort", "priority"]);
  const metaLabelMap = {
    presenter: "汇报人",
    reporter: "汇报人",
    author: "作者",
    studentid: "学号",
    studentId: "学号",
    class: "班级",
    time: "时间",
    date: "时间",
    unit: "单位",
    department: "院系",
    school: "学校",
  };

  const pushLine = (target, text) => {
    const clean = String(text || "").replace(/\s+/g, " ").trim();
    if (!clean) return;
    if (target[target.length - 1] === clean) return;
    target.push(clean);
  };

  const renderScalarText = (input) => {
    const source = String(input || "").trim();
    if (!source) return [];

    const lines = [];
    let pendingHeading = "";
    let pendingMetric = "";
    let pendingValue = "";
    let pendingNote = "";

    const flushHeading = () => {
      if (!pendingHeading) return;
      pushLine(lines, pendingHeading);
      pendingHeading = "";
    };

    const flushMetric = () => {
      if (!pendingMetric) return;
      let line = pendingMetric;
      if (pendingValue) line = `${line}：${pendingValue}`;
      if (pendingNote) line = `${line}（${pendingNote}）`;
      pushLine(lines, line);
      pendingMetric = "";
      pendingValue = "";
      pendingNote = "";
    };

    const flushPending = () => {
      flushMetric();
      flushHeading();
    };

    source.split(/\r?\n+/).forEach((rawLine) => {
      const line = String(rawLine || "").trim();
      if (!line) return;
      const match = line.match(/^([A-Za-z_][\w-]*)\s*[:：]\s*(.*)$/);
      if (!match) {
        flushPending();
        pushLine(lines, line);
        return;
      }

      const key = normalizeKey(match[1]);
      const content = String(match[2] || "").trim();
      if (wrapperKeys.has(key) || hiddenKeys.has(key)) return;

      if (key === "title") {
        flushPending();
        pushLine(lines, content);
        return;
      }

      if (key === "subtitle") {
        flushPending();
        pushLine(lines, content ? `副标题：${content}` : "");
        return;
      }

      if (key === "heading" || key === "label") {
        flushPending();
        pendingHeading = content;
        return;
      }

      if (key === "detail") {
        if (pendingHeading) {
          pushLine(lines, content ? `${pendingHeading}：${content}` : pendingHeading);
          pendingHeading = "";
        } else {
          pushLine(lines, content);
        }
        return;
      }

      if (key === "metric") {
        flushPending();
        pendingMetric = content;
        return;
      }

      if (key === "value") {
        pendingValue = content;
        return;
      }

      if (key === "note") {
        if (pendingMetric) {
          pendingNote = content;
          flushMetric();
        } else {
          pushLine(lines, content);
        }
        return;
      }

      if (["metainfo", "summary", "abstract", "body", "content", "text"].includes(key)) {
        flushPending();
        pushLine(lines, content);
        return;
      }

      flushPending();
      pushLine(lines, content || line);
    });

    flushMetric();
    flushHeading();
    return lines;
  };

  const renderLines = (input, keyHint = "") => {
    if (input == null) return [];
    if (typeof input === "string" || typeof input === "number" || typeof input === "boolean") {
      return renderScalarText(input);
    }

    if (Array.isArray(input)) {
      return input.flatMap((item) => renderLines(item, keyHint));
    }

    if (typeof input === "object") {
      const objectKeys = Object.keys(input);
      if (objectKeys.length === 1 && wrapperKeys.has(normalizeKey(objectKeys[0]))) {
        return renderLines(input[objectKeys[0]], objectKeys[0]);
      }

      if (["metainfo", "metaInfo"].some((key) => Object.prototype.hasOwnProperty.call(input, key))) {
        const meta = input.metaInfo || input.metainfo;
        const pieces = Object.entries(meta || {})
          .map(([subKey, subValue]) => {
            const text = renderLines(subValue, subKey).join(" ");
            if (!text) return "";
            const label = metaLabelMap[subKey] || metaLabelMap[normalizeKey(subKey)] || "";
            return label ? `${label}：${text}` : text;
          })
          .filter(Boolean);
        const rest = objectKeys
          .filter((key) => !["metainfo", "metaInfo"].includes(key))
          .flatMap((key) => renderLines(input[key], key));
        return [...(pieces.length ? [pieces.join("，")] : []), ...rest];
      }

      if (Object.prototype.hasOwnProperty.call(input, "metric") || Object.prototype.hasOwnProperty.call(input, "value")) {
        const metric = renderLines(input.metric, "metric").join(" ");
        const valueText = renderLines(input.value, "value").join(" ");
        const noteText = renderLines(input.note, "note").join(" ");
        const label = renderLines(input.label, "label").join(" ");
        const heading = renderLines(input.heading, "heading").join(" ");
        let line = [label || heading, metric].filter(Boolean).join(" ");
        if (valueText) line = line ? `${line}：${valueText}` : valueText;
        if (noteText) line = line ? `${line}（${noteText}）` : noteText;
        const extra = ["detail", "summary", "abstract", "content", "body", "text"]
          .flatMap((key) => renderLines(input[key], key));
        return [...(line ? [line] : []), ...extra];
      }

      if (Object.prototype.hasOwnProperty.call(input, "heading") || Object.prototype.hasOwnProperty.call(input, "detail")) {
        const heading = renderLines(input.heading || input.label, "heading").join(" ");
        const detail = renderLines(input.detail || input.content || input.summary, "detail").join(" ");
        const line = heading ? (detail ? `${heading}：${detail}` : heading) : detail;
        return line ? [line] : [];
      }

      return objectKeys.flatMap((key) => {
        const normalizedKey = normalizeKey(key);
        if (wrapperKeys.has(normalizedKey) || hiddenKeys.has(normalizedKey)) {
          return renderLines(input[key], key);
        }
        return renderLines(input[key], key);
      });
    }

    return [];
  };

  return renderLines(value).join("\n").trim();
}

function stripMarkdownCodeFence(text) {
  const source = String(text || "").trim();
  if (source.startsWith("```")) {
    return source.replace(/^```[\w-]*\n?/, "").replace(/\n?```$/, "").trim();
  }
  return source;
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

function extractAssistantMessageText(response) {
  const content = response?.output?.choices?.[0]?.message?.content;
  if (typeof content === "string") return content.trim();
  if (!Array.isArray(content)) return "";
  return content
    .map((item) => {
      if (typeof item === "string") return item;
      if (typeof item?.text === "string") return item.text;
      return "";
    })
    .join("\n")
    .trim();
}

function installWorkflowRoutes(app, deps) {
  const {
    resolveRegion,
    parseJsonResponse,
    requestGeminiJson,
    normalizeGeminiGenerateResponse,
    buildGeminiModelUrl,
    normalizeGeminiAspectRatio,
    normalizeGeminiImageSize,
    parseDataUrl,
  } = deps;

  async function callAssistant(apiKey, region, payload) {
    const response = await fetch(`${resolveRegion(region || DEFAULT_REGION)}/api/v1/services/aigc/multimodal-generation/generation`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify(payload),
    });
    const parsed = await parseJsonResponse(response);
    if (!parsed.ok || parsed.data?.code) {
      throw new Error(parsed.data?.message || "Qwen 调用失败。");
    }
    const text = extractAssistantMessageText(parsed.data);
    if (!text) throw new Error("Qwen 返回为空。");
    return { data: parsed.data, text };
  }

  function buildAssistantPayload(systemPrompt, userPrompt) {
    return {
      model: WORKFLOW_ASSISTANT_MODEL,
      input: {
        messages: [
          { role: "system", content: [{ text: systemPrompt }] },
          { role: "user", content: [{ text: userPrompt }] },
        ],
      },
      parameters: {
        result_format: "message",
        response_format: { type: "json_object" },
        enable_thinking: false,
      },
    };
  }

  async function runAssistantJsonObject(apiKey, region, systemPrompt, userPrompt, moduleName) {
    const payload = buildAssistantPayload(systemPrompt, userPrompt);
    const { text } = await callAssistant(apiKey, region, payload);
    const parsed = safeParseJsonObject(text);
    if (!parsed || typeof parsed !== "object") {
      throw new Error(`${moduleName} 返回的 JSON 无法解析。`);
    }
    return { payload, text, parsed };
  }

  function normalizePreferenceValue(key, value) {
    const normalized = String(value || "").trim();
    return Object.prototype.hasOwnProperty.call(PREFERENCE_LABELS[key] || {}, normalized)
      ? normalized
      : DEFAULT_PREFERENCES[key];
  }

  function normalizePreferences(preferences = {}) {
    return {
      styleMode: normalizePreferenceValue("styleMode", preferences.styleMode),
      layoutVariety: normalizePreferenceValue("layoutVariety", preferences.layoutVariety),
      detailLevel: normalizePreferenceValue("detailLevel", preferences.detailLevel),
      visualDensity: normalizePreferenceValue("visualDensity", preferences.visualDensity),
      compositionFocus: normalizePreferenceValue("compositionFocus", preferences.compositionFocus),
      dataNarrative: normalizePreferenceValue("dataNarrative", preferences.dataNarrative),
      pageMood: normalizePreferenceValue("pageMood", preferences.pageMood),
    };
  }

  function normalizeDecorationLevel(level) {
    const normalized = String(level || "").trim().toLowerCase();
    return Object.prototype.hasOwnProperty.call(DECORATION_LEVELS, normalized) ? normalized : DEFAULT_DECORATION_LEVEL;
  }

  function getDecorationLevelLabel(level) {
    return DECORATION_LEVELS[normalizeDecorationLevel(level)];
  }

  function buildPreferencePromptBlock(preferences) {
    return [
      "【用户偏好】",
      `风格模式：${PREFERENCE_LABELS.styleMode[preferences.styleMode]}。`,
      `版式变化：${PREFERENCE_LABELS.layoutVariety[preferences.layoutVariety]}。`,
      `视觉细节：${PREFERENCE_LABELS.detailLevel[preferences.detailLevel]}。`,
      `留白与信息量：${PREFERENCE_LABELS.visualDensity[preferences.visualDensity]}。`,
      `图文主次：${PREFERENCE_LABELS.compositionFocus[preferences.compositionFocus]}。`,
      `数据页表达：${PREFERENCE_LABELS.dataNarrative[preferences.dataNarrative]}。`,
      `整体气质：${PREFERENCE_LABELS.pageMood[preferences.pageMood]}。`,
      "这些偏好只影响结构与视觉表达，不得新增用户未提供的事实。",
    ].join("\n");
  }

  function buildThemeDefinitionBlock(themeDefinition) {
    if (!themeDefinition) return "";
    return [
      "【全局主题模板】",
      themeDefinition.modelPrompt ? `模型总纲：${themeDefinition.modelPrompt}` : "",
      themeDefinition.basic ? `Basic：${themeDefinition.basic}` : "",
      themeDefinition.cover ? `Cover：${themeDefinition.cover}` : "",
      themeDefinition.content ? `Content：${themeDefinition.content}` : "",
      themeDefinition.data ? `Data：${themeDefinition.data}` : "",
      themeDefinition.decorationLevel ? `装饰强度：${getDecorationLevelLabel(themeDefinition.decorationLevel)}` : "",
    ].filter(Boolean).join("\n");
  }

  function buildHardConstraintBlock() {
    return [
      "【硬约束】",
      "只允许使用用户主文本和参考材料明确支持的信息，禁止臆造事实。",
      "装饰只限于无字图形、纹理、容器、线条、光效和图标，禁止新增 logo、水印、页码、角标或无关小字。",
      "页面主标题与页面主正文保持明显层级差，建议约 1.5-2 倍；二级和三级标题只做温和层级差。",
      "避免把一整页内容压成密集小字墙；优先保留留白与可读性。",
      "最终画面必须像 PPT 页面，而不是海报、长图或 UI 仪表盘。",
    ].join("\n");
  }

  function countCharacters(text) {
    const source = String(text || "").replace(/\s+/g, "");
    let total = 0;
    for (const char of source) {
      if (/[\u4e00-\u9fff]/u.test(char)) {
        total += 1;
      } else if (/[A-Za-z]/.test(char)) {
        total += 0.45;
      } else if (/[0-9]/.test(char)) {
        total += 0.4;
      } else {
        total += 0.2;
      }
    }
    return Math.round(total);
  }

  function normalizeReferenceFiles(referenceFiles = []) {
    return Array.isArray(referenceFiles)
      ? referenceFiles
        .map((item) => ({
          id: String(item?.id || crypto.randomUUID()),
          name: String(item?.name || "").trim(),
          category: String(item?.category || "").trim(),
          parseStatus: String(item?.parseStatus || "").trim(),
          parseNote: String(item?.parseNote || "").trim(),
          extractedText: String(item?.extractedText || "").trim(),
          previewText: String(item?.previewText || "").trim(),
          includeInSplit: item?.includeInSplit !== false,
        }))
        .filter((item) => item.name)
      : [];
  }

  function buildReferenceDigestInput(mainText, referenceFiles) {
    const lines = ["【用户主文本】", mainText, "", "【参考材料】"];
    referenceFiles.forEach((file, index) => {
      lines.push(`文件 ${index + 1}：${file.name}（${file.category || "unknown"}）`);
      lines.push(file.extractedText || file.previewText || "");
      lines.push("");
    });
    return lines.join("\n");
  }

  function normalizeThemeDefinition(result, fallbackThemeName, decorationLevel, preferences) {
    const normalized = {
      displaySummaryZh: stringifyStructuredField(result?.displaySummaryZh || ""),
      modelPrompt: stringifyStructuredField(result?.modelPrompt || ""),
      basic: stringifyStructuredField(result?.basic || ""),
      cover: stringifyStructuredField(result?.cover || ""),
      content: stringifyStructuredField(result?.content || ""),
      data: stringifyStructuredField(result?.data || ""),
      decorationLevel: normalizeDecorationLevel(decorationLevel),
      preferences,
      themeName: String(fallbackThemeName || "").trim(),
    };
    if (!normalized.displaySummaryZh) {
      normalized.displaySummaryZh = [
        fallbackThemeName ? `主题：${fallbackThemeName}` : "",
        normalized.basic ? `基础风格：${normalized.basic}` : "",
        normalized.content ? `内容页：${normalized.content}` : "",
      ].filter(Boolean).join("\n");
    }
    if (!normalized.modelPrompt) {
      normalized.modelPrompt = [normalized.basic, normalized.cover, normalized.content, normalized.data].filter(Boolean).join("\n");
    }
    return normalized;
  }

  function normalizePageType(value, index) {
    const normalized = String(value || "").trim().toLowerCase();
    if (["cover", "content", "data"].includes(normalized)) return normalized;
    return index === 0 ? "cover" : "content";
  }

  function normalizeSplitRisk(value, estimatedChars) {
    const normalized = String(value || "").trim().toLowerCase();
    if (["low", "medium", "high"].includes(normalized)) return normalized;
    if (estimatedChars > 250) return "high";
    if (estimatedChars > 200) return "medium";
    return "low";
  }

  function normalizeContentBand(value, fallbackText = "") {
    const normalized = String(value || "").trim().toLowerCase();
    if (["minimal", "balanced", "standard", "dense"].includes(normalized)) return normalized;
    const count = countCharacters(fallbackText);
    if (count < 50) return "minimal";
    if (count <= 150) return "balanced";
    if (count <= 250) return "standard";
    return "dense";
  }

  function normalizePagePlan(pagePlan = [], decorationLevel) {
    return (Array.isArray(pagePlan) ? pagePlan : [])
      .map((item, index) => {
        const pageContent = stringifyStructuredField(item?.pageContent || item?.page_content || "");
        const estimatedChars = Number(item?.estimatedChars || item?.estimated_chars || countCharacters(pageContent));
        return {
          id: crypto.randomUUID(),
          pageNumber: Number(item?.pageNumber || item?.page_number || index + 1),
          pageType: normalizePageType(item?.pageType || item?.page_type, index),
          pageTitle: stringifyStructuredField(item?.pageTitle || item?.page_title || `第 ${index + 1} 页`),
          pageContent,
          sectionTopic: stringifyStructuredField(item?.sectionTopic || item?.section_topic || ""),
          estimatedChars,
          splitRisk: normalizeSplitRisk(item?.splitRisk || item?.split_risk, estimatedChars),
          recommendedBand: normalizeContentBand(item?.recommendedBand || item?.recommended_band, pageContent),
          decorationLevel: normalizeDecorationLevel(item?.decorationLevel || decorationLevel),
        };
      })
      .filter((page) => page.pageTitle || page.pageContent);
  }

  function buildEmptyQualityResult() {
    return {
      pass: false,
      severity: "medium",
      issues: [],
      suggestions: [],
      metrics: {
        charCount: 0,
        estimatedMinFont: 18,
        contrastRisk: "manual-check",
        whitespaceBand: "balanced",
        fontFamilyCount: 1,
      },
      checklist: "",
    };
  }

  function createWorkflowPage(rawPage) {
    return {
      ...rawPage,
      splitDone: true,
      prepareDone: false,
      readyToGenerate: false,
      generated: false,
      generationStatus: "idle",
      generationError: "",
      onscreenContent: "",
      contentBand: rawPage.recommendedBand || "balanced",
      overflowFlag: false,
      overflowReason: "",
      revisionHint: "",
      layoutSummary: "",
      textHierarchy: "",
      visualFocus: "",
      readabilityNotes: "",
      pagePrompt: "",
      qualityResult: buildEmptyQualityResult(),
      qualityPass: false,
      riskLevel: rawPage.splitRisk === "high" ? "high" : rawPage.splitRisk === "medium" ? "medium" : "none",
      riskReason: rawPage.splitRisk === "high" ? "拆分阶段已标记为高风险页。" : "",
      extraPrompt: "",
      baseImage: "",
      resultImages: [],
      promptTrace: {},
    };
  }

  function createWorkflowJob(options) {
    const pages = options.pages.map((page) => createWorkflowPage(page));
    const job = {
      id: crypto.randomUUID(),
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      status: "running",
      userStage: "pages",
      statusText: "正在拆分并准备页面...",
      totalPages: pages.length,
      preparedPages: 0,
      failedPages: 0,
      readyToGeneratePages: 0,
      currentPageNumber: 0,
      documentSummary: options.documentSummary || "",
      splitDiagnostics: options.splitDiagnostics || "",
      referenceDigest: options.referenceDigest || null,
      themeDefinition: options.themeDefinition,
      preferences: options.preferences,
      splitPreset: options.splitPreset || "",
      promptTrace: {
        themeCore: options.themeTrace || null,
        referenceDigest: options.referenceTrace || null,
        splitPlan: options.splitTrace || null,
      },
      pages,
      errors: [],
    };
    workflowJobs.set(job.id, job);
    return job;
  }

  function refreshJobProgress(job) {
    job.pages.forEach((page) => {
      if (!page.prepareDone) return;
      page.onscreenContent = normalizeOnscreenContent(page.onscreenContent);
      page.readyToGenerate = true;
      page.riskLevel = deriveRiskLevel(page);
      page.riskReason = deriveRiskReason(page);
    });
    job.updatedAt = new Date().toISOString();
    job.preparedPages = job.pages.filter((page) => page.prepareDone).length;
    job.failedPages = job.pages.filter((page) => page.generationStatus === "error").length;
    job.readyToGeneratePages = job.pages.filter((page) => page.readyToGenerate).length;
    if (job.preparedPages >= job.totalPages) {
      job.status = "ready";
      job.statusText = job.pages.some((page) => page.riskLevel !== "none")
        ? `页面准备完成，${job.readyToGeneratePages} 页都可以生成，风险页会继续保留提醒。`
        : `页面准备完成，${job.readyToGeneratePages} 页都可以生成。`;
    }
  }

  function publicJobSnapshot(job) {
    return {
      id: job.id,
      status: job.status,
      userStage: job.userStage,
      statusText: job.statusText,
      createdAt: job.createdAt,
      updatedAt: job.updatedAt,
      totalPages: job.totalPages,
      preparedPages: job.preparedPages,
      failedPages: job.failedPages,
      readyToGeneratePages: job.readyToGeneratePages,
      currentPageNumber: job.currentPageNumber,
      documentSummary: job.documentSummary,
      splitDiagnostics: job.splitDiagnostics,
      referenceDigest: job.referenceDigest,
      themeDefinition: job.themeDefinition,
      preferences: job.preferences,
      promptTrace: job.promptTrace,
      pages: job.pages,
      errors: job.errors,
    };
  }

  function deriveRiskLevel(page) {
    if (page.overflowFlag || page.splitRisk === "high" || page.qualityResult.severity === "high") return "high";
    if (page.splitRisk === "medium" || page.qualityResult.severity === "medium" || !page.qualityPass) return "medium";
    return "none";
  }

  function deriveRiskReason(page) {
    if (page.overflowFlag) return page.overflowReason || "上屏内容偏长，建议人工确认。";
    if (!page.qualityPass && page.qualityResult.issues.length) return page.qualityResult.issues[0];
    if (page.splitRisk === "high") return "拆分页时已识别为高风险页。";
    if (page.splitRisk === "medium") return "拆分页时已识别为中风险页。";
    return "";
  }

  async function runThemeDefinition(apiKey, region, themeName, decorationLevel, preferences) {
    const systemPrompt = [
      "你是一位 PPT 视觉系统设计师。",
      "请只返回 JSON object，不要输出 markdown。",
      "字段必须包含 displaySummaryZh、modelPrompt、basic、cover、content、data。",
      "displaySummaryZh 给人看，要中文、清晰、简洁。",
      "modelPrompt 给后续模型编排使用，要专业、精炼、结构化。",
    ].join("\n");

    const userPrompt = [
      "请为一个以 Nano Banana 为最终生图模型的 PPT 生成系统设计全局主题模板。",
      `主题关键词：${themeName}`,
      `装饰强度：${getDecorationLevelLabel(decorationLevel)}`,
      buildPreferencePromptBlock(preferences),
      "要求区分封面页、内容页、数据页的视觉表达。",
      "装饰只允许无字图形，不允许让装饰变成额外文字。",
      "返回 JSON：",
      "{\"displaySummaryZh\":\"...\",\"modelPrompt\":\"...\",\"basic\":\"...\",\"cover\":\"...\",\"content\":\"...\",\"data\":\"...\"}",
    ].join("\n\n");

    const result = await runAssistantJsonObject(apiKey, region, systemPrompt, userPrompt, "主题模板");
    return {
      themeDefinition: normalizeThemeDefinition(result.parsed, themeName, decorationLevel, preferences),
      trace: { systemPrompt, userPrompt, responseText: result.text },
    };
  }

  async function runReferenceDigest(apiKey, region, mainText, referenceFiles, preferences, themeDefinition) {
    if (!referenceFiles.length) {
      return { digest: null, trace: null };
    }

    const systemPrompt = [
      "你是一位 PPT 资料整合助手。",
      "请只返回 JSON object，不要输出 markdown。",
      "字段必须包含 summary、usableFacts、cautions。",
      "summary 用于后续拆分，是一段简洁材料综述。",
      "usableFacts 是可补充给拆分页的事实点列表。",
      "cautions 是需要谨慎使用的信息提醒列表。",
    ].join("\n");

    const userPrompt = [
      buildPreferencePromptBlock(preferences),
      buildThemeDefinitionBlock(themeDefinition),
      "请先理解用户主文本，再把参考材料中真正能补充主文本的内容提炼出来。",
      "用户主文本优先级最高；参考材料只能补充明确存在、且与主题直接相关的信息。",
      "不要把文件里的所有内容都塞进摘要，不要延展到用户没有要求的方向。",
      buildReferenceDigestInput(mainText, referenceFiles),
      "返回 JSON：",
      "{\"summary\":\"...\",\"usableFacts\":[\"...\"],\"cautions\":[\"...\"]}",
    ].join("\n\n");

    const result = await runAssistantJsonObject(apiKey, region, systemPrompt, userPrompt, "参考材料摘要");
    return {
      digest: {
        summary: stringifyStructuredField(result.parsed.summary || ""),
        usableFacts: Array.isArray(result.parsed.usableFacts) ? result.parsed.usableFacts.map((item) => stringifyStructuredField(item)).filter(Boolean) : [],
        cautions: Array.isArray(result.parsed.cautions) ? result.parsed.cautions.map((item) => stringifyStructuredField(item)).filter(Boolean) : [],
      },
      trace: { systemPrompt, userPrompt, responseText: result.text },
    };
  }

  async function runSplitPlan(apiKey, region, options) {
    const { mainText, pageCount, splitPreset, referenceDigest, preferences, themeDefinition, decorationLevel } = options;

    const systemPrompt = [
      "你是一位 PPT 内容策划师。",
      "这一步只做拆分页和逻辑分段，不做最终上屏文案。",
      "请只返回 JSON object，不要输出 markdown。",
      "字段必须包含 documentSummary、splitDiagnostics、pagePlan。",
      "pagePlan 每一项必须包含 pageNumber、pageType、pageTitle、pageContent、sectionTopic、estimatedChars、splitRisk、recommendedBand。",
    ].join("\n");

    const userPrompt = [
      buildHardConstraintBlock(),
      buildPreferencePromptBlock(preferences),
      buildThemeDefinitionBlock({ ...themeDefinition, decorationLevel }),
      "【拆分原则】",
      `目标页数：${pageCount}。`,
      "第 1 页必须是 cover。",
      "pageType 只能是 cover、content、data。",
      "每页只承载一个主题或一个完整思想单元。",
      "每页文字推荐 50-150 字，硬上限 200 字，超过 250 字视为 high 风险。",
      "优先按主题转换、逻辑递进、对比关系切页。",
      "这里不要做最终上屏润色，不要为了美观偷删逻辑。",
      splitPreset ? `【本次拆分模板】\n${splitPreset}` : "",
      referenceDigest?.summary ? `【参考材料摘要】\n${referenceDigest.summary}` : "",
      referenceDigest?.usableFacts?.length ? `【可用补充事实】\n${referenceDigest.usableFacts.join("\n")}` : "",
      referenceDigest?.cautions?.length ? `【参考材料注意事项】\n${referenceDigest.cautions.join("\n")}` : "",
      "【用户主文本】",
      mainText,
      "返回 JSON：",
      "{\"documentSummary\":\"...\",\"splitDiagnostics\":\"...\",\"pagePlan\":[{\"pageNumber\":1,\"pageType\":\"cover\",\"pageTitle\":\"...\",\"pageContent\":\"...\",\"sectionTopic\":\"...\",\"estimatedChars\":90,\"splitRisk\":\"low|medium|high\",\"recommendedBand\":\"minimal|balanced|standard|dense\"}]}",
    ].filter(Boolean).join("\n\n");

    const result = await runAssistantJsonObject(apiKey, region, systemPrompt, userPrompt, "拆分页");
    return {
      pages: normalizePagePlan(result.parsed.pagePlan || result.parsed.pages || [], decorationLevel),
      documentSummary: stringifyStructuredField(result.parsed.documentSummary || result.parsed.summary || ""),
      splitDiagnostics: stringifyStructuredField(result.parsed.splitDiagnostics || ""),
      trace: { systemPrompt, userPrompt, responseText: result.text },
    };
  }

  async function prepareSinglePage(apiKey, region, job, page) {
    const onscreenSystemPrompt = [
      "你是一位 PPT 上屏内容编辑助手。",
      "这一步只决定这一页真正要上屏的内容，不负责页面版式。",
      "请只返回 JSON object，不要输出 markdown。",
      "字段必须包含 onscreenContent、contentBand、overflowFlag、overflowReason、revisionHint。",
      "onscreenContent 输出给用户编辑，只保留最终真正上屏的内容。",
      "onscreenContent 必须是纯上屏文本，不要带 title:、subtitle:、metaInfo:、abstract: 这类字段标签。",
    ].join("\n");

    const onscreenUserPrompt = [
      buildHardConstraintBlock(),
      buildPreferencePromptBlock(job.preferences),
      buildThemeDefinitionBlock(job.themeDefinition),
      "【上屏原则】",
      "最终上屏内容应可直接给用户整体确认与修改。",
      "标题控制在 20 字以内，正文尽量控制在 3-5 个语义块内。",
      "少于 50 字可走极简页；50-150 字走轻量页；150-250 字走标准页；超过建议阈值时可设置 overflowFlag=true 作为提醒。",
      "如果内容过多，不要挤成小字墙；请明确给出 revisionHint，但不要因为超字数就阻止用户继续生成。",
      `页面类型：${page.pageType}`,
      `页面标题：${page.pageTitle}`,
      `页面原始内容：\n${page.pageContent}`,
      "返回 JSON：",
      "{\"onscreenContent\":\"...\",\"contentBand\":\"minimal|balanced|standard|dense\",\"overflowFlag\":false,\"overflowReason\":\"...\",\"revisionHint\":\"...\"}",
    ].join("\n\n");

    const onscreenResult = await runAssistantJsonObject(apiKey, region, onscreenSystemPrompt, onscreenUserPrompt, `第 ${page.pageNumber} 页上屏内容`);
    page.promptTrace.onscreen = {
      systemPrompt: onscreenSystemPrompt,
      userPrompt: onscreenUserPrompt,
      responseText: onscreenResult.text,
    };
    page.onscreenContent = normalizeOnscreenContent(onscreenResult.parsed.onscreenContent || "");
    page.contentBand = normalizeContentBand(onscreenResult.parsed.contentBand, page.onscreenContent);
    page.overflowFlag = Boolean(onscreenResult.parsed.overflowFlag);
    page.overflowReason = stringifyStructuredField(onscreenResult.parsed.overflowReason || "");
    page.revisionHint = stringifyStructuredField(onscreenResult.parsed.revisionHint || "");

    const layoutSystemPrompt = [
      "你是一位 PPT 页面排版与生图提示词设计助手。",
      "这一步只负责页面层级、留白、图文关系和排版策略，不再重新做内容取舍。",
      "请只返回 JSON object，不要输出 markdown。",
      "字段必须包含 layoutSummary、textHierarchy、visualFocus、readabilityNotes、pagePrompt。",
    ].join("\n");

    const layoutUserPrompt = [
      buildHardConstraintBlock(),
      buildPreferencePromptBlock(job.preferences),
      buildThemeDefinitionBlock(job.themeDefinition),
      "【排版原则】",
      "使用 12 列网格，保证约 30% 留白。",
      "全文字体总数不超过 2 种。",
      "页面主标题与页面主正文保持约 1.5-2 倍层级差。",
      "二级和三级标题只做温和层级，不做夸张跳变。",
      "正文默认按不低于 18pt 的可读性风险标准审视。",
      "装饰只允许无字图形、纹理、容器、光效和图标。",
      `页面类型：${page.pageType}`,
      `页面标题：${page.pageTitle}`,
      `上屏内容：\n${page.onscreenContent}`,
      `装饰强度：${getDecorationLevelLabel(page.decorationLevel)}`,
      "返回 JSON：",
      "{\"layoutSummary\":\"...\",\"textHierarchy\":\"...\",\"visualFocus\":\"...\",\"readabilityNotes\":\"...\",\"pagePrompt\":\"...\"}",
    ].join("\n\n");

    const layoutResult = await runAssistantJsonObject(apiKey, region, layoutSystemPrompt, layoutUserPrompt, `第 ${page.pageNumber} 页排版`);
    page.promptTrace.layout = {
      systemPrompt: layoutSystemPrompt,
      userPrompt: layoutUserPrompt,
      responseText: layoutResult.text,
    };
    page.layoutSummary = stringifyStructuredField(layoutResult.parsed.layoutSummary || "");
    page.textHierarchy = stringifyStructuredField(layoutResult.parsed.textHierarchy || "");
    page.visualFocus = stringifyStructuredField(layoutResult.parsed.visualFocus || "");
    page.readabilityNotes = stringifyStructuredField(layoutResult.parsed.readabilityNotes || "");
    page.pagePrompt = stringifyStructuredField(layoutResult.parsed.pagePrompt || "");

    const qualitySystemPrompt = [
      "你是一位 PPT 质量检查助手。",
      "请只返回 JSON object，不要输出 markdown。",
      "字段必须包含 qualityChecklist、passStatus、severity、issues、fixSuggestions、metrics。",
      "metrics 至少包含 charCount、estimatedMinFont、contrastRisk、whitespaceBand、fontFamilyCount。",
    ].join("\n");

    const qualityUserPrompt = [
      buildHardConstraintBlock(),
      buildPreferencePromptBlock(job.preferences),
      buildThemeDefinitionBlock(job.themeDefinition),
      "【质量检查原则】",
      "对比度至少按 4.5:1 风险标准检查。",
      "最小字号不低于 18pt 的可读性标准。",
      "字体总数不超过 2 种。",
      "页面留白约 30%，避免密集小字墙。",
      "检查是否新增未授权文字、是否偏离风格、是否超出安全显示区。",
      "检查跨设备兼容性、投影可读性、图表与引用准确性。",
      `页面类型：${page.pageType}`,
      `页面标题：${page.pageTitle}`,
      `上屏内容：\n${page.onscreenContent}`,
      page.layoutSummary ? `布局摘要：${page.layoutSummary}` : "",
      page.textHierarchy ? `文字层级：${page.textHierarchy}` : "",
      page.visualFocus ? `视觉焦点：${page.visualFocus}` : "",
      page.readabilityNotes ? `可读性说明：${page.readabilityNotes}` : "",
      page.pagePrompt ? `最终生图提示词：\n${page.pagePrompt}` : "",
      "返回 JSON：",
      "{\"qualityChecklist\":\"...\",\"passStatus\":\"pass|review\",\"severity\":\"low|medium|high\",\"issues\":[\"...\"],\"fixSuggestions\":[\"...\"],\"metrics\":{\"charCount\":0,\"estimatedMinFont\":18,\"contrastRisk\":\"...\",\"whitespaceBand\":\"...\",\"fontFamilyCount\":1}}",
    ].filter(Boolean).join("\n\n");

    const qualityResult = await runAssistantJsonObject(apiKey, region, qualitySystemPrompt, qualityUserPrompt, `第 ${page.pageNumber} 页质检`);
    page.promptTrace.quality = {
      systemPrompt: qualitySystemPrompt,
      userPrompt: qualityUserPrompt,
      responseText: qualityResult.text,
    };
    page.qualityResult = {
      pass: String(qualityResult.parsed.passStatus || "").trim().toLowerCase() === "pass",
      severity: ["low", "medium", "high"].includes(String(qualityResult.parsed.severity || "").trim().toLowerCase())
        ? String(qualityResult.parsed.severity || "").trim().toLowerCase()
        : "medium",
      issues: Array.isArray(qualityResult.parsed.issues) ? qualityResult.parsed.issues.map((item) => stringifyStructuredField(item)).filter(Boolean) : [],
      suggestions: Array.isArray(qualityResult.parsed.fixSuggestions) ? qualityResult.parsed.fixSuggestions.map((item) => stringifyStructuredField(item)).filter(Boolean) : [],
      metrics: {
        charCount: Number(qualityResult.parsed.metrics?.charCount || countCharacters(page.onscreenContent)),
        estimatedMinFont: Number(qualityResult.parsed.metrics?.estimatedMinFont || 18),
        contrastRisk: stringifyStructuredField(qualityResult.parsed.metrics?.contrastRisk || "manual-check"),
        whitespaceBand: stringifyStructuredField(qualityResult.parsed.metrics?.whitespaceBand || page.contentBand),
        fontFamilyCount: Number(qualityResult.parsed.metrics?.fontFamilyCount || 1),
      },
      checklist: stringifyStructuredField(qualityResult.parsed.qualityChecklist || ""),
    };
    page.qualityPass = page.qualityResult.pass;
    page.prepareDone = true;
    page.readyToGenerate = true;
    page.riskLevel = deriveRiskLevel(page);
    page.riskReason = deriveRiskReason(page);
  }

  async function prepareWorkflowJob(job, apiKey, region) {
    for (const page of job.pages) {
      job.currentPageNumber = page.pageNumber;
      job.status = "running";
      job.statusText = `正在准备第 ${page.pageNumber}/${job.totalPages} 页...`;
      job.updatedAt = new Date().toISOString();
      try {
        await prepareSinglePage(apiKey, region, job, page);
      } catch (error) {
        page.prepareDone = true;
        page.readyToGenerate = false;
        page.riskLevel = "high";
        page.riskReason = error.message || `第 ${page.pageNumber} 页准备失败。`;
        page.qualityPass = false;
        page.qualityResult = {
          ...buildEmptyQualityResult(),
          issues: [page.riskReason],
          suggestions: ["请先修改这页的上屏内容，再重新整理。"],
        };
        job.errors.push(page.riskReason);
      }
      refreshJobProgress(job);
    }
    job.currentPageNumber = 0;
    refreshJobProgress(job);
  }

  function getWorkflowJobOrThrow(jobId) {
    const job = workflowJobs.get(String(jobId || "").trim());
    if (!job) {
      const error = new Error("找不到对应的工作流任务。");
      error.status = 404;
      throw error;
    }
    return job;
  }

  function getWorkflowPageOrThrow(job, pageId) {
    const page = job.pages.find((item) => item.id === String(pageId || "").trim());
    if (!page) {
      const error = new Error("找不到对应的页面。");
      error.status = 404;
      throw error;
    }
    return page;
  }

  function buildFinalImagePrompt(job, page, extraPrompt = "") {
    const cleanOnscreenContent = normalizeOnscreenContent(page.onscreenContent || page.pageContent);
    return [
      buildHardConstraintBlock(),
      "【主题风格】",
      job.themeDefinition?.modelPrompt || "",
      job.themeDefinition?.[page.pageType === "cover" ? "cover" : page.pageType === "data" ? "data" : "content"] || "",
      "【页面信息】",
      `页面类型：${page.pageType}`,
      `页面标题：${page.pageTitle}`,
      `装饰强度：${getDecorationLevelLabel(page.decorationLevel)}`,
      "【上屏内容】",
      cleanOnscreenContent,
      "【排版意图】",
      page.layoutSummary || "",
      page.textHierarchy ? `文字层级：${page.textHierarchy}` : "",
      page.visualFocus ? `视觉焦点：${page.visualFocus}` : "",
      page.readabilityNotes ? `可读性说明：${page.readabilityNotes}` : "",
      page.pagePrompt ? `页面构图提示：${page.pagePrompt}` : "",
      extraPrompt ? `【额外要求】\n${extraPrompt}` : "",
      "【负向约束】",
      "不要新增页面上没有授权的汉字、英文、数字串、logo、水印、页码、角标、品牌名或解释性小字。",
      "不要把整页做成过满的海报，不要破坏 PPT 页面化结构，不要过度花哨。",
    ].filter(Boolean).join("\n\n");
  }

  app.post("/api/workflow/theme", async (req, res) => {
    const { apiKey, region, themeName, decorationLevel, preferences } = req.body || {};
    if (!apiKey) {
      return res.status(400).json({ code: "MissingApiKey", message: "请先填写 DashScope / Qwen API Key。" });
    }
    if (!String(themeName || "").trim()) {
      return res.status(400).json({ code: "MissingThemeName", message: "请先输入风格主题。" });
    }

    try {
      const normalizedPreferences = normalizePreferences(preferences);
      const result = await runThemeDefinition(
        apiKey,
        region || DEFAULT_REGION,
        String(themeName || "").trim(),
        decorationLevel,
        normalizedPreferences,
      );
      return res.json({
        ok: true,
        themeDefinition: result.themeDefinition,
        promptTrace: { themeCore: result.trace },
      });
    } catch (error) {
      return res.status(500).json({
        code: "WorkflowThemeFailed",
        message: error.message || "生成主题模板失败。",
      });
    }
  });

  app.post("/api/workflow/split", async (req, res) => {
    const {
      apiKey,
      region,
      content,
      pageCount,
      splitTemplate,
      referenceFiles,
      themeDefinition,
      preferences,
      decorationLevel,
    } = req.body || {};

    if (!apiKey) {
      return res.status(400).json({ code: "MissingApiKey", message: "请先填写 DashScope / Qwen API Key。" });
    }

    const mainText = String(content || "").trim();
    if (!mainText) {
      return res.status(400).json({ code: "MissingContent", message: "请先输入需要拆分的主文本。" });
    }

    try {
      const normalizedPreferences = normalizePreferences(preferences);
      const normalizedThemeDefinition = normalizeThemeDefinition(
        themeDefinition || {},
        themeDefinition?.themeName || "",
        decorationLevel,
        normalizedPreferences,
      );
      const normalizedReferenceFiles = normalizeReferenceFiles(referenceFiles)
        .filter((item) => item.includeInSplit && item.parseStatus === "parsed" && (item.extractedText || item.previewText));
      const referenceDigestResult = await runReferenceDigest(
        apiKey,
        region || DEFAULT_REGION,
        mainText,
        normalizedReferenceFiles,
        normalizedPreferences,
        normalizedThemeDefinition,
      );
      const splitResult = await runSplitPlan(apiKey, region || DEFAULT_REGION, {
        mainText,
        pageCount: Math.max(1, Number(pageCount) || 6),
        splitPreset: String(splitTemplate || "").trim(),
        referenceDigest: referenceDigestResult.digest,
        preferences: normalizedPreferences,
        themeDefinition: normalizedThemeDefinition,
        decorationLevel,
      });

      if (!splitResult.pages.length) {
        return res.status(502).json({
          code: "EmptySplitResult",
          message: "拆分页结果为空，请调整文本或拆分模板后重试。",
        });
      }

      const job = createWorkflowJob({
        documentSummary: splitResult.documentSummary || `已拆分为 ${splitResult.pages.length} 页。`,
        splitDiagnostics: splitResult.splitDiagnostics || "本次拆分没有返回额外诊断。",
        referenceDigest: referenceDigestResult.digest,
        themeDefinition: normalizedThemeDefinition,
        preferences: normalizedPreferences,
        splitPreset: String(splitTemplate || "").trim(),
        pages: splitResult.pages,
        themeTrace: null,
        referenceTrace: referenceDigestResult.trace,
        splitTrace: splitResult.trace,
      });

      setTimeout(() => {
        prepareWorkflowJob(job, apiKey, region || DEFAULT_REGION).catch((error) => {
          job.status = "error";
          job.statusText = error.message || "页面准备失败。";
          job.errors.push(job.statusText);
          refreshJobProgress(job);
        });
      }, 0);

      return res.json({
        ok: true,
        jobId: job.id,
        job: publicJobSnapshot(job),
      });
    } catch (error) {
      return res.status(500).json({
        code: "WorkflowSplitFailed",
        message: error.message || "拆分工作流失败。",
      });
    }
  });

  app.get("/api/workflow/jobs/:jobId", (req, res) => {
    try {
      const job = getWorkflowJobOrThrow(req.params.jobId);
      refreshJobProgress(job);
      return res.json({ ok: true, job: publicJobSnapshot(job) });
    } catch (error) {
      return res.status(error.status || 500).json({
        code: "WorkflowJobNotFound",
        message: error.message || "读取工作流任务失败。",
      });
    }
  });

  app.post("/api/workflow/page/reprepare", async (req, res) => {
    const { apiKey, region, jobId, pageId, onscreenContent } = req.body || {};
    if (!apiKey) {
      return res.status(400).json({ code: "MissingApiKey", message: "请先填写 DashScope / Qwen API Key。" });
    }

    try {
      const job = getWorkflowJobOrThrow(jobId);
      const page = getWorkflowPageOrThrow(job, pageId);
      const nextOnscreenContent = String(onscreenContent || "").trim();
      if (!nextOnscreenContent) {
        return res.status(400).json({ code: "MissingOnscreenContent", message: "请先填写这页的上屏内容。" });
      }

      page.onscreenContent = normalizeOnscreenContent(nextOnscreenContent);
      page.contentBand = normalizeContentBand("", page.onscreenContent);
      page.overflowFlag = countCharacters(page.onscreenContent) > 250;
      page.overflowReason = page.overflowFlag ? "当前上屏内容偏长，建议你确认后再决定是否精简。" : "";
      page.revisionHint = page.overflowFlag ? "如果你接受当前密度，可以直接继续生成；若想更清爽，再手动压缩。" : "";
      page.readyToGenerate = true;
      page.generated = false;
      page.promptTrace.userEditedOnscreen = {
        updatedAt: new Date().toISOString(),
        onscreenContent: nextOnscreenContent,
      };

      await prepareSinglePage(apiKey, region || DEFAULT_REGION, job, page);
      refreshJobProgress(job);
      return res.json({
        ok: true,
        job: publicJobSnapshot(job),
        page,
      });
    } catch (error) {
      return res.status(error.status || 500).json({
        code: "WorkflowPageReprepareFailed",
        message: error.message || "重新整理页面失败。",
      });
    }
  });

  app.post("/api/workflow/page/generate-v2", async (req, res) => {
    const {
      apiKey,
      googleApiKey,
      region,
      imageModel,
      jobId,
      pageId,
      slideAspect,
      size,
      seed,
      extraPrompt,
      canvasImage,
    } = req.body || {};

    const selectedImageModel = String(imageModel || WORKFLOW_IMAGE_MODEL).trim() || WORKFLOW_IMAGE_MODEL;
    const useGemini = selectedImageModel === "gemini-3-pro-image-preview";

    if (useGemini && !googleApiKey) {
      return res.status(400).json({ code: "MissingGoogleApiKey", message: "请先填写 Google API Key。" });
    }
    if (!useGemini && !apiKey) {
      return res.status(400).json({ code: "MissingApiKey", message: "请先填写 DashScope / Qwen API Key。" });
    }

    try {
      const job = getWorkflowJobOrThrow(jobId);
      const page = getWorkflowPageOrThrow(job, pageId);
      const finalPrompt = buildFinalImagePrompt(job, page, String(extraPrompt || "").trim());
      page.promptTrace.finalImage = {
        builtAt: new Date().toISOString(),
        model: selectedImageModel,
        prompt: finalPrompt,
        extraPrompt: String(extraPrompt || "").trim(),
        hasCanvasImage: Boolean(canvasImage),
      };

      const parsedCanvasImage = parseDataUrl(String(canvasImage || "").trim());
      let responsePayload = null;
      let images = [];

      if (useGemini) {
        const userParts = [{ text: finalPrompt }];
        if (parsedCanvasImage) {
          userParts.push({
            inlineData: {
              mimeType: parsedCanvasImage.mimeType || "image/png",
              data: parsedCanvasImage.buffer.toString("base64"),
            },
          });
        }

        const geminiBody = {
          contents: [{ role: "user", parts: userParts }],
          generationConfig: {
            responseModalities: ["TEXT", "IMAGE"],
            imageConfig: {
              aspectRatio: normalizeGeminiAspectRatio(slideAspect, size),
              imageSize: normalizeGeminiImageSize(size),
            },
          },
        };
        if (String(seed || "").trim()) {
          geminiBody.generationConfig.seed = Number(seed);
        }

        const parsed = await requestGeminiJson({
          method: "POST",
          url: `${buildGeminiModelUrl(selectedImageModel)}:generateContent`,
          headers: { "x-goog-api-key": googleApiKey },
          body: JSON.stringify(geminiBody),
        });

        if (!parsed.ok) {
          return res.status(parsed.status).json({
            code: "WorkflowPageGenerateFailed",
            message: parsed.data?.error?.message || parsed.data?.message || "生成页面失败。",
            details: parsed.data,
          });
        }

        responsePayload = await normalizeGeminiGenerateResponse(parsed.data, selectedImageModel);
        images = (responsePayload.output?.choices?.[0]?.message?.content || [])
          .filter((item) => item.type === "image" && item.image)
          .map((item) => item.image);
      } else {
        const content = [{ text: finalPrompt }];
        if (parsedCanvasImage) {
          content.push({ image: String(canvasImage || "").trim() });
        }

        const dashscopePayload = {
          model: selectedImageModel,
          input: {
            messages: [
              {
                role: "user",
                content,
              },
            ],
          },
          parameters: {
            size: size || "2K",
            n: 1,
          },
        };

        if (String(seed || "").trim()) {
          dashscopePayload.parameters.seed = Number(seed);
        }

        const response = await fetch(`${resolveRegion(region || DEFAULT_REGION)}/api/v1/services/aigc/multimodal-generation/generation`, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${apiKey}`,
          },
          body: JSON.stringify(dashscopePayload),
        });

        const parsed = await parseJsonResponse(response);
        if (!parsed.ok || parsed.data?.code) {
          return res.status(parsed.status || 500).json({
            code: "WorkflowPageGenerateFailed",
            message: parsed.data?.message || "生成页面失败。",
            details: parsed.data,
          });
        }

        responsePayload = parsed.data;
        images = ((parsed.data.output?.choices || [])[0]?.message?.content || [])
          .filter((item) => item.type === "image" && item.image)
          .map((item) => item.image);
      }

      page.generated = images.length > 0;
      page.generationStatus = images.length > 0 ? "done" : "error";
      page.generationError = images.length > 0 ? "" : "没有拿到图片结果。";
      page.extraPrompt = String(extraPrompt || "").trim();
      page.resultImages = images;
      if (images[0]) page.baseImage = images[0];
      refreshJobProgress(job);

      return res.json({
        ok: true,
        page,
        response: responsePayload,
      });
    } catch (error) {
      return res.status(error.status || 500).json({
        code: "WorkflowPageGenerateFailed",
        message: error.message || "生成页面失败。",
      });
    }
  });

  app.post("/api/workflow/page/generate", async (req, res) => {
    const { googleApiKey, jobId, pageId, slideAspect, size, seed, extraPrompt, canvasImage } = req.body || {};

    if (!googleApiKey) {
      return res.status(400).json({ code: "MissingGoogleApiKey", message: "请先填写 Google API Key。" });
    }

    try {
      const job = getWorkflowJobOrThrow(jobId);
      const page = getWorkflowPageOrThrow(job, pageId);
      const finalPrompt = buildFinalImagePrompt(job, page, String(extraPrompt || "").trim());
      page.promptTrace.finalImage = {
        builtAt: new Date().toISOString(),
        prompt: finalPrompt,
        extraPrompt: String(extraPrompt || "").trim(),
        hasCanvasImage: Boolean(canvasImage),
      };

      const userParts = [{ text: finalPrompt }];
      const parsedCanvasImage = parseDataUrl(String(canvasImage || "").trim());
      if (parsedCanvasImage) {
        userParts.push({
          inlineData: {
            mimeType: parsedCanvasImage.mimeType || "image/png",
            data: parsedCanvasImage.buffer.toString("base64"),
          },
        });
      }

      const geminiBody = {
        contents: [{ role: "user", parts: userParts }],
        generationConfig: {
          responseModalities: ["TEXT", "IMAGE"],
          imageConfig: {
            aspectRatio: normalizeGeminiAspectRatio(slideAspect, size),
            imageSize: normalizeGeminiImageSize(size),
          },
        },
      };
      if (String(seed || "").trim()) {
        geminiBody.generationConfig.seed = Number(seed);
      }

      const parsed = await requestGeminiJson({
        method: "POST",
        url: `${buildGeminiModelUrl(WORKFLOW_IMAGE_MODEL)}:generateContent`,
        headers: { "x-goog-api-key": googleApiKey },
        body: JSON.stringify(geminiBody),
      });

      if (!parsed.ok) {
        return res.status(parsed.status).json({
          code: "WorkflowPageGenerateFailed",
          message: parsed.data?.error?.message || parsed.data?.message || "生成页面失败。",
          details: parsed.data,
        });
      }

      const normalized = await normalizeGeminiGenerateResponse(parsed.data, WORKFLOW_IMAGE_MODEL);
      const images = (normalized.output?.choices?.[0]?.message?.content || [])
        .filter((item) => item.type === "image" && item.image)
        .map((item) => item.image);

      page.generated = images.length > 0;
      page.generationStatus = images.length > 0 ? "done" : "error";
      page.generationError = images.length > 0 ? "" : "没有拿到图片结果。";
      page.extraPrompt = String(extraPrompt || "").trim();
      page.resultImages = images;
      if (images[0]) page.baseImage = images[0];
      refreshJobProgress(job);

      return res.json({
        ok: true,
        page,
        response: normalized,
      });
    } catch (error) {
      return res.status(error.status || 500).json({
        code: "WorkflowPageGenerateFailed",
        message: error.message || "生成页面失败。",
      });
    }
  });
}

module.exports = {
  installWorkflowRoutes,
};
