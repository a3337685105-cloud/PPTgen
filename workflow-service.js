const crypto = require("crypto");

const WORKFLOW_ASSISTANT_MODEL = "qwen3.6-plus";
const WORKFLOW_IMAGE_MODEL = "gemini-3.1-flash-image-preview";
const GEMINI_WORKFLOW_IMAGE_MODELS = new Set([
  "gemini-3.1-flash-image-preview",
  "gemini-3-pro-image-preview",
  "gemini-2.5-flash-image",
]);
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

const AI_PROCESSING_MODES = {
  strict: "原汁原味",
  balanced: "适度润色",
  creative: "深度扩写",
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
    const source = String(input || "")
      .replace(/\r/g, "")
      .replace(/```[\s\S]*?```/g, "")
      .replace(/^\s*(blocks|items|points|entries|sections|visualelements?|datapoints?)\s*[:：].*$/gim, "")
      .replace(/^\s*(index|highlight|type|order|sort|priority)\s*[:：].*$/gim, "")
      .trim();
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

      if (key === "metainfo") {
        flushPending();
        pushLine(lines, content ? `补充信息：${content}` : "");
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

      if (/^[a-z][a-z0-9_]{1,30}$/i.test(key)) {
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

function normalizeContentKey(input) {
  return String(input || "").replace(/[\s_-]+/g, "").toLowerCase();
}

function splitPageContentBuckets(value) {
  const source = stringifyStructuredField(value)
    .replace(/\r/g, "")
    .replace(/```[\s\S]*?```/g, "")
    .trim();
  if (!source) {
    return {
      onscreenContentText: "",
      visualElementsPrompt: "",
      visualElementsDisplay: "",
    };
  }

  const visualKeySet = new Set([
    "visual",
    "visualelement",
    "visualelements",
    "visualhint",
    "visualhints",
    "visualsuggestion",
    "visualsuggestions",
    "designnote",
    "designnotes",
    "layoutnote",
    "layoutnotes",
    "compositionnote",
    "compositionnotes",
    "artdirection",
  ]);
  const wrapperKeySet = new Set(["blocks", "items", "points", "entries", "sections"]);
  const hiddenKeySet = new Set(["type", "highlight", "index", "order", "sort", "priority"]);
  const chineseVisualPrefix = /^(?:视觉元素|视觉建议|画面建议|设计说明|版式说明|构图说明|画面说明|视觉)\s*[:：]\s*(.*)$/i;
  const chineseTextPrefix = /^(?:标题|副标题|补充信息|摘要|正文|内容)\s*[:：]\s*(.*)$/i;
  const visualHeuristics = [
    /(?:页面|本页|画面).{0,8}(?:采用|使用|适合|建议).{0,16}(?:布局|排版|构图|网格|Bento|卡片)/i,
    /(?:左文右图|右文左图|上文下图|上图下文|三栏|四宫格|时间轴|网格|卡片化|分栏|分块|图表化)/i,
    /(?:右下角|左上角|居中突出|黄金分割|背景装饰|作为背景|视觉焦点|主视觉)/i,
    /用.+(?:图标|示意图|结构图|流程图|时间轴|图表|插画|配图)/i,
  ];
  const textLines = [];
  const visualLines = [];

  const pushUnique = (target, text) => {
    const clean = String(text || "").replace(/\s+/g, " ").trim();
    if (!clean) return;
    if (target[target.length - 1] === clean) return;
    target.push(clean);
  };

  source.split(/\r?\n+/).forEach((rawLine) => {
    const line = String(rawLine || "").trim();
    if (!line) return;

    const visualMatch = line.match(chineseVisualPrefix);
    if (visualMatch) {
      pushUnique(visualLines, visualMatch[1] || "");
      return;
    }

    const textMatch = line.match(chineseTextPrefix);
    if (textMatch) {
      pushUnique(textLines, textMatch[1] || "");
      return;
    }

    const match = line.match(/^([A-Za-z_][\w-]*)\s*[:：]\s*(.*)$/);
    if (!match) {
      if (visualHeuristics.some((pattern) => pattern.test(line))) {
        pushUnique(visualLines, line);
        return;
      }
      pushUnique(textLines, line);
      return;
    }

    const key = normalizeContentKey(match[1]);
    const content = String(match[2] || "").trim();
    if (visualKeySet.has(key)) {
      pushUnique(visualLines, content);
      return;
    }
    if (wrapperKeySet.has(key) || hiddenKeySet.has(key)) {
      return;
    }
    if (visualHeuristics.some((pattern) => pattern.test(content || line))) {
      pushUnique(visualLines, content || line);
      return;
    }
    pushUnique(textLines, content || line);
  });

  return {
    onscreenContentText: textLines.join("\n").trim(),
    visualElementsPrompt: visualLines.join("\n").trim(),
    visualElementsDisplay: visualLines.join("\n").trim(),
  };
}

function normalizeVisualElements(value) {
  return splitPageContentBuckets(value).visualElementsPrompt;
}

function normalizeOnscreenContent(value) {
  return splitPageContentBuckets(value).onscreenContentText;
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

  function buildResponsesUrl(region) {
    return `${resolveRegion(region || DEFAULT_REGION)}/compatible-mode/v1/responses`;
  }

  function buildChatCompletionsUrl(region) {
    return `${resolveRegion(region || DEFAULT_REGION)}/compatible-mode/v1/chat/completions`;
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

  function extractChatCompletionText(data) {
    const content = data?.choices?.[0]?.message?.content;
    return typeof content === "string" ? content.trim() : "";
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
            name: "workflow_research_supplements",
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
          message: "联网补充结果修复后仍无法解析为 JSON。",
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

  async function runExpansionResearch(apiKey, region, page, referenceDigest) {
    const prompt = [
      "You are a lightweight web research assistant for PPT slide writing.",
      "Use web search only for the exact topic of this PPT page.",
      "Search the web and propose 0 to 4 factual supplements for expanding one slide.",
      "Every supplement must be directly supported by sources and suitable for slide writing after human review.",
      "Do not rewrite existing user text.",
      "Do not invent organizations, authors, dates, numbers, conclusions, or background stories.",
      "Prefer concise milestone facts, parameter ranges, short definitions, representative applications, or brief comparison points.",
      "Candidate text must be in Simplified Chinese and should stay short.",
      "If no reliable supplement is needed, return an empty candidates array.",
      "Return pure JSON only. No markdown. No prose.",
      "JSON schema:",
      "{\"summary\":\"...\",\"candidates\":[{\"text\":\"...\",\"why\":\"...\",\"sources\":[{\"title\":\"...\",\"url\":\"https://...\"}]}]}",
      "",
      `page_type: ${String(page?.pageType || "content")}`,
      page?.pageTitle ? `page_title: ${String(page.pageTitle).slice(0, 200)}` : "",
      page?.pageContent ? `page_content: ${String(page.pageContent).slice(0, 1600)}` : "",
      referenceDigest?.summary ? `reference_summary: ${String(referenceDigest.summary).slice(0, 1200)}` : "",
      referenceDigest?.usableFacts?.length ? `reference_facts:\n${referenceDigest.usableFacts.slice(0, 12).join("\n")}` : "",
    ].filter(Boolean).join("\n");

    const response = await fetch(buildResponsesUrl(region), {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model: "qwen3.5-plus",
        input: prompt,
        tools: [{ type: "web_search" }],
        enable_thinking: false,
      }),
    });

    const parsed = await parseJsonResponse(response);
    if (!parsed.ok) {
      throw new Error(parsed.data?.message || "联网扩写检索失败。");
    }

    const upstreamError = extractResponsesError(parsed.data);
    if (upstreamError) {
      throw new Error(upstreamError);
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
      if (repaired.ok) {
        parsedJson = repaired.data;
        normalizedRaw = repaired.raw || normalizedRaw;
      }
    }

    const candidates = normalizeResearchCandidates(parsedJson?.candidates);
    return {
      summary: String(parsedJson?.summary || "").trim(),
      candidates,
      raw: normalizedRaw,
    };
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

function normalizeAiProcessingMode(value) {
  const normalized = String(value || "").trim().toLowerCase();
  return ["strict", "balanced", "creative"].includes(normalized) ? normalized : "balanced";
}

function getAiProcessingModeLabel(value) {
  return AI_PROCESSING_MODES[normalizeAiProcessingMode(value)];
}

  function buildPreferencePromptBlock(preferences) {
    return [
      "【用户偏好】",
      `风格定位：${PREFERENCE_LABELS.styleMode[preferences.styleMode]}。`,
      `版式节奏：${PREFERENCE_LABELS.layoutVariety[preferences.layoutVariety]}。`,
      `细节密度：${PREFERENCE_LABELS.detailLevel[preferences.detailLevel]}。`,
      `版面疏密：${PREFERENCE_LABELS.visualDensity[preferences.visualDensity]}。`,
      `图文重心：${PREFERENCE_LABELS.compositionFocus[preferences.compositionFocus]}。`,
      `数据表达：${PREFERENCE_LABELS.dataNarrative[preferences.dataNarrative]}。`,
      `整体氛围：${PREFERENCE_LABELS.pageMood[preferences.pageMood]}。`,
    ].join("\n");
  }

  function buildThemeDefinitionBlock(themeDefinition) {
    if (!themeDefinition) return "";
    return [
      "【全局主题模板】",
      themeDefinition.modelPrompt ? `模型总纲：${themeDefinition.modelPrompt}` : "",
      themeDefinition.basic ? `Basic：${themeDefinition.basic}` : "",
      themeDefinition.cover ? `Cover：${themeDefinition.cover}` : "",
      themeDefinition.catalog ? `Catalog：${themeDefinition.catalog}` : "",
      themeDefinition.chapter ? `Chapter：${themeDefinition.chapter}` : "",
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
      catalog: stringifyStructuredField(result?.catalog || ""),
      chapter: stringifyStructuredField(result?.chapter || ""),
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
        normalized.catalog ? `目录页：${normalized.catalog}` : "",
        normalized.chapter ? `章节页：${normalized.chapter}` : "",
        normalized.content ? `内容页：${normalized.content}` : "",
      ].filter(Boolean).join("\n");
    }
    if (!normalized.modelPrompt) {
      normalized.modelPrompt = [normalized.basic, normalized.cover, normalized.catalog, normalized.chapter, normalized.content, normalized.data].filter(Boolean).join("\n");
    }
    return normalized;
  }

  function normalizePageType(value, index) {
    const normalized = String(value || "").trim().toLowerCase();
    if (["cover", "catalog", "chapter", "content", "data"].includes(normalized)) return normalized;
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
      pass: true,
      severity: "low",
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
      onscreenContentText: "",
      visualElementsPrompt: "",
      visualElementsDisplay: "",
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

  function preparePageForGeneration(job, page, source = "split") {
    const separatedContent = splitPageContentBuckets(page.onscreenContent || page.onscreenContentText || page.pageContent || page.pageTitle || "");
    const cleanOnscreenContent = separatedContent.onscreenContentText;
    const visualElementsPrompt = separatedContent.visualElementsPrompt;
    const charCount = countCharacters(cleanOnscreenContent);

    page.onscreenContent = cleanOnscreenContent;
    page.onscreenContentText = cleanOnscreenContent;
    page.visualElementsPrompt = visualElementsPrompt;
    page.visualElementsDisplay = separatedContent.visualElementsDisplay;
    page.contentBand = normalizeContentBand(page.contentBand || page.recommendedBand, cleanOnscreenContent);
    page.overflowFlag = charCount > 250;
    page.overflowReason = page.overflowFlag
      ? `当前页内容约 ${charCount} 字（按中英数字加权估算），建议关注页面密度。`
      : "";
    page.revisionHint = page.overflowFlag
      ? "如果你接受当前信息密度，可以直接生成；如果想更清爽，再手动精简。"
      : "";
    page.layoutSummary = "";
    page.textHierarchy = "";
    page.visualFocus = "";
    page.readabilityNotes = "";
    page.pagePrompt = "";
    page.qualityResult = {
      ...buildEmptyQualityResult(),
      metrics: {
        charCount,
        estimatedMinFont: 18,
        contrastRisk: "disabled",
        whitespaceBand: page.contentBand,
        fontFamilyCount: 1,
      },
      checklist: "已停用自动质检，当前只保留排版风险提醒。",
    };
    page.qualityPass = true;
    page.prepareDone = true;
    page.readyToGenerate = true;
    page.promptTrace.simplifiedPrepare = {
      mode: "local",
      source,
      strategy: "normalize-onscreen-only",
      imagePromptStrategy: "basic + pageTypeTemplate + languageInstruction + onscreenContent",
      quality: "disabled",
      layout: "disabled",
      onscreenContent: cleanOnscreenContent,
      visualElementsPrompt,
    };
    page.riskLevel = deriveRiskLevel(page);
    page.riskReason = deriveRiskReason(page);
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
      aiProcessingMode: normalizeAiProcessingMode(options.aiProcessingMode),
      splitPreset: options.splitPreset || "",
        promptTrace: {
          themeCore: options.themeTrace || null,
          referenceDigest: options.referenceTrace || null,
          splitPlan: options.splitTrace || null,
          expansionPlan: options.expansionTrace || null,
          lengthControlPlan: options.expansionTrace || null,
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
      page.onscreenContentText = page.onscreenContent;
      page.visualElementsPrompt = normalizeVisualElements(page.visualElementsPrompt || page.pageContent || "");
      page.visualElementsDisplay = page.visualElementsPrompt;
      if (page.generationStatus === "running") {
        page.generationStatus = page.generated ? "done" : "idle";
      }
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
    const publicPages = job.pages.map((page) => ({
      ...page,
      baseImage: String(page.baseImage || "").startsWith("data:") ? "" : page.baseImage,
      resultImages: Array.isArray(page.resultImages)
        ? page.resultImages.filter((item) => !String(item || "").startsWith("data:"))
        : [],
    }));
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
      aiProcessingMode: job.aiProcessingMode,
      promptTrace: job.promptTrace,
      pages: publicPages,
      errors: job.errors,
    };
  }

  function deriveRiskLevel(page) {
    if (page.overflowFlag || page.splitRisk === "high") return "high";
    if (page.splitRisk === "medium") return "medium";
    return "none";
  }

  function deriveRiskReason(page) {
    if (page.overflowFlag) return page.overflowReason || "当前页内容偏长，建议关注页面密度。";
    if (page.splitRisk === "high") return "拆分页时识别为高排版风险页。";
    if (page.splitRisk === "medium") return "拆分页时识别为中排版风险页。";
    return "";
  }

  async function runThemeDefinition(apiKey, region, themeName, decorationLevel, preferences) {
    const systemPrompt = [
      "你是一位专注 PPT 视觉表达的高级艺术总监及 UI/UX 专家。",
"你的任务是为顶级生图大模型 Nano Banana 2 撰写极具画面感、物理质感的【中文】系统级提示词。",
      "请只返回 JSON object，不要输出 markdown。",
      "字段必须包含 displaySummaryZh、modelPrompt、basic、cover、catalog、chapter、content、data。",
    ].join("\n");

    const effectiveSystemPrompt = [
      "你是一位精通 PPT 视觉设计、演示文稿美化与版式统筹的高级设计师。",
      "你的任务是为顶级生图大模型 Nano Banana 2 撰写极具画面感、物理质感、并且严格服务于 PPT 场景的【中文】系统级提示词。",
      "请只返回 JSON object，不要输出 markdown。",
      "字段必须包含 displaySummaryZh、modelPrompt、basic、cover、catalog、chapter、content、data。",
    ].join("\n");

    const userPrompt = [
`请为“${themeName}”设计全局主题模板。目标生图模型：Nano Banana 2。`,
      `装饰强度：${getDecorationLevelLabel(decorationLevel)}`,
      buildPreferencePromptBlock(preferences),
"【Nano Banana 2 模型专属特性要求】",
      "1. 画面感极值：该模型对【材质】、【光影】和【渲染引擎】极其敏感。必须大量使用如“Octane Render渲染”、“虚幻引擎5”、“电影级光影”、“丁达尔效应”、“表面微磨砂”、“SSS次表面散射”等顶级CG行业术语。",
      "2. 绝不用抽象词：该模型不懂“好看”、“干净”等抽象词汇。必须转化为物理描述，如“背景为无反光的摄影棚纯色背景”、“光线为顶部柔和漫反射”。",
      "3. 负向约束明确：该模型容易“画蛇添足”，必须在提示词中强硬规定不准出现的内容（如：“绝不允许出现多余的UI元素、发光的赛博朋克线条或杂乱的3D碎屑”）。",
      "【极其重要的输出格式与语言要求】",
      "1. basic, cover, catalog, chapter, content, data 字段的输出【必须全部使用极其专业、富有画面感的中文】，绝不允许使用十六进制色号（如 #FFFFFF），英文仅限专业渲染术语（如 Octane Render）。",
      "2. 颜色必须使用自然界或工业界的名词（如：钛银色、深海绿、实验室纯白）。",
      "3. basic 字段的开头，必须使用角色扮演句式设定世界观。例如：“你是一位专注于 [某领域] 的平面设计视觉专家，请生成一份专业、高精度的16:9演示文稿。视觉风格融合...”",
      "4. 各 PageType 字段，必须采用结构化的书写方式。必须包含：“构图逻辑：...”、“视觉焦点：...”、“负向约束：...”。",
      "【排版与视觉原则（动态响应用户偏好）】",
      "1. 秩序与层级：要求使用模块化网格系统，明确标题与正文呈现极端的视觉落差（如 1:0.618）。",
      "2. 具象化的插图（关键！）：不要只谈排版。必须在各模块中，为模型提供具体的【物理插图隐喻】。必须让模型知道“具体画什么物体”。",
      "3. 页面类型指南（根据偏好动态调整，以下为基准建议）：",
      "   - cover: 确立唯一视觉焦点（如海报模式）。若偏好【视觉主导】，则放大3D主视觉物体；若偏好【内容主导/简约】，则克制图形，让超大标题占据重心。",
      "   - catalog: 建立显性的纵向或横向阵列排布。依据【版式变化】偏好决定是规整列表还是创意散点阵列。",
      "   - chapter: 色块或材质的巨大分割。利用巨大的序号作为实体水印或空间重塑的锚点。",
      "   - content: 确立信息容器（如 Bento 便当盒或双栏网格）。若偏好【留白更多】，则采用极简物理底板；若偏好【信息更满】，则引入紧凑的卡片堆叠或仪表盘化背景。",
      "   - data: 图表高精物理模型化。若偏好【适度信息图/清晰克制】，则如同无尘实验台上的实体；若偏好【更强视觉化】，可加入全息、发光光纤等具象化数据流表达。",
      "返回 JSON：",
      "{\"displaySummaryZh\":\"...\",\"modelPrompt\":\"...\",\"basic\":\"(必须输出纯中文描述)\",\"cover\":\"(必须输出纯中文描述)\",\"catalog\":\"(必须输出纯中文描述)\",\"chapter\":\"(必须输出纯中文描述)\",\"content\":\"(必须输出纯中文描述)\",\"data\":\"(必须输出纯中文描述)\"}",
    ].join("\n\n");

    const effectiveUserPrompt = [
      `请为“${themeName}”设计全局主题模板。目标生图模型：Nano Banana 2。`,
      `装饰强度：${getDecorationLevelLabel(decorationLevel)}`,
      "【最高优先级】",
      "必须严格继承用户提供的主题描述、配色倾向、字体气质、材质隐喻和风格问卷；这些信息要贯穿 displaySummaryZh、modelPrompt、basic、cover、catalog、chapter、content、data 全部字段。",
      "如果用户已经明确指定了风格方向，不要被通用模板覆盖，不要改写成默认审美。",
      buildPreferencePromptBlock(preferences),
      "【Nano Banana 2 表达要求】",
      "1. 语言必须具体、可视、物理化，优先描述材质、光影、空间、镜头、渲染质感。",
      "2. 避免抽象空话；把“高级、干净、好看”改写成可见画面描述。",
      "3. 必须明确负向约束，禁止出现网页、UI 面板、仪表盘感、按钮控件、无关装饰和杂乱碎屑。",
      "【输出要求】",
      "1. 所有字段必须使用专业中文；不要使用十六进制色号；颜色请用自然、工业或材料命名。",
      "2. basic 负责定义整套 PPT 的视觉世界观、版式秩序、字体层级、材质和气质，不要写成网页或 UI 设计说明。",
      "3. cover、catalog、chapter、content、data 必须分别针对页型给出结构化提示，至少包含：构图逻辑、视觉焦点、信息容器或版式策略、负向约束。",
      "【排版与视觉原则】",
      "1. 使用模块化网格系统，标题与正文保持明确层级，字号比例按 1:0.618 控制。",
      "2. 不要只谈排版；每个页型都要提供具体可画的物体、装置、结构或物理插图隐喻。",
      "3. cover 强调单一焦点；catalog 强调阵列与浏览节奏；chapter 强调章节切割感；content 强调信息容器；data 强调数据图表的视觉化承载。",
      "返回 JSON：",
      "{\"displaySummaryZh\":\"...\",\"modelPrompt\":\"...\",\"basic\":\"...\",\"cover\":\"...\",\"catalog\":\"...\",\"chapter\":\"...\",\"content\":\"...\",\"data\":\"...\"}",
    ].join("\n\n");

    const result = await runAssistantJsonObject(apiKey, region, effectiveSystemPrompt, effectiveUserPrompt, "主题模板");
    return {
      themeDefinition: normalizeThemeDefinition(result.parsed, themeName, decorationLevel, preferences),
      trace: { systemPrompt: effectiveSystemPrompt, userPrompt: effectiveUserPrompt, responseText: result.text },
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
    const { mainText, pageCount, splitPreset, referenceDigest, preferences, themeDefinition, decorationLevel, aiProcessingMode } = options;

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
      `【AI 处理模式】\n${getAiProcessingModeLabel(aiProcessingMode)}`,
      "【拆分原则】",
      `目标页数：${pageCount}。`,
      "第 1 页必须是 cover。",
      "pageType 只能是 cover、catalog、chapter、content、data。",
      "需要目录请用 catalog，需要章节过渡请用 chapter。",
      "每页只承载一个主题或一个完整思想单元。",
      "每页文字推荐 50-150 字，硬上限 200 字，超过 250 字视为 high 风险。",
      "优先按主题转换、逻辑递进、对比关系切页。",
      "这里不要做最终上屏润色，不要为了美观偷删逻辑。",
      aiProcessingMode === "strict"
        ? "Strict 模式：禁止扩写、禁止缩写、禁止改写语气。尽量保留用户原句，只做硬性分页、标题抽取和类型标记。"
        : "",
      aiProcessingMode === "balanced"
        ? "Balanced 模式：允许轻度梳理逻辑、提炼小标题，但不要过度总结，也不要随意补充新事实。"
        : "",
      aiProcessingMode === "creative"
        ? "Creative 模式：可以在不违背主文本主题的前提下，适度补全桥接句、案例提示和说明语气，让单页更完整。若参考材料明确支持，可引入其中的事实补充。"
        : "",
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

  async function runPageExpansion(apiKey, region, options) {
    const { enableExpansion, targetChars, maxChars, pages, referenceDigest } = options;
    const normalizedMaxChars = Math.max(0, Number(maxChars) || 0);
    const normalizedTargetChars = Boolean(enableExpansion)
      ? Math.max(0, Math.min(Number(targetChars) || 0, normalizedMaxChars || Number(targetChars) || 0))
      : 0;
    const candidates = (Array.isArray(pages) ? pages : [])
      .map((page) => {
        const pageType = String(page.pageType || "").toLowerCase();
        const currentChars = countCharacters(page.pageContent || "");
        const shouldExpand =
          normalizedTargetChars > 0 &&
          ["content", "data"].includes(pageType) &&
          currentChars < normalizedTargetChars;
        const shouldCondense =
          normalizedMaxChars > 0 &&
          ["catalog", "content", "data"].includes(pageType) &&
          currentChars > normalizedMaxChars;
        if (!shouldExpand && !shouldCondense) return null;
        return {
          ...page,
          currentChars,
          lengthMode: shouldCondense ? "condense" : "expand",
          desiredChars: shouldCondense ? normalizedMaxChars : normalizedTargetChars,
        };
      })
      .filter(Boolean);

    if (!candidates.length) {
      return {
        pages,
        trace: null,
      };
    }

    const researchByPageNumber = new Map();
    const researchTrace = [];
    if (normalizedTargetChars > 0) {
      for (const page of candidates.filter((item) => item.lengthMode === "expand")) {
        try {
          const research = await runExpansionResearch(apiKey, region, page, referenceDigest);
          researchByPageNumber.set(Number(page.pageNumber || 0), research);
          researchTrace.push({
            pageNumber: Number(page.pageNumber || 0),
            pageTitle: page.pageTitle,
            summary: research.summary,
            candidateCount: research.candidates.length,
            candidates: research.candidates.map((item) => ({
              text: item.text,
              sources: item.sources.map((source) => source.title),
            })),
          });
        } catch (error) {
          researchTrace.push({
            pageNumber: Number(page.pageNumber || 0),
            pageTitle: page.pageTitle,
            candidateCount: 0,
            error: error.message || "联网扩写检索失败。",
          });
        }
      }
    }

    const systemPrompt = [
      "你是一位 PPT 单页内容长度调节助手。",
      "你要根据每页的长度指示，决定是扩写还是压缩。",
      "不要改变 pageNumber、pageType、pageTitle、sectionTopic。",
      "不得编造没有依据的新事实。",
      "不要生成 visualElements、designNotes、layoutNotes 或任何视觉字段。",
      "如果是 expand：只允许使用当前页面原文、参考摘要、参考事实，以及联网检索得到的已验证补充事实来扩写。",
      "如果联网检索没有给出可靠补充，请宁可保持略短，也不要为了凑字数乱写。",
      "如果是 condense：保留关键信息、比较关系和结论，删去重复、冗长说明，必须压到不超过目标字数。",
      "只返回 JSON object。",
    ].join("\n");

    const userPrompt = [
      normalizedTargetChars ? `【扩写目标字数】\n${normalizedTargetChars}` : "",
      normalizedMaxChars ? `【压缩上限字数】\n${normalizedMaxChars}` : "",
      referenceDigest?.summary ? `【参考摘要】\n${referenceDigest.summary}` : "",
      referenceDigest?.usableFacts?.length ? `【可用客观事实】\n${referenceDigest.usableFacts.join("\n")}` : "",
      "【待调整页面】",
      JSON.stringify(candidates.map((page) => {
        const research = researchByPageNumber.get(Number(page.pageNumber || 0));
        return {
          pageNumber: page.pageNumber,
          pageType: page.pageType,
          pageTitle: page.pageTitle,
          sectionTopic: page.sectionTopic,
          pageContent: page.pageContent,
          currentChars: page.currentChars,
          lengthMode: page.lengthMode,
          desiredChars: page.desiredChars,
          estimatedChars: page.estimatedChars,
          researchedSummary: research?.summary || "",
          researchedSupplements: Array.isArray(research?.candidates)
            ? research.candidates.map((item) => ({
              text: item.text,
              sourceTitles: item.sources.map((source) => source.title),
            }))
            : [],
        };
      }), null, 2),
      "返回 JSON：",
      "{\"pages\":[{\"pageNumber\":1,\"lengthMode\":\"expand|condense\",\"pageContent\":\"...\",\"estimatedChars\":140}]}",
    ].filter(Boolean).join("\n\n");

    const result = await runAssistantJsonObject(apiKey, region, systemPrompt, userPrompt, "单页长度调节");
    const expandedPages = Array.isArray(result.parsed.pages) ? result.parsed.pages : [];
    const expandedByNumber = new Map(
      expandedPages
        .map((item) => ({
          pageNumber: Number(item?.pageNumber || 0),
          pageContent: stringifyStructuredField(item?.pageContent || ""),
          estimatedChars: Number(item?.estimatedChars || 0),
        }))
        .filter((item) => item.pageNumber > 0 && item.pageContent)
        .map((item) => [item.pageNumber, item])
    );

    const mergedPages = pages.map((page) => {
      const expanded = expandedByNumber.get(Number(page.pageNumber || 0));
      if (!expanded) return page;
      return {
        ...page,
        pageContent: expanded.pageContent,
        estimatedChars: expanded.estimatedChars || countCharacters(expanded.pageContent),
      };
    });

    return {
      pages: mergedPages,
      trace: {
        systemPrompt,
        userPrompt,
        responseText: result.text,
        targetChars: normalizedTargetChars,
        maxChars: normalizedMaxChars,
        enableExpansion: Boolean(enableExpansion),
        research: researchTrace,
      },
    };
  }

  async function prepareSinglePage(apiKey, region, job, page) {
    preparePageForGeneration(job, page, "local-reprepare");
  }

  async function prepareWorkflowJob(job, apiKey, region, options = {}) {
    const {
      enableExpansion = false,
      targetChars = 0,
      maxChars = 0,
      referenceDigest = null,
    } = options;

    for (const page of job.pages) {
      job.currentPageNumber = page.pageNumber;
      job.status = "running";
      job.statusText = `正在整理第 ${page.pageNumber}/${job.totalPages} 页...`;
      job.updatedAt = new Date().toISOString();
      try {
        const lengthResult = await runPageExpansion(apiKey, region, {
          enableExpansion,
          targetChars,
          maxChars,
          pages: [page],
          referenceDigest,
        });
        const preparedPage = Array.isArray(lengthResult.pages) ? lengthResult.pages[0] : null;
        if (preparedPage) {
          Object.assign(page, preparedPage);
        }
        if (lengthResult.trace) {
          page.promptTrace.lengthControl = lengthResult.trace;
        }
        preparePageForGeneration(job, page, "local-batch");
      } catch (error) {
        page.prepareDone = true;
        page.readyToGenerate = true;
        page.riskLevel = "high";
        page.riskReason = error.message || `第 ${page.pageNumber} 页准备失败。`;
        page.qualityPass = true;
        page.qualityResult = {
          ...buildEmptyQualityResult(),
          issues: [page.riskReason],
          suggestions: ["你可以直接生成，也可以先修改当前页内容。"],
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

  function buildPptLayoutPrinciplesBlock() {
    return [
      "【PPT排版原则｜高优先级】",
      "对比：标题、结论、关键数字必须显著大于正文；标题与正文字号比例严格为 1:0.618。",
      "对比：标题、结论、关键数字必须显著大于正文；标题与正文字号比例严格为 1:0.618。",
      "亲密：相关内容靠近成组，无关内容主动拉开，避免信息粘连。",
      "对齐：文字、图形、卡片、数据区遵循统一网格与明确边线，保持秩序感。",
      "重复：同层级元素保持一致的字体、字重、颜色、间距与容器样式。",
      "留白：保留呼吸感，内容过多时优先分栏、分块、卡片化或图表化，不压成小字墙。",
      "可读：正文适合投影阅读，禁止细字、浅字、低对比文字压在复杂背景上。",
    ].join("\n");
  }

  function buildFinalImagePrompt(job, page, extraPrompt = "") {
    const cleanOnscreenContent = normalizeOnscreenContent(page.onscreenContentText || page.onscreenContent || page.pageContent);
    const visualElementsPrompt = normalizeVisualElements(page.visualElementsPrompt || page.pageContent || "");
    const pageTypeKey = ["cover", "catalog", "chapter", "content", "data"].includes(page.pageType) ? page.pageType : "content";
    const pageTypeTemplate = job.themeDefinition?.[pageTypeKey] || "";
    const pptLayoutPrinciples = buildPptLayoutPrinciplesBlock();
    return [
      job.themeDefinition?.basic || "",
      pptLayoutPrinciples,
      pageTypeTemplate,
      `本页标题：${page.pageTitle}`,
      cleanOnscreenContent ? `以下是我的文字内容：\n${cleanOnscreenContent}` : "",
      visualElementsPrompt ? `画面补充建议：\n${visualElementsPrompt}` : "",
      extraPrompt ? `补充要求：\n${normalizeOnscreenContent(extraPrompt)}` : "",
    ].filter(Boolean).join("\n\n");
  }

  function extractGeminiSearchMetadata(data) {
    const metadata = [];
    const candidates = Array.isArray(data?.candidates) ? data.candidates : [];
    candidates.forEach((candidate) => {
      if (candidate?.groundingMetadata) metadata.push(candidate.groundingMetadata);
      if (candidate?.citationMetadata) metadata.push(candidate.citationMetadata);
    });
    return metadata.length ? metadata : null;
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
      aiProcessingMode,
      enableExpansion,
      targetChars,
      maxChars,
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
        aiProcessingMode: normalizeAiProcessingMode(aiProcessingMode),
        referenceDigest: referenceDigestResult.digest,
        preferences: normalizedPreferences,
        themeDefinition: normalizedThemeDefinition,
        decorationLevel,
      });
      const normalizedEnableExpansion = Boolean(enableExpansion);
      const normalizedTargetChars = Math.max(0, Number(targetChars) || 0);
      const normalizedMaxChars = Math.max(0, Number(maxChars) || 0);

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
        aiProcessingMode: normalizeAiProcessingMode(aiProcessingMode),
        pages: splitResult.pages,
        themeTrace: null,
        referenceTrace: referenceDigestResult.trace,
        splitTrace: splitResult.trace,
        expansionTrace: {
          mode: "background-page-by-page",
          enableExpansion: normalizedEnableExpansion,
          targetChars: normalizedTargetChars,
          maxChars: normalizedMaxChars,
        },
      });
      job.status = "running";
      job.statusText = "拆分完成，正在逐页准备结果...";
      refreshJobProgress(job);

      setTimeout(() => {
        prepareWorkflowJob(job, apiKey, region || DEFAULT_REGION, {
          enableExpansion: normalizedEnableExpansion,
          targetChars: normalizedTargetChars,
          maxChars: normalizedMaxChars,
          referenceDigest: referenceDigestResult.digest,
        }).catch((error) => {
          job.status = "error";
          job.statusText = error.message || "逐页准备失败。";
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
    const { apiKey, region, jobId, pageId, onscreenContent, autoExpandToMaxChars } = req.body || {};
    try {
      const job = getWorkflowJobOrThrow(jobId);
      const page = getWorkflowPageOrThrow(job, pageId);
      const nextOnscreenContent = String(onscreenContent || "").trim();
      if (!nextOnscreenContent) {
        return res.status(400).json({ code: "MissingOnscreenContent", message: "请先填写这页的上屏内容。" });
      }

      page.onscreenContent = normalizeOnscreenContent(nextOnscreenContent);
      page.onscreenContentText = page.onscreenContent;
      page.pageContent = page.onscreenContent;
      page.generated = false;
      page.promptTrace.userEditedOnscreen = {
        updatedAt: new Date().toISOString(),
        onscreenContent: nextOnscreenContent,
      };

      if (autoExpandToMaxChars) {
        const lengthPlan = job.promptTrace?.lengthControlPlan || job.promptTrace?.expansionPlan || null;
        const normalizedMaxChars = Math.max(0, Number(lengthPlan?.maxChars) || 0);
        if (!normalizedMaxChars) {
          return res.status(400).json({
            code: "MissingMaxChars",
            message: "当前工作流没有可用的最大字数设置，无法执行 AI 一键重润。",
          });
        }
        if (!String(apiKey || "").trim()) {
          return res.status(400).json({
            code: "MissingApiKey",
            message: "请先填写 API Key，再执行 AI 一键重润。",
          });
        }
        const lengthResult = await runPageExpansion(apiKey, region || DEFAULT_REGION, {
          enableExpansion: true,
          targetChars: normalizedMaxChars,
          maxChars: normalizedMaxChars,
          pages: [{
            ...page,
            pageContent: page.onscreenContent,
          }],
          referenceDigest: job.referenceDigest || null,
        });
        const repolishedPage = Array.isArray(lengthResult.pages) ? lengthResult.pages[0] : null;
        if (repolishedPage) {
          Object.assign(page, repolishedPage);
          page.onscreenContent = normalizeOnscreenContent(page.pageContent || page.onscreenContent);
          page.onscreenContentText = page.onscreenContent;
        }
        if (lengthResult.trace) {
          page.promptTrace.lengthControl = lengthResult.trace;
          page.promptTrace.aiRepolish = {
            updatedAt: new Date().toISOString(),
            mode: "expand-to-max-chars",
            maxChars: normalizedMaxChars,
          };
        }
      }

      preparePageForGeneration(job, page, "user-edit");
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
      onscreenContent,
      enableGeminiGoogleSearch,
    } = req.body || {};

    const selectedImageModel = String(imageModel || WORKFLOW_IMAGE_MODEL).trim() || WORKFLOW_IMAGE_MODEL;
    const useGemini = GEMINI_WORKFLOW_IMAGE_MODELS.has(selectedImageModel);

    if (useGemini && !googleApiKey) {
      return res.status(400).json({ code: "MissingGoogleApiKey", message: "请先填写 Google API Key。" });
    }
    if (!useGemini && !apiKey) {
      return res.status(400).json({ code: "MissingApiKey", message: "请先填写 DashScope / Qwen API Key。" });
    }

    try {
      const job = getWorkflowJobOrThrow(jobId);
      const page = getWorkflowPageOrThrow(job, pageId);
      if (String(onscreenContent || "").trim()) {
        page.onscreenContent = normalizeOnscreenContent(String(onscreenContent || "").trim());
        page.onscreenContentText = page.onscreenContent;
        preparePageForGeneration(job, page, "generate-inline");
      }
      const finalPrompt = buildFinalImagePrompt(job, page, String(extraPrompt || "").trim());
      page.promptTrace.finalImage = {
        builtAt: new Date().toISOString(),
        model: selectedImageModel,
        prompt: finalPrompt,
        extraPrompt: String(extraPrompt || "").trim(),
        hasCanvasImage: Boolean(canvasImage),
        googleSearchEnabled: Boolean(enableGeminiGoogleSearch && useGemini),
        searchMetadata: null,
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
        if (enableGeminiGoogleSearch) {
          geminiBody.tools = [{ google_search: {} }];
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
        page.promptTrace.finalImage.searchMetadata = extractGeminiSearchMetadata(parsed.data);
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

      page.generationStatus = images.length > 0 ? "done" : "error";
      page.generationError = images.length > 0 ? "" : "没有拿到图片结果。";
      page.extraPrompt = String(extraPrompt || "").trim();
      const previousImages = Array.isArray(page.resultImages) ? page.resultImages.filter(Boolean) : [];
      const mergedImages = Array.from(new Set([
        ...images.filter(Boolean),
        ...previousImages,
      ]));
      page.generated = mergedImages.length > 0;
      page.resultImages = mergedImages;
      if (images[0]) page.baseImage = images[0];
      if (!page.baseImage && mergedImages[0]) page.baseImage = mergedImages[0];
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

      page.generationStatus = images.length > 0 ? "done" : "error";
      page.generationError = images.length > 0 ? "" : "没有拿到图片结果。";
      page.extraPrompt = String(extraPrompt || "").trim();
      const previousImages = Array.isArray(page.resultImages) ? page.resultImages.filter(Boolean) : [];
      const mergedImages = Array.from(new Set([
        ...images.filter(Boolean),
        ...previousImages,
      ]));
      page.generated = mergedImages.length > 0;
      page.resultImages = mergedImages;
      if (images[0]) page.baseImage = images[0];
      if (!page.baseImage && mergedImages[0]) page.baseImage = mergedImages[0];
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
