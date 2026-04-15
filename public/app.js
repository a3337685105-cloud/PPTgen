const STORAGE_KEY = "ppt-image-studio-settings";
const MAX_BOXES_PER_IMAGE = 2;
const THEME_WORKFLOW_VERSION = 3;
const LIBRARY_DOC_VERSION = 1;
const WORKFLOW_PLAN_LIBRARY_LIMIT = 24;
const DEFAULT_REGION = "beijing";
const GEMINI_IMAGE_MODELS = new Set(["gemini-3-pro-image-preview"]);

const SLIDE_ASPECTS = {
  "16:9": { label: "16:9", canvasWidth: 800, canvasHeight: 450, outputWidth: 3840, outputHeight: 2160 },
  "4:3": { label: "4:3", canvasWidth: 800, canvasHeight: 600, outputWidth: 3840, outputHeight: 2880 },
  "1:1": { label: "1:1", canvasWidth: 700, canvasHeight: 700, outputWidth: 3840, outputHeight: 3840 },
};

const TEMPLATE_BUILDERS = {
  blank: () => [],
  cover: () => [
    { label: "标题区", x: 0.07, y: 0.1, w: 0.5, h: 0.18 },
    { label: "副标题区", x: 0.07, y: 0.3, w: 0.44, h: 0.12 },
    { label: "主视觉主体", x: 0.5, y: 0.08, w: 0.42, h: 0.78 },
  ],
  "title-left": () => [
    { label: "标题与文案区", x: 0.06, y: 0.12, w: 0.34, h: 0.62 },
    { label: "主视觉区", x: 0.46, y: 0.08, w: 0.46, h: 0.8 },
    { label: "页脚或 Logo 区", x: 0.06, y: 0.8, w: 0.22, h: 0.08 },
  ],
  compare: () => [
    { label: "页标题区", x: 0.07, y: 0.08, w: 0.62, h: 0.12 },
    { label: "左侧对比区", x: 0.07, y: 0.26, w: 0.38, h: 0.58 },
    { label: "右侧对比区", x: 0.55, y: 0.26, w: 0.38, h: 0.58 },
  ],
  data: () => [
    { label: "标题区", x: 0.06, y: 0.08, w: 0.42, h: 0.12 },
    { label: "装饰主视觉区", x: 0.07, y: 0.25, w: 0.26, h: 0.52 },
    { label: "图表内容留白区", x: 0.4, y: 0.22, w: 0.52, h: 0.58 },
  ],
};

const ASSISTANT_MODE_LABELS = {
  generate: "生成单页图片提示词",
  revise: "修改当前图片提示词",
  edit: "生成改图提示词",
  pages: "生成多页视觉草案",
};

const PAGE_TYPE_LABELS = {
  cover: "封面页",
  content: "内容页",
  data: "数据页",
};

const DECORATION_LEVELS = {
  plain: {
    label: "朴素",
    prompt: "装饰强度：朴素。只保留少量不影响内容表达的修饰性图案，可使用轻微几何线条、分隔、纹理或材质层次，避免大面积花纹、复杂边框和抢眼背景。",
    overridePrompt: "最终装饰强度覆写：本页按“朴素”执行。修饰性图案务必克制，只做轻量陪衬，不得抢过标题、正文和关键数据。",
  },
  medium: {
    label: "中等",
    prompt: "装饰强度：中等。可以加入适量修饰性图案、几何形状、纹理、光影或辅助装饰，增强主题氛围，但必须保持信息层级清晰，不能盖过内容。",
    overridePrompt: "最终装饰强度覆写：本页按“中等”执行。允许适量主题化修饰图案，但仍需优先保证文字阅读、重点信息和版式秩序。",
  },
  complex: {
    label: "复杂",
    prompt: "装饰强度：复杂。可以在不影响阅读的前提下加入更丰富的修饰性图案、层叠几何、材质纹理、装饰边框和主题化背景细节，但所有装饰都只能服务内容，不得制造信息噪声。",
    overridePrompt: "最终装饰强度覆写：本页按“复杂”执行。允许更丰富的装饰性图案与层次细节，但必须避免压住文字、关键数字和内容容器。",
  },
};
const DEFAULT_DECORATION_LEVEL = "medium";

function normalizeDecorationLevel(value) {
  const normalized = String(value || "").trim().toLowerCase();
  return Object.prototype.hasOwnProperty.call(DECORATION_LEVELS, normalized)
    ? normalized
    : DEFAULT_DECORATION_LEVEL;
}

function getDecorationLevelLabel(value) {
  return DECORATION_LEVELS[normalizeDecorationLevel(value)]?.label || DECORATION_LEVELS[DEFAULT_DECORATION_LEVEL].label;
}

function getGlobalDecorationLevel() {
  return normalizeDecorationLevel(el?.themeDecorationLevel?.value);
}

function getPageDecorationLevel(page) {
  return normalizeDecorationLevel(page?.decorationLevel || getGlobalDecorationLevel());
}

function buildNoExtraTextConstraintBlock(options = {}) {
  const lines = [];
  if (options.scope === "decoration") {
    lines.push("装饰强度变量只控制非文字装饰图形的数量、密度和复杂度，不控制也不允许新增任何文字内容。");
  } else {
    lines.push("硬限制：除明确给出的上屏文字外，不得新增任何可读文字内容。");
  }
  if (options.requireWhitelist) {
    lines.push("凡是不在当前文字白名单里的内容，都不要出现在画面里。");
  }
  lines.push("禁止擅自生成任何汉字、英文单词、字母串、数字、年份、页码、角标、logo 文案、水印、来源、注释、标签、说明、占位词或伪文字纹理。");
  lines.push("允许新增的只可以是无字图形元素，例如几何图案、线条、纹理、光效、图标、箭头、容器、边框、材质层次和结构示意；这些元素内部也不得藏可读字符。");
  return lines.join("\n");
}

function buildDecorationPromptBlock(level, options = {}) {
  const normalized = normalizeDecorationLevel(level);
  const entry = DECORATION_LEVELS[normalized] || DECORATION_LEVELS[DEFAULT_DECORATION_LEVEL];
  return [
    options.override ? entry.overridePrompt : entry.prompt,
    "这些装饰只针对不影响内容的修饰性图案，不新增需要阅读的新文字、数字、标签或事实信息。",
    buildNoExtraTextConstraintBlock({ scope: "decoration" }),
  ].join("\n");
}

const TAB_DEFAULTS = {
  main: "smart",
};

const PPT_THEME_PRESETS = {
  papercraft: {
    label: "纸艺主题",
    summary: "参考 n8n 工作流里的纸艺方案，强调分层剪纸、折纸体块、舞台布景感，以及适合 PPT 的稳定留白。",
    basic: "你是一位顶级纸艺展示视觉设计师，请为 16:9 PPT 页面生成高端、克制、层次清晰的纸艺主视觉。整体采用立体折纸、层叠剪纸、卡纸结构和舞台布景式空间感，强调整洁留白、清晰层级、适合标题和正文排版，不要做成杂乱海报。",
    cover: "当前页面是封面页。请采用极简海报构图，只保留一个强视觉中心和大面积干净留白，为标题预留稳定区域，避免碎片化元素和复杂图表。",
    content: "当前页面是内容页。请采用卡片堆叠、纸艺图标、层级背景和易于放置正文的版式，重点服务于标题、要点和解释性文本排版。",
    data: "当前页面是数据页。请让图表或数据视觉化元素拥有纸艺拼贴和立体层次，同时保留足够空白给数字、标签和关键结论。",
  },
  inkwash: {
    label: "水墨主题",
    summary: "参考 n8n 工作流里的水墨方案，强调东方留白、流体水墨、宣纸肌理、玉石材质和现代简洁演示感。",
    basic: "你是一位顶级东方美学与现代演示设计师，请生成适合 16:9 PPT 的高雅水墨视觉。整体融合宣纸肌理、流体水墨、玉石与雾感光影，画面克制、空灵、现代，不要做成传统国画海报，而要适合商业演示文稿排版。",
    cover: "当前页面是封面页。请采用极强留白的东方封面构图，用一个主水墨意象作为视觉锚点，其余区域尽量纯净，方便放置标题。",
    content: "当前页面是内容页。请使用轻量的水墨块面、宣纸容器、简洁结构和适合文字排版的留白，避免复杂笔触遮挡正文。",
    data: "当前页面是数据页。请将图表感转译为水墨环形、流线、层叠山形或玉石仪表结构，保持数据展示区域清晰可用。",
  },
};

const state = {
  images: [],
  palette: [],
  slideRegions: [],
  themeLibrary: [],
  themeDefinition: null,
  themeDefinitionRaw: "",
  themeDefinitionSource: "",
  themeConfirmed: false,
  themeConfirmedSource: "",
  themeReferenceImages: [],
  editTargetImageId: null,
  currentTaskId: null,
  pollTimer: null,
  lastRequestId: null,
  resultImages: [],
  savedResults: {},
  slideDraft: null,
  assistantParsed: null,
  assistantRaw: "",
  workflowPlanSummary: "",
  workflowPlanRaw: "",
  workflowPages: [],
  workflowPlanLibrary: [],
  workflowPlanLibraryActiveId: "",
  workflowRunning: false,
  workflowPlanAbortController: null,
  workflowPlanStopRequested: false,
  workflowDetailPageId: null,
  workflowDetailDraft: { pageId: "", title: "", body: "", dirty: false },
  imagePreview: { title: "", url: "" },
  libraryDocLoaded: false,
  pptHarnessPack: null,
  pptHarnessReady: false,
  pptHarnessLoadError: "",
  activeTabs: { ...TAB_DEFAULTS },
};

let persistLibraryTimer = null;

const el = {
  craftStudioPanel: document.querySelector("#craftStudioPanel"),
  manualComposeHost: document.querySelector("#manualComposeHost"),
  reviseComposeHost: document.querySelector("#reviseComposeHost"),
  studioSectionLabel: document.querySelector("#studioSectionLabel"),
  studioTitle: document.querySelector("#studioTitle"),
  studioHint: document.querySelector("#studioHint"),
  promptLabel: document.querySelector("#promptLabel"),
  manualHelperPanel: document.querySelector("#manualHelperPanel"),
  assistantPanel: document.querySelector("#assistantPanel"),
  studioPrimaryActions: document.querySelector("#studioPrimaryActions"),
  debugPanel: document.querySelector("#debugPanel"),
  manualPageGoal: document.querySelector("#manualPageGoal"),
  manualBuildPromptBtn: document.querySelector("#manualBuildPromptBtn"),
  manualAppendPromptBtn: document.querySelector("#manualAppendPromptBtn"),
  manualCopyBriefBtn: document.querySelector("#manualCopyBriefBtn"),
  manualLayoutPreview: document.querySelector("#manualLayoutPreview"),
  themeName: document.querySelector("#themeName"),
  themeDecorationLevel: document.querySelector("#themeDecorationLevel"),
  generateThemeBtn: document.querySelector("#generateThemeBtn"),
  confirmThemeBtn: document.querySelector("#confirmThemeBtn"),
  themeStatus: document.querySelector("#themeStatus"),
  themeReferenceInput: document.querySelector("#themeReferenceInput"),
  learnThemeBtn: document.querySelector("#learnThemeBtn"),
  themeReferenceHint: document.querySelector("#themeReferenceHint"),
  themeReferenceList: document.querySelector("#themeReferenceList"),
  themeLibrarySelect: document.querySelector("#themeLibrarySelect"),
  pptHarnessEnabled: document.querySelector("#pptHarnessEnabled"),
  themeProgress: document.querySelector("#themeProgress"),
  themeReviewPanel: document.querySelector("#themeReviewPanel"),
  themeBasicPreview: document.querySelector("#themeBasicPreview"),
  themeCoverPreview: document.querySelector("#themeCoverPreview"),
  themeContentPreview: document.querySelector("#themeContentPreview"),
  themeDataPreview: document.querySelector("#themeDataPreview"),
  themeDefinitionPreview: document.querySelector("#themeDefinitionPreview"),
  apiKeyStatus: document.querySelector("#apiKeyStatus"),
  testApiKeyBtn: document.querySelector("#testApiKeyBtn"),
  autoReplaceTarget: document.querySelector("#autoReplaceTarget"),
  undoTargetBtn: document.querySelector("#undoTargetBtn"),
  apiKey: document.querySelector("#apiKey"),
  googleApiKey: document.querySelector("#googleApiKey"),
  requestMode: document.querySelector("#requestMode"),
  region: document.querySelector("#region"),
  model: document.querySelector("#model"),
  prompt: document.querySelector("#prompt"),
  fileInput: document.querySelector("#fileInput"),
  imageUrlInput: document.querySelector("#imageUrlInput"),
  addImageUrlBtn: document.querySelector("#addImageUrlBtn"),
  bboxEnabledImagePanel: document.querySelector("#bboxEnabledImagePanel"),
  editTargetSelect: document.querySelector("#editTargetSelect"),
  editTargetHint: document.querySelector("#editTargetHint"),
  imageList: document.querySelector("#imageList"),
  sizeMode: document.querySelector("#sizeMode"),
  presetSizeField: document.querySelector("#presetSizeField"),
  presetSize: document.querySelector("#presetSize"),
  customSizeField: document.querySelector("#customSizeField"),
  customWidth: document.querySelector("#customWidth"),
  customHeight: document.querySelector("#customHeight"),
  imageCount: document.querySelector("#imageCount"),
  seed: document.querySelector("#seed"),
  bboxEnabled: document.querySelector("#bboxEnabled"),
  enableSequential: document.querySelector("#enableSequential"),
  thinkingMode: document.querySelector("#thinkingMode"),
  watermark: document.querySelector("#watermark"),
  addPaletteBtn: document.querySelector("#addPaletteBtn"),
  paletteList: document.querySelector("#paletteList"),
  sendBtn: document.querySelector("#sendBtn"),
  refreshTaskBtn: document.querySelector("#refreshTaskBtn"),
  copyPayloadBtn: document.querySelector("#copyPayloadBtn"),
  downloadAllBtn: document.querySelector("#downloadAllBtn"),
  payloadPreview: document.querySelector("#payloadPreview"),
  responsePreview: document.querySelector("#responsePreview"),
  statusBar: document.querySelector("#statusBar"),
  usageBox: document.querySelector("#usageBox"),
  resultImages: document.querySelector("#resultImages"),
  slideAspect: document.querySelector("#slideAspect"),
  slideTemplate: document.querySelector("#slideTemplate"),
  includeRegionsInPrompt: document.querySelector("#includeRegionsInPrompt"),
  applyTemplateBtn: document.querySelector("#applyTemplateBtn"),
  syncSlideSizeBtn: document.querySelector("#syncSlideSizeBtn"),
  insertRegionsBtn: document.querySelector("#insertRegionsBtn"),
  clearRegionsBtn: document.querySelector("#clearRegionsBtn"),
  slideCanvas: document.querySelector("#slideCanvas"),
  slideCanvasOverlay: document.querySelector("#slideCanvasOverlay"),
  slideCanvasHint: document.querySelector("#slideCanvasHint"),
  slideRegionList: document.querySelector("#slideRegionList"),
  slideRegionPreview: document.querySelector("#slideRegionPreview"),
  assistantMode: document.querySelector("#assistantMode"),
  assistantUseImages: document.querySelector("#assistantUseImages"),
  assistantUseRegions: document.querySelector("#assistantUseRegions"),
  assistantUsePrompt: document.querySelector("#assistantUsePrompt"),
  assistantThinking: document.querySelector("#assistantThinking"),
  assistantRequest: document.querySelector("#assistantRequest"),
  assistantSendBtn: document.querySelector("#assistantSendBtn"),
  assistantApplyBtn: document.querySelector("#assistantApplyBtn"),
  assistantAppendBtn: document.querySelector("#assistantAppendBtn"),
  assistantSummary: document.querySelector("#assistantSummary"),
  assistantPreview: document.querySelector("#assistantPreview"),
  assistantPages: document.querySelector("#assistantPages"),
  workflowPageCount: document.querySelector("#workflowPageCount"),
  workflowTheme: document.querySelector("#workflowTheme"),
  workflowContent: document.querySelector("#workflowContent"),
  workflowPlanLibrarySelect: document.querySelector("#workflowPlanLibrarySelect"),
  workflowPlanHistoryPanel: document.querySelector("#workflowPlanHistoryPanel"),
  workflowPlanHistoryMeta: document.querySelector("#workflowPlanHistoryMeta"),
  workflowPlanHistoryLead: document.querySelector("#workflowPlanHistoryLead"),
  workflowPlanHistorySummary: document.querySelector("#workflowPlanHistorySummary"),
  workflowPlanHistoryPages: document.querySelector("#workflowPlanHistoryPages"),
  workflowPlanDeleteBtn: document.querySelector("#workflowPlanDeleteBtn"),
  workflowPlanBtn: document.querySelector("#workflowPlanBtn"),
  workflowPlanStopBtn: document.querySelector("#workflowPlanStopBtn"),
  workflowBatchBtn: document.querySelector("#workflowBatchBtn"),
  workflowCopyBtn: document.querySelector("#workflowCopyBtn"),
  workflowClearBtn: document.querySelector("#workflowClearBtn"),
  workflowGateHint: document.querySelector("#workflowGateHint"),
  workflowProgress: document.querySelector("#workflowProgress"),
  workflowThemePreview: document.querySelector("#workflowThemePreview"),
  workflowPlanSummary: document.querySelector("#workflowPlanSummary"),
  workflowPlanCards: document.querySelector("#workflowPlanCards"),
  workflowDetailModal: document.querySelector("#workflowDetailModal"),
  workflowDetailBackdrop: document.querySelector("#workflowDetailBackdrop"),
  workflowDetailPanel: document.querySelector("#workflowDetailPanel"),
  workflowDetailTitle: document.querySelector("#workflowDetailTitle"),
  workflowDetailMeta: document.querySelector("#workflowDetailMeta"),
  workflowDetailContent: document.querySelector("#workflowDetailContent"),
  workflowDetailDisplayTitleInput: document.querySelector("#workflowDetailDisplayTitleInput"),
  workflowDetailDecorationLevel: document.querySelector("#workflowDetailDecorationLevel"),
  workflowDetailVisibleTextEditor: document.querySelector("#workflowDetailVisibleTextEditor"),
  workflowDetailConfirmContentBtn: document.querySelector("#workflowDetailConfirmContentBtn"),
  workflowDetailConfirmStatus: document.querySelector("#workflowDetailConfirmStatus"),
  workflowDetailSuggestedContent: document.querySelector("#workflowDetailSuggestedContent"),
  workflowDetailConfirmedContent: document.querySelector("#workflowDetailConfirmedContent"),
  workflowDetailResearchQueryInput: document.querySelector("#workflowDetailResearchQueryInput"),
  workflowDetailResearchBtn: document.querySelector("#workflowDetailResearchBtn"),
  workflowDetailResearchApplyBtn: document.querySelector("#workflowDetailResearchApplyBtn"),
  workflowDetailResearchStatus: document.querySelector("#workflowDetailResearchStatus"),
  workflowDetailResearchList: document.querySelector("#workflowDetailResearchList"),
  workflowDetailVisibleResetBtn: document.querySelector("#workflowDetailVisibleResetBtn"),
  workflowDetailLayout: document.querySelector("#workflowDetailLayout"),
  workflowDetailTheme: document.querySelector("#workflowDetailTheme"),
  workflowDetailPrompt: document.querySelector("#workflowDetailPrompt"),
  workflowDetailUseBtn: document.querySelector("#workflowDetailUseBtn"),
  workflowDetailCopyBtn: document.querySelector("#workflowDetailCopyBtn"),
  workflowDetailRunBtn: document.querySelector("#workflowDetailRunBtn"),
  workflowDetailCloseBtn: document.querySelector("#workflowDetailCloseBtn"),
  imagePreviewModal: document.querySelector("#imagePreviewModal"),
  imagePreviewBackdrop: document.querySelector("#imagePreviewBackdrop"),
  imagePreviewTitle: document.querySelector("#imagePreviewTitle"),
  imagePreviewImage: document.querySelector("#imagePreviewImage"),
  imagePreviewPrompt: document.querySelector("#imagePreviewPrompt"),
  imagePreviewUseTargetBtn: document.querySelector("#imagePreviewUseTargetBtn"),
  imagePreviewOpenReviseBtn: document.querySelector("#imagePreviewOpenReviseBtn"),
  imagePreviewEditBtn: document.querySelector("#imagePreviewEditBtn"),
  imagePreviewEditHint: document.querySelector("#imagePreviewEditHint"),
  imagePreviewOpenNew: document.querySelector("#imagePreviewOpenNew"),
  imagePreviewCloseBtn: document.querySelector("#imagePreviewCloseBtn"),
};

const uid = () => `${Date.now()}-${Math.random().toString(16).slice(2)}`;
const clamp = (value, min, max) => Math.min(max, Math.max(min, value));
const formatJSON = (value) => JSON.stringify(value, null, 2);

function sanitizeForPreview(value) {
  if (Array.isArray(value)) return value.map((item) => sanitizeForPreview(item));
  if (!value || typeof value !== "object") return value;
  const next = {};
  Object.entries(value).forEach(([key, item]) => {
    if (key === "image" && typeof item === "string" && item.startsWith("data:")) {
      next[key] = `[base64 image omitted, ${item.length} chars]`;
      return;
    }
    next[key] = typeof item === "object" ? sanitizeForPreview(item) : item;
  });
  return next;
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function normalizeRegion(value) {
  return value === DEFAULT_REGION ? value : DEFAULT_REGION;
}

function ensureGoogleApiKeyField() {
  if (el.googleApiKey) return el.googleApiKey;
  const settingsGrid = document.querySelector(".settings-grid");
  const dashScopeField = el.apiKey?.closest(".field");
  if (!settingsGrid || !dashScopeField) return null;

  const field = document.createElement("label");
  field.className = "field";
  field.innerHTML = [
    "<span>Google API Key (Nano Banana)</span>",
    '<input id="googleApiKey" type="password" placeholder="AIza..." autocomplete="off" />',
  ].join("");
  settingsGrid.insertBefore(field, dashScopeField.nextElementSibling);
  el.googleApiKey = field.querySelector("#googleApiKey");
  return el.googleApiKey;
}

function isGeminiImageModel(model = el?.model?.value) {
  return GEMINI_IMAGE_MODELS.has(String(model || "").trim());
}

function getDashScopeApiKey() {
  return String(el?.apiKey?.value || "").trim();
}

function getGoogleApiKey() {
  return String(el?.googleApiKey?.value || "").trim();
}

function getImageGenerationApiKey(model = el?.model?.value) {
  return isGeminiImageModel(model) ? getGoogleApiKey() : getDashScopeApiKey();
}

function getImageProviderLabel(model = el?.model?.value) {
  return isGeminiImageModel(model) ? "Nano Banana Pro" : "DashScope";
}

function getImageApiKeyLabel(model = el?.model?.value) {
  return isGeminiImageModel(model) ? "Google API Key" : "DashScope API Key";
}

function shouldUseAsyncImageGeneration(model = el?.model?.value) {
  return !isGeminiImageModel(model) && el.requestMode?.value === "async";
}

function syncImageProviderUi() {
  const gemini = isGeminiImageModel();

  if (el.requestMode) {
    if (gemini) el.requestMode.value = "sync";
    el.requestMode.disabled = gemini;
    el.requestMode.title = gemini ? "Nano Banana 暂时只支持同步出图。" : "";
  }

  if (el.region) {
    el.region.disabled = gemini;
    el.region.title = gemini ? "Nano Banana 不使用 DashScope 区域设置。" : "";
  }

  if (el.testApiKeyBtn) {
    const idleLabel = gemini ? "测试 Nano Banana Key" : "测试 DashScope Key";
    el.testApiKeyBtn.dataset.idleLabel = idleLabel;
    if (!el.testApiKeyBtn.disabled) {
      el.testApiKeyBtn.textContent = idleLabel;
    }
  }
}

function enforceRegionOptions() {
  if (!el.region) return;
  el.region.innerHTML = '<option value="beijing">北京</option>';
  el.region.value = DEFAULT_REGION;
}

function formatHistoryTime(timestamp) {
  const value = Number(timestamp);
  if (!value) return "时间未知";
  const date = new Date(value);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hour = String(date.getHours()).padStart(2, "0");
  const minute = String(date.getMinutes()).padStart(2, "0");
  return `${year}-${month}-${day} ${hour}:${minute}`;
}

function maskApiKey(value) {
  const text = String(value || "").trim();
  if (!text) return "";
  if (text.length <= 10) return text;
  return `${text.slice(0, 6)}...${text.slice(-4)}`;
}

function renderApiKeyStatus(message, tone = "idle") {
  if (!el.apiKeyStatus) return;
  const key = el.apiKey.value.trim();
  const defaultMessage = key
    ? `当前已保存 API Key：${maskApiKey(key)}。它会持久化保存在项目文档 data/studio-library.json。`
    : "还没有填写 API Key。填一次后会长期保存在项目文档 data/studio-library.json。";
  el.apiKeyStatus.textContent = message || defaultMessage;
  el.apiKeyStatus.dataset.tone = tone;
}

function syncApiKeyFeedback(response, data) {
  const message = String(data?.message || "");
  if (response?.status === 401 || /invalid api-?key/i.test(message)) {
    renderApiKeyStatus(
      "这串 API Key 已经保存在本地，但当前调用被上游接口以 401 / Invalid API-key 拒绝。请检查是否复制完整、是否已过期，或是否用了其他环境的 Key。",
      "error",
    );
    return;
  }

  if (response?.ok && el.apiKey.value.trim()) {
    renderApiKeyStatus();
  }
}

async function testApiKey() {
  const apiKey = el.apiKey.value.trim();
  if (!apiKey) {
    renderApiKeyStatus("请先输入 API Key，再点击测试。", "error");
    setStatus("请先输入 API Key。", "error");
    return;
  }

  saveSettings();

  const payload = {
    model: "qwen3.6-plus",
    input: {
      messages: [
        {
          role: "user",
          content: [{ text: "请只回复 OK" }],
        },
      ],
    },
    parameters: {
      result_format: "message",
      enable_thinking: false,
    },
  };

  setButtonLoading(el.testApiKeyBtn, true, "测试中...");
  renderApiKeyStatus(`正在测试 ${maskApiKey(apiKey)}，会验证当前区域和 Qwen 调用链路。`, "running");
  setStatus("正在测试 API Key...", "running");

  try {
    const response = await fetch("/api/assistant", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        apiKey,
        region: el.region.value,
        payload,
      }),
    });
    const data = await response.json();
    syncApiKeyFeedback(response, data);

    if (!response.ok || data.code) {
      const requestId = data?.request_id ? ` request_id: ${data.request_id}` : "";
      const message = data?.message || "API Key 测试未通过。";
      renderApiKeyStatus(`API Key 测试未通过：${message}${requestId}`, "error");
      setStatus(message, "error");
      return;
    }

    renderApiKeyStatus(
      `API Key 测试通过：${maskApiKey(apiKey)} 在当前区域可正常调用 Qwen。这个测试只验证鉴权和 Qwen 链路，图片模型仍以实际生图为准。`,
      "success",
    );
    setStatus("API Key 测试通过，可以继续生成。", "success");
  } catch (error) {
    renderApiKeyStatus(`API Key 测试失败：${error.message || "网络请求失败。"}`, "error");
    setStatus(error.message || "API Key 测试失败。", "error");
  } finally {
    setButtonLoading(el.testApiKeyBtn, false);
  }
}

function setStatus(message, tone = "idle") {
  el.statusBar.textContent = message;
  el.statusBar.dataset.tone = tone;
}

function stopPolling() {
  if (state.pollTimer) {
    clearInterval(state.pollTimer);
    state.pollTimer = null;
  }
}

function createAbortError(message = "已停止提交。") {
  const error = new Error(message);
  error.name = "AbortError";
  return error;
}

function isAbortError(error) {
  return error?.name === "AbortError";
}

function syncWorkflowPlanStopButton() {
  if (!el.workflowPlanStopBtn) return;
  const running = Boolean(state.workflowPlanAbortController);
  el.workflowPlanStopBtn.hidden = !running;
  el.workflowPlanStopBtn.classList.toggle("hidden", !running);
  el.workflowPlanStopBtn.disabled = !running || state.workflowPlanStopRequested;
  el.workflowPlanStopBtn.textContent = state.workflowPlanStopRequested ? "停止中..." : "停止";
}

function beginWorkflowPlanRun() {
  if (state.workflowPlanAbortController) {
    state.workflowPlanAbortController.abort(createAbortError("已停止上一轮拆分页。"));
  }
  state.workflowPlanAbortController = new AbortController();
  state.workflowPlanStopRequested = false;
  syncWorkflowPlanStopButton();
  return state.workflowPlanAbortController;
}

function endWorkflowPlanRun(controller) {
  if (!controller || state.workflowPlanAbortController === controller) {
    state.workflowPlanAbortController = null;
    state.workflowPlanStopRequested = false;
    syncWorkflowPlanStopButton();
  }
}

function cancelWorkflowPlanRun() {
  if (!state.workflowPlanAbortController || state.workflowPlanStopRequested) return;
  state.workflowPlanStopRequested = true;
  syncWorkflowPlanStopButton();
  state.workflowPlanAbortController.abort(createAbortError("已停止拆分页。"));
  setStatus("正在停止拆分页...", "running");
}

function getCurrentAspectMeta() {
  return SLIDE_ASPECTS[el.slideAspect.value] || SLIDE_ASPECTS["16:9"];
}

function getCanvasRelativePoint(event, canvas) {
  const rect = canvas.getBoundingClientRect();
  const scaleX = rect.width ? canvas.width / rect.width : 1;
  const scaleY = rect.height ? canvas.height / rect.height : 1;
  return {
    x: clamp((event.clientX - rect.left) * scaleX, 0, canvas.width),
    y: clamp((event.clientY - rect.top) * scaleY, 0, canvas.height),
  };
}

function extractResultImages(response) {
  const images = [];
  (response.output?.choices || []).forEach((choice) => {
    (choice.message?.content || []).forEach((item) => {
      if (item.type === "image" && item.image) images.push(item.image);
    });
  });
  return images;
}

function serializeWorkflowPage(page) {
  return {
    pageNumber: page.pageNumber,
    pageType: page.pageType,
    pageTitle: page.pageTitle,
    pageContent: page.pageContent,
    decorationLevel: getPageDecorationLevel(page),
    layoutSummary: page.layoutSummary || "",
    textHierarchy: page.textHierarchy || "",
    visualFocus: page.visualFocus || "",
    readabilityNotes: page.readabilityNotes || "",
    pagePrompt: page.pagePrompt || "",
    resultImages: Array.isArray(page.resultImages) ? page.resultImages.filter(Boolean) : [],
    savedResults: page.savedResults && typeof page.savedResults === "object" ? page.savedResults : {},
    requestId: page.requestId || "",
    taskId: page.taskId || "",
    detailBackdropUrl: page.detailBackdropUrl || "",
  };
}

function hydrateWorkflowPages(savedPages) {
  if (!Array.isArray(savedPages)) return [];
  return savedPages
    .map((page, index) => {
      const pageNumber = Number(page.pageNumber ?? page.page_number ?? index + 1) || index + 1;
      const pageType = normalizeWorkflowPageType(page.pageType ?? page.page_type, index);
      const pageTitle = String(page.pageTitle ?? page.page_title ?? `第 ${pageNumber} 页`).trim();
      const pageContent = String(page.pageContent ?? page.page_content ?? pageTitle).trim();
      return {
        id: uid(),
        pageNumber,
        pageType,
        pageTitle: pageTitle || `第 ${pageNumber} 页`,
        pageContent,
        decorationLevel: normalizeDecorationLevel(page.decorationLevel || page.decoration_level),
        layoutSummary: String(page.layoutSummary || page.layout_summary || "").trim(),
        textHierarchy: String(page.textHierarchy || page.text_hierarchy || "").trim(),
        visualFocus: String(page.visualFocus || page.visual_focus || "").trim(),
        readabilityNotes: String(page.readabilityNotes || page.readability_notes || "").trim(),
        pagePrompt: String(page.pagePrompt || page.page_prompt || "").trim(),
        layoutStatus: "idle",
        layoutError: "",
        layoutPromise: null,
        status: "idle",
        error: "",
        resultImages: Array.isArray(page.resultImages || page.result_images)
          ? (page.resultImages || page.result_images).filter(Boolean)
          : [],
        savedResults: page.savedResults && typeof page.savedResults === "object" ? page.savedResults : {},
        requestId: String(page.requestId || page.request_id || "").trim(),
        taskId: String(page.taskId || page.task_id || "").trim(),
        detailBackdropUrl: String(page.detailBackdropUrl || page.detail_backdrop_url || "").trim(),
      };
    })
    .filter((page) => page.pageTitle || page.pageContent)
    .sort((a, b) => a.pageNumber - b.pageNumber)
    .map((page, index) => {
      const nextPage = {
        ...page,
        pageNumber: index + 1,
        pageType: index === 0 ? "cover" : page.pageType,
      };
      if (nextPage.pagePrompt) {
        nextPage.layoutStatus = "ready";
      } else {
        try {
          nextPage.pagePrompt = buildWorkflowPagePrompt(nextPage);
        } catch {
          nextPage.pagePrompt = "";
        }
      }
      return nextPage;
    });
}

function normalizeWorkflowPlanLibraryEntry(entry) {
  const label = String(entry?.label || "").trim();
  const workflowContent = String(entry?.workflowContent || "").trim();
  const pages = hydrateWorkflowPages(entry?.pages);
  if (!label || !workflowContent || !pages.length) return null;
  return {
    id: String(entry?.id || uid()),
    label,
    themeLabel: String(entry?.themeLabel || entry?.themeName || "").trim(),
    workflowContent,
    workflowPageCount: String(entry?.workflowPageCount || pages.length),
    workflowPlanSummary: String(entry?.workflowPlanSummary || "").trim(),
    pages,
    createdAt: Number(entry?.createdAt || entry?.updatedAt || Date.now()),
    updatedAt: Number(entry?.updatedAt || Date.now()),
  };
}

function buildSettingsSnapshot() {
  return {
    themeWorkflowVersion: THEME_WORKFLOW_VERSION,
    apiKey: getDashScopeApiKey(),
    googleApiKey: getGoogleApiKey(),
    requestMode: el.requestMode.value,
    region: normalizeRegion(el.region.value),
    model: el.model.value,
    sizeMode: el.sizeMode.value,
    presetSize: el.presetSize.value,
    customWidth: el.customWidth.value,
    customHeight: el.customHeight.value,
    imageCount: el.imageCount.value,
    seed: el.seed.value,
    enableSequential: el.enableSequential.checked,
    thinkingMode: el.thinkingMode.checked,
    watermark: el.watermark.checked,
    bboxEnabled: el.bboxEnabled.checked,
    editTargetImageId: state.editTargetImageId,
    palette: state.palette,
    slideAspect: el.slideAspect.value,
    slideTemplate: el.slideTemplate.value,
    includeRegionsInPrompt: el.includeRegionsInPrompt.checked,
    slideRegions: state.slideRegions,
    assistantMode: el.assistantMode.value,
    assistantUseImages: el.assistantUseImages.checked,
    assistantUseRegions: el.assistantUseRegions.checked,
    assistantUsePrompt: el.assistantUsePrompt.checked,
    assistantThinking: el.assistantThinking.checked,
    assistantRequest: el.assistantRequest.value,
    manualPageGoal: el.manualPageGoal.value,
    autoReplaceTarget: el.autoReplaceTarget.checked,
    themeName: el.themeName.value,
    themeDecorationLevel: getGlobalDecorationLevel(),
    themeLibrary: state.themeLibrary,
    themeDefinition: state.themeDefinition,
    themeDefinitionRaw: state.themeDefinitionRaw,
    themeDefinitionSource: state.themeDefinitionSource,
    themeConfirmed: state.themeConfirmed,
    themeConfirmedSource: state.themeConfirmedSource,
    workflowPageCount: el.workflowPageCount.value,
    workflowTheme: el.workflowTheme.value || el.themeName.value.trim(),
    workflowContent: el.workflowContent.value,
    workflowPlanSummary: state.workflowPlanSummary,
    workflowPages: state.workflowPages.map((page) => serializeWorkflowPage(page)),
    workflowPlanLibrary: state.workflowPlanLibrary.map((entry) => ({
      ...entry,
      pages: entry.pages.map((page) => serializeWorkflowPage(page)),
    })),
    workflowPlanLibraryActiveId: state.workflowPlanLibraryActiveId,
    activeTabs: state.activeTabs,
  };
}

function loadSettings() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}");
    const useThemeWorkflowV2 = saved.themeWorkflowVersion === THEME_WORKFLOW_VERSION;
    el.apiKey.value = saved.apiKey || "";
    if (el.googleApiKey) el.googleApiKey.value = saved.googleApiKey || "";
    el.requestMode.value = saved.requestMode || "sync";
    el.region.value = normalizeRegion(saved.region);
  el.model.value = saved.model || "gemini-3-pro-image-preview";
    el.sizeMode.value = saved.sizeMode || "preset";
    el.presetSize.value = saved.presetSize || "4K";
    el.customWidth.value = saved.customWidth || "";
    el.customHeight.value = saved.customHeight || "";
    el.imageCount.value = saved.imageCount || "1";
    el.seed.value = saved.seed || "";
    el.enableSequential.checked = Boolean(saved.enableSequential);
    el.thinkingMode.checked = saved.thinkingMode !== false;
    el.watermark.checked = Boolean(saved.watermark);
    el.bboxEnabled.checked = Boolean(saved.bboxEnabled);
    el.bboxEnabledImagePanel.checked = Boolean(saved.bboxEnabled);
    el.slideAspect.value = saved.slideAspect || "16:9";
    el.slideTemplate.value = saved.slideTemplate || "blank";
    el.includeRegionsInPrompt.checked = saved.includeRegionsInPrompt !== false;
    el.assistantMode.value = saved.assistantMode || "generate";
    el.assistantUseImages.checked = saved.assistantUseImages !== false;
    el.assistantUseRegions.checked = saved.assistantUseRegions !== false;
    el.assistantUsePrompt.checked = saved.assistantUsePrompt !== false;
    el.assistantThinking.checked = Boolean(saved.assistantThinking);
    el.assistantRequest.value = saved.assistantRequest || "";
    el.manualPageGoal.value = saved.manualPageGoal || "";
    el.autoReplaceTarget.checked = saved.autoReplaceTarget !== false;
    el.themeName.value = useThemeWorkflowV2 ? (saved.themeName || "") : "";
    if (el.themeDecorationLevel) {
      el.themeDecorationLevel.value = normalizeDecorationLevel(saved.themeDecorationLevel);
    }
    el.workflowPageCount.value = saved.workflowPageCount || "4";
    el.workflowTheme.value = useThemeWorkflowV2 ? (saved.workflowTheme || saved.themeName || "") : "";
    el.workflowContent.value = saved.workflowContent || "";
    const main = saved.activeTabs?.main;
    state.activeTabs = {
      main: ["smart", "manual", "revise"].includes(main) ? main : TAB_DEFAULTS.main,
    };
    state.palette = Array.isArray(saved.palette) ? saved.palette : [];
    state.slideRegions = Array.isArray(saved.slideRegions) ? saved.slideRegions : [];
    state.themeLibrary = useThemeWorkflowV2 && Array.isArray(saved.themeLibrary)
      ? saved.themeLibrary.map((entry) => normalizeThemeLibraryEntry(entry)).filter(Boolean)
      : [];
    state.themeDefinition = useThemeWorkflowV2 ? (saved.themeDefinition || null) : null;
    state.themeDefinitionRaw = useThemeWorkflowV2 ? (saved.themeDefinitionRaw || "") : "";
    state.themeDefinitionSource = useThemeWorkflowV2 ? (saved.themeDefinitionSource || "") : "";
    state.themeConfirmed = useThemeWorkflowV2 ? Boolean(saved.themeConfirmed) : false;
    state.themeConfirmedSource = useThemeWorkflowV2 ? (saved.themeConfirmedSource || "") : "";
    state.workflowPlanSummary = useThemeWorkflowV2 ? (saved.workflowPlanSummary || "") : "";
    state.workflowPages = useThemeWorkflowV2 ? hydrateWorkflowPages(saved.workflowPages) : [];
    state.workflowPlanLibrary = useThemeWorkflowV2 && Array.isArray(saved.workflowPlanLibrary)
      ? saved.workflowPlanLibrary.map((entry) => normalizeWorkflowPlanLibraryEntry(entry)).filter(Boolean)
      : [];
    state.workflowPlanLibraryActiveId = useThemeWorkflowV2 ? (saved.workflowPlanLibraryActiveId || "") : "";
    state.editTargetImageId = saved.editTargetImageId || null;
    if (!el.region.value) el.region.value = DEFAULT_REGION;
  } catch {
    state.palette = [];
    state.slideRegions = [];
    state.themeLibrary = [];
    state.themeDefinition = null;
    state.themeDefinitionRaw = "";
    state.themeDefinitionSource = "";
    state.themeConfirmed = false;
    state.themeConfirmedSource = "";
    state.workflowPlanSummary = "";
    state.workflowPages = [];
    state.workflowPlanLibrary = [];
    state.workflowPlanLibraryActiveId = "";
    state.editTargetImageId = null;
    state.activeTabs = { ...TAB_DEFAULTS };
  }
}

function saveSettings() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(buildSettingsSnapshot()));
  schedulePersistentLibrarySave();
}

function buildLibraryDoc() {
  const settings = buildSettingsSnapshot();
  return {
    version: LIBRARY_DOC_VERSION,
    savedAt: new Date().toISOString(),
    settings: {
      apiKey: settings.apiKey,
      googleApiKey: settings.googleApiKey,
      requestMode: settings.requestMode,
      region: settings.region,
      model: settings.model,
      sizeMode: settings.sizeMode,
      presetSize: settings.presetSize,
      customWidth: settings.customWidth,
      customHeight: settings.customHeight,
      imageCount: settings.imageCount,
      seed: settings.seed,
      enableSequential: settings.enableSequential,
      thinkingMode: settings.thinkingMode,
      watermark: settings.watermark,
      slideAspect: settings.slideAspect,
      includeRegionsInPrompt: settings.includeRegionsInPrompt,
      assistantThinking: settings.assistantThinking,
      themeName: settings.themeName,
      themeDecorationLevel: settings.themeDecorationLevel,
      themeDefinition: settings.themeDefinition,
      themeDefinitionRaw: settings.themeDefinitionRaw,
      themeDefinitionSource: settings.themeDefinitionSource,
      themeConfirmed: settings.themeConfirmed,
      themeConfirmedSource: settings.themeConfirmedSource,
      workflowPlanLibraryActiveId: settings.workflowPlanLibraryActiveId,
    },
    themes: state.themeLibrary.map((entry) => ({
      ...entry,
      definition: entry.definition,
    })),
    workflowPlans: state.workflowPlanLibrary.map((entry) => ({
      ...entry,
      pages: entry.pages.map((page) => serializeWorkflowPage(page)),
    })),
  };
}

function applyLibrarySettings(settings = {}) {
  const assignValue = (node, value) => {
    if (node && value !== undefined && value !== null) node.value = String(value);
  };
  const assignChecked = (node, value) => {
    if (node && value !== undefined) node.checked = Boolean(value);
  };

  assignValue(el.apiKey, settings.apiKey);
  assignValue(el.googleApiKey, settings.googleApiKey);
  assignValue(el.requestMode, settings.requestMode);
  assignValue(el.region, normalizeRegion(settings.region));
  assignValue(el.model, settings.model);
  assignValue(el.sizeMode, settings.sizeMode);
  assignValue(el.presetSize, settings.presetSize);
  assignValue(el.customWidth, settings.customWidth);
  assignValue(el.customHeight, settings.customHeight);
  assignValue(el.imageCount, settings.imageCount);
  assignValue(el.seed, settings.seed);
  assignChecked(el.enableSequential, settings.enableSequential);
  assignChecked(el.thinkingMode, settings.thinkingMode);
  assignChecked(el.watermark, settings.watermark);
  assignValue(el.slideAspect, settings.slideAspect);
  assignChecked(el.includeRegionsInPrompt, settings.includeRegionsInPrompt);
  assignChecked(el.assistantThinking, settings.assistantThinking);
  if (!el.region.value) el.region.value = DEFAULT_REGION;

  if (typeof settings.themeName === "string") {
    el.themeName.value = settings.themeName;
    el.workflowTheme.value = settings.themeName;
  }
  if (el.themeDecorationLevel && settings.themeDecorationLevel !== undefined) {
    el.themeDecorationLevel.value = normalizeDecorationLevel(settings.themeDecorationLevel);
  }
  if (settings.themeDefinition && typeof settings.themeDefinition === "object") {
    state.themeDefinition = normalizeThemeDefinition(settings.themeDefinition, settings.themeName || settings.themeDefinitionSource || "");
  }
  if (typeof settings.themeDefinitionRaw === "string") {
    state.themeDefinitionRaw = settings.themeDefinitionRaw;
  }
  if (typeof settings.themeDefinitionSource === "string") {
    state.themeDefinitionSource = settings.themeDefinitionSource;
  }
  if (settings.themeConfirmed !== undefined) {
    state.themeConfirmed = Boolean(settings.themeConfirmed);
  }
  if (typeof settings.themeConfirmedSource === "string") {
    state.themeConfirmedSource = settings.themeConfirmedSource;
  }
  if (typeof settings.workflowPlanLibraryActiveId === "string") {
    state.workflowPlanLibraryActiveId = settings.workflowPlanLibraryActiveId;
  }
}

async function persistLibraryDoc() {
  const response = await fetch("/api/library", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ doc: buildLibraryDoc() }),
  });
  const data = await response.json();
  if (!response.ok || data.code) {
    throw new Error(data.message || "保存项目资料库失败。");
  }
}

function schedulePersistentLibrarySave() {
  if (!state.libraryDocLoaded) return;
  if (persistLibraryTimer) {
    clearTimeout(persistLibraryTimer);
  }
  persistLibraryTimer = setTimeout(async () => {
    persistLibraryTimer = null;
    try {
      await persistLibraryDoc();
    } catch (error) {
      setStatus(error.message || "保存项目资料库失败。", "error");
    }
  }, 400);
}

async function loadPersistentLibrary() {
  try {
    const response = await fetch("/api/library");
    const data = await response.json();
    if (!response.ok || data.code) {
      throw new Error(data.message || "读取项目资料库失败。");
    }

    const doc = data.doc || {};
    applyLibrarySettings(doc.settings || {});

    const themeEntries = Array.isArray(doc.themes)
      ? doc.themes.map((entry) => normalizeThemeLibraryEntry(entry)).filter(Boolean)
      : [];
    if (themeEntries.length) {
      state.themeLibrary = themeEntries;
    }

    const workflowPlanEntries = Array.isArray(doc.workflowPlans)
      ? doc.workflowPlans.map((entry) => normalizeWorkflowPlanLibraryEntry(entry)).filter(Boolean)
      : [];
    if (workflowPlanEntries.length) {
      state.workflowPlanLibrary = workflowPlanEntries;
    }
  } catch (error) {
    setStatus(error.message || "读取项目资料库失败。", "error");
  } finally {
    state.libraryDocLoaded = true;
    saveSettings();
  }
}

function applyTabState() {
  const main = state.activeTabs.main || TAB_DEFAULTS.main;

  document.querySelectorAll("[data-main-tab]").forEach((button) => {
    const active = button.dataset.mainTab === main;
    button.classList.toggle("active", active);
    button.setAttribute("aria-selected", active ? "true" : "false");
  });

  document.querySelectorAll("[data-main-panel]").forEach((panel) => {
    panel.hidden = panel.dataset.mainPanel !== main;
  });

  if (main === "manual") {
    el.manualComposeHost.appendChild(el.craftStudioPanel);
    el.craftStudioPanel.hidden = false;
    el.craftStudioPanel.classList.remove("revise-simple");
    el.manualHelperPanel.hidden = false;
    if (el.assistantPanel) el.assistantPanel.hidden = false;
    if (el.downloadAllBtn) el.downloadAllBtn.hidden = false;
    if (el.debugPanel) el.debugPanel.hidden = true;
    if (el.usageBox) {
      el.usageBox.classList.add("hidden");
      el.usageBox.innerHTML = "";
    }
    el.studioSectionLabel.textContent = "流程 2 右侧面板";
    el.studioTitle.textContent = "手动流：固定提示词、润色并出图";
    el.studioHint.textContent = "先规划区域和文案，再把主题风格写进主提示词骨架，需要时可继续让 Qwen 帮你润色。";
    el.promptLabel.textContent = "本页主提示词";
    el.sendBtn.textContent = "开始调用";
    el.prompt.rows = 7;
    renderSlideCanvas();
  } else if (main === "revise") {
    el.reviseComposeHost.appendChild(el.craftStudioPanel);
    el.craftStudioPanel.hidden = false;
    el.craftStudioPanel.classList.add("revise-simple");
    el.manualHelperPanel.hidden = true;
    if (el.assistantPanel) el.assistantPanel.hidden = true;
    if (el.downloadAllBtn) el.downloadAllBtn.hidden = true;
    if (el.debugPanel) el.debugPanel.hidden = true;
    if (el.usageBox) {
      el.usageBox.classList.add("hidden");
      el.usageBox.innerHTML = "";
    }
    el.studioSectionLabel.textContent = "流程 3 右侧面板";
    el.studioTitle.textContent = "修改提示词与提交";
    el.studioHint.textContent = "右边直接写你想改什么，提交后结果会回到下面。";
    el.promptLabel.textContent = "修改提示词";
    el.sendBtn.textContent = "提交修改";
    el.prompt.rows = 10;
  } else {
    el.craftStudioPanel.hidden = true;
  }

  renderApiKeyStatus();
  renderManualLayoutPreview();
  renderEditTargetControls();
}

function bindTabNavigation() {
  document.querySelectorAll("[data-main-tab]").forEach((button) => {
    button.addEventListener("click", () => {
      state.activeTabs.main = button.dataset.mainTab || TAB_DEFAULTS.main;
      saveSettings();
      applyTabState();
    });
  });
}

function toggleSizeMode() {
  const custom = el.sizeMode.value === "custom";
  el.presetSizeField.classList.toggle("hidden", custom);
  el.customSizeField.classList.toggle("hidden", !custom);
}

function buildSizeValue() {
  if (el.sizeMode.value === "preset") {
    if (el.model.value === "wan2.7-image" && el.presetSize.value === "4K") {
      throw new Error("wan2.7-image 不支持 4K 预设规格。");
    }
    return el.presetSize.value;
  }
  const width = Number(el.customWidth.value);
  const height = Number(el.customHeight.value);
  if (!width || !height) throw new Error("自定义尺寸模式下请同时填写宽和高。");
  return `${width}*${height}`;
}

function buildSlideRegionText() {
  if (!state.slideRegions.length) return "";
  const aspect = getCurrentAspectMeta();
  const lines = [`PPT版式比例：${aspect.label}（建议输出尺寸 ${aspect.outputWidth}x${aspect.outputHeight}）`, "PPT版式区域要求："];
  state.slideRegions.forEach((region, index) => {
    const x = (region.x * 100).toFixed(1);
    const y = (region.y * 100).toFixed(1);
    const w = (region.w * 100).toFixed(1);
    const h = (region.h * 100).toFixed(1);
    lines.push(`${index + 1}. ${region.label}：位置 ${x}%/${y}%，大小 ${w}% x ${h}%`);
  });
  lines.push("请根据这些区域安排主体位置、留白和文字可读性，不要让重要主体压住标题和正文。");
  return lines.join("\n");
}

function buildPromptText() {
  const prompt = el.prompt.value.trim();
  if (!prompt) throw new Error("请先填写提示词。");
  if (!el.includeRegionsInPrompt.checked || !state.slideRegions.length) return prompt;
  return `${prompt}\n\n${buildSlideRegionText()}`;
}

function getEditTargetImage() {
  return state.images.find((image) => image.id === state.editTargetImageId) || null;
}

function ensureEditTargetSelection() {
  if (!state.images.length) {
    state.editTargetImageId = null;
    return;
  }
  const exists = state.images.some((image) => image.id === state.editTargetImageId);
  if (!exists) {
    state.editTargetImageId = state.images[0].id;
  }
}

function getOrderedImagesForPayload() {
  if (!el.bboxEnabled.checked) {
    return state.images;
  }
  const target = getEditTargetImage();
  if (!target) {
    throw new Error("启用定点改图时，请先选择一张改图底图。");
  }
  const refs = state.images.filter((image) => image.id !== target.id);
  return [...refs, target];
}

function buildEditTargetText() {
  if (!el.bboxEnabled.checked) return "";
  const target = getEditTargetImage();
  if (!target) return "";
  const boxLines = target.boxes.map((box, index) => `区域 ${index + 1}: [${box.join(", ")}]`).join("\n");
  return [
    "定点改图要求：",
    "请以最后一张底图为基础，只修改框选区域，其他页面结构、文字排版和非框选内容尽量保持稳定。",
    `改图底图：${target.name}`,
    boxLines || "当前还没有框选区域。",
  ].join("\n");
}

function syncBboxToggles(checked) {
  el.bboxEnabled.checked = checked;
  el.bboxEnabledImagePanel.checked = checked;
}

function updatePayloadPreview() {
  try {
    el.payloadPreview.textContent = formatJSON(sanitizeForPreview(buildPayload()));
  } catch (error) {
    el.payloadPreview.textContent = `// ${error.message}`;
  }
}

function renderPalette() {
  if (!state.palette.length) {
    el.paletteList.innerHTML = '<p class="hint">当前未配置调色板。</p>';
    saveSettings();
    updatePayloadPreview();
    return;
  }
  el.paletteList.innerHTML = "";
  state.palette.forEach((item) => {
    const row = document.createElement("div");
    row.className = "palette-row";
    row.innerHTML = `
      <input type="color" value="${item.hex || "#C2D1E6"}" />
      <input type="text" value="${item.hex || ""}" placeholder="#C2D1E6" />
      <input type="text" value="${item.ratio || ""}" placeholder="25.00%" />
      <button type="button" class="ghost-btn">删除</button>
    `;
    const inputs = row.querySelectorAll("input");
    row.querySelector("button").addEventListener("click", () => {
      state.palette = state.palette.filter((entry) => entry.id !== item.id);
      renderPalette();
    });
    inputs[0].addEventListener("input", () => { item.hex = inputs[0].value.toUpperCase(); inputs[1].value = item.hex; saveSettings(); updatePayloadPreview(); });
    inputs[1].addEventListener("input", () => { item.hex = inputs[1].value; saveSettings(); updatePayloadPreview(); });
    inputs[2].addEventListener("input", () => { item.ratio = inputs[2].value; saveSettings(); updatePayloadPreview(); });
    el.paletteList.appendChild(row);
  });
  saveSettings();
  updatePayloadPreview();
}

function addPaletteRow() {
  state.palette.push({ id: uid(), hex: "#C2D1E6", ratio: "" });
  renderPalette();
}

function buildTemplateRegions(name) {
  return (TEMPLATE_BUILDERS[name] || TEMPLATE_BUILDERS.blank)().map((region) => ({ id: uid(), ...region }));
}

function renderSlideRegionList() {
  if (!state.slideRegions.length) {
    el.slideRegionList.innerHTML = '<p class="hint">当前还没有 PPT 区域，可以先拖拽绘制或应用模板。</p>';
    return;
  }
  el.slideRegionList.innerHTML = "";
  state.slideRegions.forEach((region, index) => {
    const item = document.createElement("div");
    item.className = "region-item";
    item.innerHTML = `
      <div class="region-item-head"><strong>区域 ${index + 1}</strong><button type="button" class="ghost-btn">删除</button></div>
      <input type="text" value="${escapeHtml(region.label)}" />
      <p>位置：${(region.x * 100).toFixed(1)}%，${(region.y * 100).toFixed(1)}%；大小：${(region.w * 100).toFixed(1)}% x ${(region.h * 100).toFixed(1)}%</p>
    `;
    const input = item.querySelector("input");
    input.addEventListener("input", () => {
      region.label = input.value || `区域 ${index + 1}`;
      renderSlideCanvas();
      renderSlideRegionPreview();
      refreshWorkflowPagePrompts();
      saveSettings();
    });
    item.querySelector("button").addEventListener("click", () => {
      state.slideRegions = state.slideRegions.filter((entry) => entry.id !== region.id);
      renderSlidePlanner();
      refreshWorkflowPagePrompts();
      updatePayloadPreview();
    });
    el.slideRegionList.appendChild(item);
  });
}

function renderSlideRegionPreview() {
  el.slideRegionPreview.textContent = buildSlideRegionText() || "// 当前还没有 PPT 区域描述。";
}

function renderSlideCanvas() {
  const canvas = el.slideCanvas;
  const ctx = canvas.getContext("2d");
  const aspect = getCurrentAspectMeta();
  canvas.width = aspect.canvasWidth;
  canvas.height = aspect.canvasHeight;
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.fillStyle = "#fbf6ef";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.strokeStyle = "rgba(19, 34, 56, 0.12)";
  ctx.lineWidth = 1;
  for (let x = 0; x <= canvas.width; x += canvas.width / 8) { ctx.beginPath(); ctx.moveTo(x, 0); ctx.lineTo(x, canvas.height); ctx.stroke(); }
  for (let y = 0; y <= canvas.height; y += canvas.height / 6) { ctx.beginPath(); ctx.moveTo(0, y); ctx.lineTo(canvas.width, y); ctx.stroke(); }
  state.slideRegions.forEach((region, index) => {
    const x = region.x * canvas.width;
    const y = region.y * canvas.height;
    const width = region.w * canvas.width;
    const height = region.h * canvas.height;
    ctx.fillStyle = "rgba(255, 107, 87, 0.16)";
    ctx.strokeStyle = "#ff6b57";
    ctx.lineWidth = 2;
    ctx.fillRect(x, y, width, height);
    ctx.strokeRect(x, y, width, height);
    ctx.fillStyle = "rgba(255, 107, 87, 0.12)";
    ctx.fillRect(x, y, width, Math.min(height, 34));
    ctx.fillStyle = "#132238";
    ctx.font = 'bold 12px "Segoe UI", sans-serif';
    ctx.fillText(region.label || `区域 ${index + 1}`, x + 10, y + 22);
  });
  if (state.slideDraft) {
    ctx.strokeStyle = "#32c8ff";
    ctx.setLineDash([6, 4]);
    ctx.strokeRect(state.slideDraft.x, state.slideDraft.y, state.slideDraft.width, state.slideDraft.height);
    ctx.setLineDash([]);
  }
}

function renderSlideCanvasOverlay() {
  if (!el.slideCanvasOverlay) return;

  el.slideCanvasOverlay.innerHTML = "";
  state.slideRegions.forEach((region, index) => {
    const editor = document.createElement("div");
    editor.className = "region-editor";
    editor.style.left = `${region.x * 100}%`;
    editor.style.top = `${region.y * 100}%`;
    editor.style.width = `${region.w * 100}%`;
    editor.style.height = `${region.h * 100}%`;
    editor.innerHTML = `
      <div class="region-editor-head">
        <span class="region-editor-index">区域 ${index + 1}</span>
        <button type="button" class="ghost-btn region-editor-delete">删</button>
      </div>
      <input type="text" value="${escapeHtml(region.label || `区域 ${index + 1}`)}" placeholder="区域标题" />
      <textarea rows="3" placeholder="直接在框里写这块内容或视觉目标">${escapeHtml(region.content || "")}</textarea>
      <p class="region-editor-meta">${(region.w * 100).toFixed(1)}% x ${(region.h * 100).toFixed(1)}%</p>
    `;

    const titleInput = editor.querySelector("input");
    const contentInput = editor.querySelector("textarea");
    const deleteBtn = editor.querySelector(".region-editor-delete");

    titleInput.addEventListener("input", () => {
      region.label = titleInput.value || `区域 ${index + 1}`;
      renderSlideCanvas();
      renderSlideRegionPreview();
      renderManualLayoutPreview();
      updatePayloadPreview();
      saveSettings();
    });

    contentInput.addEventListener("input", () => {
      region.content = contentInput.value;
      renderSlideRegionPreview();
      renderManualLayoutPreview();
      updatePayloadPreview();
      saveSettings();
    });

    deleteBtn.addEventListener("click", () => {
      state.slideRegions = state.slideRegions.filter((entry) => entry.id !== region.id);
      renderSlidePlanner();
      updatePayloadPreview();
    });

    el.slideCanvasOverlay.appendChild(editor);
  });
}

function renderSlidePlanner() {
  renderSlideCanvas();
  renderSlideCanvasOverlay();
  renderSlideRegionList();
  renderSlideRegionPreview();
  saveSettings();
}

function setupSlideCanvas() {
  const canvas = el.slideCanvas;
  const getCoords = (event) => getCanvasRelativePoint(event, canvas);
  let start = null;
  canvas.onmousedown = (event) => { start = getCoords(event); };
  canvas.onmousemove = (event) => {
    if (!start) return;
    const current = getCoords(event);
    state.slideDraft = { x: Math.min(start.x, current.x), y: Math.min(start.y, current.y), width: Math.abs(current.x - start.x), height: Math.abs(current.y - start.y) };
    renderSlideCanvas();
  };
  canvas.onmouseup = (event) => {
    if (!start) return;
    const end = getCoords(event);
    const x = Math.min(start.x, end.x) / canvas.width;
    const y = Math.min(start.y, end.y) / canvas.height;
    const w = Math.abs(end.x - start.x) / canvas.width;
    const h = Math.abs(end.y - start.y) / canvas.height;
    start = null;
    state.slideDraft = null;
    if (w < 0.03 || h < 0.03) return renderSlideCanvas();
    state.slideRegions.push({ id: uid(), label: `区域 ${state.slideRegions.length + 1}`, content: "", x, y, w, h });
    renderSlidePlanner();
    updatePayloadPreview();
  };
  canvas.onmouseleave = () => {
    if (!start) return;
    start = null;
    state.slideDraft = null;
    renderSlideCanvas();
  };
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error(`读取文件失败：${file.name}`));
    reader.readAsDataURL(file);
  });
}

function buildAutoThemeReferenceLabel() {
  const stamp = new Date();
  const month = String(stamp.getMonth() + 1).padStart(2, "0");
  const day = String(stamp.getDate()).padStart(2, "0");
  const hour = String(stamp.getHours()).padStart(2, "0");
  const minute = String(stamp.getMinutes()).padStart(2, "0");
  return `截图学习风格 ${month}-${day} ${hour}:${minute}`;
}

function renderThemeReferenceAssets() {
  if (!el.themeReferenceHint || !el.themeReferenceList) return;

  if (!state.themeReferenceImages.length) {
    el.themeReferenceHint.textContent = "还没有风格截图";
    el.themeReferenceList.innerHTML = "";
    if (el.learnThemeBtn) el.learnThemeBtn.disabled = true;
    return;
  }

  el.themeReferenceHint.textContent = `已载入 ${state.themeReferenceImages.length} 张风格截图`;
  if (el.learnThemeBtn) el.learnThemeBtn.disabled = false;
  el.themeReferenceList.innerHTML = state.themeReferenceImages.map((item, index) => `
    <article class="theme-reference-card">
      <img src="${escapeHtml(item.source)}" alt="${escapeHtml(item.name)}" />
      <div class="theme-reference-card-body">
        <strong>${escapeHtml(item.name)}</strong>
        <button type="button" class="ghost-btn theme-reference-remove" data-index="${index}">移除</button>
      </div>
    </article>
  `).join("");

  el.themeReferenceList.querySelectorAll(".theme-reference-remove").forEach((button) => {
    button.addEventListener("click", () => {
      const index = Number(button.dataset.index);
      if (!Number.isInteger(index)) return;
      state.themeReferenceImages.splice(index, 1);
      renderThemeReferenceAssets();
    });
  });
}

async function handleThemeReferenceFiles(event) {
  const files = Array.from(event.target.files || []);
  if (!files.length) return;

  const additions = [];
  for (const file of files) {
    const source = await fileToDataUrl(file);
    additions.push({
      id: uid(),
      name: file.name,
      source: String(source),
    });
  }

  state.themeReferenceImages = [...state.themeReferenceImages, ...additions].slice(-6);
  if (el.themeReferenceInput) el.themeReferenceInput.value = "";
  renderThemeReferenceAssets();
  setStatus(`已载入 ${additions.length} 张风格截图，可用于学习风格。`, "success");
}

function getImageDimensions(src) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight });
    img.onerror = () => reject(new Error("图片加载失败"));
    img.src = src;
  });
}

async function addImage({ name, source }) {
  let width = 0;
  let height = 0;
  try { ({ width, height } = await getImageDimensions(source)); } catch {}
  state.images.push({ id: uid(), name, source, width, height, boxes: [] });
  renderImages();
}

function moveImage(id, direction) {
  const index = state.images.findIndex((image) => image.id === id);
  const nextIndex = index + direction;
  if (index < 0 || nextIndex < 0 || nextIndex >= state.images.length) return;
  [state.images[index], state.images[nextIndex]] = [state.images[nextIndex], state.images[index]];
  renderImages();
}

function renderImageCanvas(card, image) {
  const canvas = card.querySelector("canvas");
  const info = card.querySelector(".box-info");
  const ctx = canvas.getContext("2d");
  const preview = new Image();
  const isEditTarget = image.id === state.editTargetImageId;
  preview.onload = () => {
    const dpr = Math.max(1, window.devicePixelRatio || 1);
    const cardWidth = Math.max(240, Math.round(card.clientWidth || 320));
    const maxDisplayWidth = Math.min(560, Math.max(240, cardWidth - 28));
    const maxDisplayHeight = 720;
    const scale = Math.min(
      1,
      maxDisplayWidth / preview.naturalWidth,
      maxDisplayHeight / preview.naturalHeight,
    );
    const displayWidth = Math.max(1, Math.round(preview.naturalWidth * scale));
    const displayHeight = Math.max(1, Math.round(preview.naturalHeight * scale));
    const labelHeight = 20 * dpr;
    const labelWidth = 56 * dpr;
    const labelOffset = 22 * dpr;
    const lineWidth = Math.max(2, Math.round(2 * dpr));
    const fontSize = Math.max(12, Math.round(12 * dpr));

    canvas.style.width = `${displayWidth}px`;
    canvas.style.height = `${displayHeight}px`;
    canvas.width = Math.max(1, Math.round(displayWidth * dpr));
    canvas.height = Math.max(1, Math.round(displayHeight * dpr));
    const draw = (draft) => {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      ctx.imageSmoothingEnabled = true;
      ctx.imageSmoothingQuality = "high";
      ctx.drawImage(preview, 0, 0, canvas.width, canvas.height);
      ctx.lineWidth = lineWidth;
      ctx.strokeStyle = "#ff6b57";
      ctx.fillStyle = "rgba(255, 107, 87, 0.18)";
      image.boxes.forEach((box, index) => {
        const ratioX = canvas.width / image.width;
        const ratioY = canvas.height / image.height;
        const x1 = box[0] * ratioX;
        const y1 = box[1] * ratioY;
        const x2 = box[2] * ratioX;
        const y2 = box[3] * ratioY;
        ctx.fillRect(x1, y1, x2 - x1, y2 - y1);
        ctx.strokeRect(x1, y1, x2 - x1, y2 - y1);
        ctx.fillStyle = "#132238";
        ctx.fillRect(x1, Math.max(0, y1 - labelOffset), labelWidth, labelHeight);
        ctx.fillStyle = "#f7f3eb";
        ctx.font = `${fontSize}px "Segoe UI", sans-serif`;
        ctx.fillText(`框 ${index + 1}`, x1 + (7 * dpr), Math.max(13 * dpr, y1 - (8 * dpr)));
        ctx.fillStyle = "rgba(255, 107, 87, 0.18)";
      });
      if (draft) {
        ctx.strokeStyle = "#32c8ff";
        ctx.setLineDash([6 * dpr, 4 * dpr]);
        ctx.strokeRect(draft.x, draft.y, draft.width, draft.height);
        ctx.setLineDash([]);
      }
    };
    draw();
    if (!el.bboxEnabled.checked) return (info.textContent = "未启用定点改图，当前图片将作为参考图或编辑输入。");
    if (!isEditTarget) return (info.textContent = "当前图片是参考图。定点改图时只在选中的底图上绘制区域。");
    if (!image.width || !image.height) return (info.textContent = "这张图无法读取尺寸，仍可发送，但不能在页面里框选。");
    info.textContent = image.boxes.length >= MAX_BOXES_PER_IMAGE ? "每张底图最多支持 2 个框。点击清框后可重新绘制。" : "当前是改图底图，在图片上拖拽即可绘制定点修改区域。";
    const getCoords = (event) => {
      const point = getCanvasRelativePoint(event, canvas);
      const canvasX = point.x;
      const canvasY = point.y;
      return { canvasX, canvasY, imageX: (canvasX / canvas.width) * image.width, imageY: (canvasY / canvas.height) * image.height };
    };
    let start = null;
    canvas.onmousedown = (event) => { if (image.boxes.length < MAX_BOXES_PER_IMAGE) start = getCoords(event); };
    canvas.onmousemove = (event) => {
      if (!start) return;
      const current = getCoords(event);
      draw({ x: Math.min(start.canvasX, current.canvasX), y: Math.min(start.canvasY, current.canvasY), width: Math.abs(current.canvasX - start.canvasX), height: Math.abs(current.canvasY - start.canvasY) });
    };
    canvas.onmouseup = (event) => {
      if (!start) return;
      const end = getCoords(event);
      const box = [Math.round(clamp(Math.min(start.imageX, end.imageX), 0, image.width)), Math.round(clamp(Math.min(start.imageY, end.imageY), 0, image.height)), Math.round(clamp(Math.max(start.imageX, end.imageX), 0, image.width)), Math.round(clamp(Math.max(start.imageY, end.imageY), 0, image.height))];
      start = null;
      if (box[2] - box[0] < 5 || box[3] - box[1] < 5) return draw();
      image.boxes.push(box);
      renderImages();
    };
    canvas.onmouseleave = () => { if (start) { start = null; draw(); } };
  };
  preview.onerror = () => { info.textContent = "这张图无法在页面中预览，但仍会按 URL 或 Base64 正常发送。"; };
  preview.src = image.source;
}

function renderImages() {
  ensureEditTargetSelection();
  renderEditTargetControls();
  if (!state.images.length) {
    el.imageList.className = "image-list empty";
    el.imageList.innerHTML = "<p>还没有添加图片。文生图可以只填提示词；图像编辑、多图参考和局部框选请先添加图片。</p>";
    updatePayloadPreview();
    return;
  }
  el.imageList.className = "image-list";
  el.imageList.innerHTML = "";
  state.images.forEach((image, index) => {
    const isEditTarget = image.id === state.editTargetImageId;
    const card = document.createElement("article");
    card.className = "image-card";
    card.innerHTML = `
      <div class="image-card-head">
        <div>
          <div class="image-meta">
            <strong>图 ${index + 1}</strong>
            ${isEditTarget ? '<span class="target-badge">改图底图</span>' : '<span class="ref-badge">参考图</span>'}
          </div>
          <p>${escapeHtml(image.name)}</p>
        </div>
        <div class="card-actions">
          <button type="button" class="ghost-btn set-target">${isEditTarget ? "当前底图" : "设为底图"}</button>
          <button type="button" class="ghost-btn clear-boxes">清框</button>
          <button type="button" class="ghost-btn remove-image">删除</button>
        </div>
      </div>
      <canvas></canvas><p class="box-info"></p><div class="box-list"></div>
    `;
    card.querySelector(".set-target").addEventListener("click", () => {
      state.editTargetImageId = image.id;
      saveSettings();
      renderImages();
      setStatus(`已将图 ${index + 1} 设为改图底图。`, "success");
    });
    card.querySelector(".clear-boxes").addEventListener("click", () => { image.boxes = []; renderImages(); });
    card.querySelector(".remove-image").addEventListener("click", () => { state.images = state.images.filter((entry) => entry.id !== image.id); renderImages(); });
    image.boxes.forEach((box, boxIndex) => {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "box-chip";
      chip.textContent = `框 ${boxIndex + 1}: [${box.join(", ")}]`;
      chip.addEventListener("click", () => { image.boxes = image.boxes.filter((_, idx) => idx !== boxIndex); renderImages(); });
      card.querySelector(".box-list").appendChild(chip);
    });
    el.imageList.appendChild(card);
    renderImageCanvas(card, image);
  });
  updatePayloadPreview();
}

function renderEditTargetControls() {
  ensureEditTargetSelection();
  const target = getEditTargetImage();

  if (!state.images.length) {
    el.editTargetSelect.innerHTML = '<option value="">请先上传图片</option>';
    el.editTargetHint.textContent = "选中后，这张图会自动作为最后一张发送，其它图片作为参考图。";
    saveSettings();
    return;
  }

  el.editTargetSelect.innerHTML = state.images
    .map((image, index) => `<option value="${image.id}">图 ${index + 1} - ${escapeHtml(image.name)}</option>`)
    .join("");

  if (state.editTargetImageId) {
    el.editTargetSelect.value = state.editTargetImageId;
  }

  if (!el.bboxEnabled.checked) {
    el.editTargetHint.textContent = "启用定点改图后，可选择哪一张上传图作为 PPT 改图底图。";
    saveSettings();
    return;
  }

  el.editTargetHint.textContent = target
    ? `当前改图底图：${target.name}。发送时它会自动排在最后，其它图片作为参考图。`
    : "启用定点改图后，请先选择一张改图底图。";
  saveSettings();
}

function buildColorPalette() {
  const rows = state.palette
    .map((item) => ({ hex: item.hex.trim(), ratio: item.ratio.trim() }))
    .filter((item) => item.hex && item.ratio);
  if (!rows.length) return null;
  if (rows.length < 3 || rows.length > 10) throw new Error("color_palette 需要 3 到 10 个颜色。");
  const total = rows.reduce((sum, item) => sum + Number(item.ratio.replace("%", "")), 0);
  if (Math.abs(total - 100) > 0.01) throw new Error("color_palette 的 ratio 总和必须为 100.00%。");
  return rows;
}

function buildGenerationParameters({
  n,
  orderedImages = [],
  enableSequential = el.enableSequential.checked,
  enableBbox = el.bboxEnabled.checked,
} = {}) {
  if (isGeminiImageModel()) {
    return {
      size: buildSizeValue(),
    };
  }

  const imageCount = Number(n ?? el.imageCount.value ?? 1);
  if (enableSequential ? imageCount < 1 || imageCount > 12 : imageCount < 1 || imageCount > 4) {
    throw new Error(enableSequential ? "组图模式下 n 的范围应为 1 到 12。" : "非组图模式下 n 的范围应为 1 到 4。");
  }

  const parameters = {
    size: buildSizeValue(),
    n: imageCount,
    watermark: el.watermark.checked,
  };

  if (enableSequential) parameters.enable_sequential = true;
  if (el.seed.value.trim()) parameters.seed = Number(el.seed.value.trim());
  if (el.thinkingMode.checked && !enableSequential && orderedImages.length === 0) {
    parameters.thinking_mode = true;
  }
  if (enableBbox && orderedImages.length) {
    parameters.bbox_list = orderedImages.map((image) => (
      image.id === state.editTargetImageId ? image.boxes : []
    ));
  }
  if (!enableSequential) {
    const palette = buildColorPalette();
    if (palette) parameters.color_palette = palette;
  }

  return parameters;
}

function buildPayload() {
  const orderedImages = getOrderedImagesForPayload();
  const parameters = buildGenerationParameters({
    orderedImages,
    enableSequential: el.enableSequential.checked,
    enableBbox: el.bboxEnabled.checked,
  });
  const promptText = buildPromptText();
  const editText = buildEditTargetText();
  return {
    model: el.model.value,
    input: {
      messages: [{
        role: "user",
        content: [
          ...orderedImages.map((image) => ({ image: image.source })),
          { text: editText ? `${promptText}\n\n${editText}` : promptText },
        ],
      }],
    },
    parameters,
  };
}

function buildWorkflowGenerationPayload(page) {
  return {
    model: el.model.value,
    input: {
      messages: [
        {
          role: "user",
          content: [{ text: page.pagePrompt }],
        },
      ],
    },
    parameters: buildGenerationParameters({
      n: 1,
      orderedImages: [],
      enableSequential: false,
      enableBbox: false,
    }),
  };
}

function renderUsage(response) {
  if (state.activeTabs.main === "revise") {
    el.usageBox.classList.add("hidden");
    el.usageBox.innerHTML = "";
    return;
  }
  const usage = response.usage;
  const output = response.output || {};
  if (!usage && !output.task_status) {
    el.usageBox.classList.add("hidden");
    el.usageBox.innerHTML = "";
    return;
  }
  const rows = [];
  if (response.request_id) rows.push(`<div><span>request_id</span><strong>${escapeHtml(response.request_id)}</strong></div>`);
  if (output.task_id) rows.push(`<div><span>task_id</span><strong>${escapeHtml(output.task_id)}</strong></div>`);
  if (output.task_status) rows.push(`<div><span>task_status</span><strong>${escapeHtml(output.task_status)}</strong></div>`);
  if (usage?.image_count !== undefined) rows.push(`<div><span>image_count</span><strong>${usage.image_count}</strong></div>`);
  if (usage?.size) rows.push(`<div><span>size</span><strong>${escapeHtml(usage.size)}</strong></div>`);
  if (output.submit_time) rows.push(`<div><span>submit_time</span><strong>${escapeHtml(output.submit_time)}</strong></div>`);
  if (output.end_time) rows.push(`<div><span>end_time</span><strong>${escapeHtml(output.end_time)}</strong></div>`);
  el.usageBox.classList.remove("hidden");
  el.usageBox.innerHTML = rows.join("");
}

async function saveResultImage(imageUrl, index) {
  const response = await fetch("/api/download", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ imageUrl, requestId: state.lastRequestId || "result", index: index + 1 }),
  });
  const data = await response.json();
  if (!response.ok || data.code) throw new Error(data.message || "保存到本地失败。");
  state.savedResults[imageUrl] = data;
  renderResultImages();
  return data;
}

function renderResultImages() {
  if (!state.resultImages.length) {
    el.resultImages.innerHTML = "";
    return;
  }
  el.resultImages.innerHTML = "";
  state.resultImages.forEach((imageUrl, index) => {
    const saved = state.savedResults[imageUrl];
    const card = document.createElement("article");
    card.className = "result-card";
    card.innerHTML = `
      <img src="${imageUrl}" alt="生成结果 ${index + 1}" />
      <div class="result-card-foot">
        <div>
          <strong>结果 ${index + 1}</strong>
          <div class="result-status">${saved ? `已保存到：${escapeHtml(saved.fileName)}` : "尚未保存到本地"}</div>
        </div>
        <div class="result-actions">
          <button type="button" class="ghost-btn preview-image-btn">打开原图</button>
          <button type="button" class="ghost-btn use-target-btn">作为改图底图</button>
          <button type="button" class="ghost-btn save-local-btn">下载到本地</button>
          ${saved ? `<a href="${saved.localUrl}" target="_blank" rel="noreferrer" class="ghost-btn">打开本地文件</a>` : ""}
        </div>
      </div>
    `;
    card.querySelector(".preview-image-btn").addEventListener("click", () => {
      openImagePreview({
        title: `结果 ${index + 1} 预览`,
        imageUrl,
      });
    });
    card.querySelector(".use-target-btn").addEventListener("click", async () => {
      await addImage({ name: `结果 ${index + 1}（改图底图）`, source: imageUrl });
      state.editTargetImageId = state.images[state.images.length - 1]?.id || null;
      syncBboxToggles(true);
      saveSettings();
      renderImages();
      setStatus(`结果 ${index + 1} 已加入输入区，并设为改图底图。`, "success");
    });
    card.querySelector(".save-local-btn").addEventListener("click", async () => {
      setStatus(`正在保存结果 ${index + 1} 到 generated-images...`, "running");
      try {
        const data = await saveResultImage(imageUrl, index);
        setStatus(`结果 ${index + 1} 已保存到 ${data.savedPath}`, "success");
      } catch (error) {
        setStatus(error.message || "保存失败。", "error");
      }
    });
    el.resultImages.appendChild(card);
  });
}

function renderResponse(response) {
  if (response.output?.task_id) state.currentTaskId = response.output.task_id;
  if (response.request_id) state.lastRequestId = response.request_id;
  state.resultImages = extractResultImages(response);
  el.responsePreview.textContent = formatJSON(response);
  renderUsage(response);
  renderResultImages();
}

function extractAssistantMessageText(response) {
  const content = response.output?.choices?.[0]?.message?.content;
  if (typeof content === "string") return content;
  if (Array.isArray(content)) return content.map((item) => item.text || "").join("\n").trim();
  return "";
}

function normalizeAssistantResult(parsed) {
  return {
    summary: parsed.summary || "",
    recommendedPrompt: parsed.recommendedPrompt || "",
    editPrompt: parsed.editPrompt || "",
    layoutAdvice: parsed.layoutAdvice || "",
    pptNotes: Array.isArray(parsed.pptNotes) ? parsed.pptNotes : [],
    pagePlan: Array.isArray(parsed.pagePlan) ? parsed.pagePlan : [],
  };
}

function renderAssistantResult() {
  if (!state.assistantParsed) {
    el.assistantSummary.classList.add("hidden");
    el.assistantSummary.innerHTML = "";
    el.assistantPages.innerHTML = "";
    return;
  }
  const data = state.assistantParsed;
  const notes = data.pptNotes.map((note) => `<li>${escapeHtml(note)}</li>`).join("");
  el.assistantSummary.classList.remove("hidden");
  el.assistantSummary.innerHTML = `
    <div class="assistant-grid">
      <div class="assistant-block"><h3>摘要</h3><p>${escapeHtml(data.summary || "暂无")}</p></div>
      <div class="assistant-block"><h3>推荐生图提示词</h3><p>${escapeHtml(data.recommendedPrompt || "暂无")}</p></div>
      <div class="assistant-block"><h3>推荐改图提示词</h3><p>${escapeHtml(data.editPrompt || "暂无")}</p></div>
      <div class="assistant-block"><h3>版式建议</h3><p>${escapeHtml(data.layoutAdvice || "暂无")}</p></div>
    </div>
    ${notes ? `<ul class="assistant-notes">${notes}</ul>` : ""}
  `;
  if (!data.pagePlan.length) {
    el.assistantPages.innerHTML = "";
    return;
  }
  el.assistantPages.innerHTML = data.pagePlan.map((page, index) => `
    <article class="page-card">
      <strong>第 ${page.pageNumber || index + 1} 页：${escapeHtml(page.pageTitle || "未命名页面")}</strong>
      <p><span>视觉目标：</span>${escapeHtml(page.visualGoal || "")}</p>
      <p><span>图片提示词：</span>${escapeHtml(page.imagePrompt || "")}</p>
      <p><span>备注：</span>${escapeHtml(page.notes || "")}</p>
    </article>
  `).join("");
}

function getCurrentThemePreset() {
  return PPT_THEME_PRESETS[el.workflowTheme.value] || PPT_THEME_PRESETS.papercraft;
}

function renderWorkflowThemePreview() {
  const theme = getCurrentThemePreset();
  el.workflowThemePreview.innerHTML = `
    <h3>${escapeHtml(theme.label)}</h3>
    <p>${escapeHtml(theme.summary)}</p>
    <div class="theme-preview-grid">
      <div>
        <strong>封面页</strong>
        <p>${escapeHtml(theme.cover)}</p>
      </div>
      <div>
        <strong>内容页</strong>
        <p>${escapeHtml(theme.content)}</p>
      </div>
      <div>
        <strong>数据页</strong>
        <p>${escapeHtml(theme.data)}</p>
      </div>
    </div>
  `;
}

function normalizeWorkflowPageType(value, index) {
  const type = String(value || "").trim().toLowerCase();
  if (index === 0) return "cover";
  if (type === "cover" || type === "content" || type === "data") return type;
  return "content";
}

function buildWorkflowPagePrompt(page) {
  const theme = getCurrentThemePreset();
  const regionText = el.includeRegionsInPrompt.checked ? buildSlideRegionText() : "";
  return [
    theme.basic,
    theme[page.pageType] || theme.content,
    regionText || "",
    "除非特殊说明，否则画面中的文字请使用与原文一致的语言。",
    `当前页面类型：${PAGE_TYPE_LABELS[page.pageType] || page.pageType}`,
    page.pageTitle ? `当前页面标题：${page.pageTitle}` : "",
    `当前页面内容：\n${page.pageContent || page.pageTitle || ""}`,
  ].filter(Boolean).join("\n\n");
}

function refreshWorkflowPagePrompts() {
  if (!state.workflowPages.length) return;
  resetAllWorkflowPageDesigns("风格或输出尺寸已变化，请重新整理逐页版式。");
  renderWorkflowPlan();
}

function normalizeWorkflowPlan(parsed) {
  const rawPages = Array.isArray(parsed?.pagePlan)
    ? parsed.pagePlan
    : Array.isArray(parsed?.pages)
      ? parsed.pages
      : Array.isArray(parsed)
        ? parsed
        : [];

  const pages = rawPages
    .map((page, index) => {
      const pageNumber = Number(
        page.pageNumber ?? page.page_number ?? index + 1,
      ) || index + 1;
      const pageType = normalizeWorkflowPageType(
        page.pageType ?? page.page_type,
        index,
      );
      const pageTitle = String(
        page.pageTitle
        ?? page.page_title
        ?? (pageType === "cover" ? page.pageContent ?? page.page_content ?? "" : `第 ${pageNumber} 页`),
      ).trim();
      const pageContent = String(
        page.pageContent ?? page.page_content ?? pageTitle,
      ).trim();

      return {
        id: uid(),
        pageNumber,
        pageType,
        pageTitle: pageTitle || `第 ${pageNumber} 页`,
        pageContent,
        layoutSummary: String(page.layoutSummary || page.layout_summary || "").trim(),
        textHierarchy: String(page.textHierarchy || page.text_hierarchy || "").trim(),
        visualFocus: String(page.visualFocus || page.visual_focus || "").trim(),
        readabilityNotes: String(page.readabilityNotes || page.readability_notes || "").trim(),
        pagePrompt: "",
        layoutStatus: "idle",
        layoutError: "",
        layoutPromise: null,
        status: "idle",
        error: "",
        resultImages: [],
        savedResults: {},
        requestId: "",
        taskId: "",
      };
    })
    .filter((page) => page.pageContent || page.pageTitle)
    .sort((a, b) => a.pageNumber - b.pageNumber)
    .map((page, index) => ({
      ...page,
      pageNumber: index + 1,
      pageType: index === 0 ? "cover" : page.pageType,
    }));

  pages.forEach((page) => {
    page.pagePrompt = buildWorkflowPagePrompt(page);
  });

  return {
    summary: String(parsed?.summary || parsed?.planSummary || "").trim(),
    pages,
  };
}

function resetWorkflowPageDesign(page, reason = "版式信息已变化，请重新整理本页。") {
  page.layoutSummary = "";
  page.textHierarchy = "";
  page.visualFocus = "";
  page.readabilityNotes = "";
  page.pagePrompt = "";
  page.layoutStatus = "idle";
  page.layoutError = reason;
  page.layoutPromise = null;
  page.status = "idle";
  page.error = reason;
  page.resultImages = [];
  page.savedResults = {};
  page.detailBackdropUrl = "";
  page.requestId = "";
  page.taskId = "";
}

function resetAllWorkflowPageDesigns(reason) {
  state.workflowPages.forEach((page) => resetWorkflowPageDesign(page, reason));
  renderWorkflowPlan();
}

function renderWorkflowPlanSummary() {
  if (!state.workflowPages.length) {
    el.workflowPlanSummary.classList.add("hidden");
    el.workflowPlanSummary.innerHTML = "";
    return;
  }

  const theme = getCurrentThemePreset();
  const coverCount = state.workflowPages.filter((page) => page.pageType === "cover").length;
  const contentCount = state.workflowPages.filter((page) => page.pageType === "content").length;
  const dataCount = state.workflowPages.filter((page) => page.pageType === "data").length;

  el.workflowPlanSummary.classList.remove("hidden");
  el.workflowPlanSummary.innerHTML = `
    <div class="assistant-grid">
      <div class="assistant-block">
        <h3>拆分页摘要</h3>
        <p>${escapeHtml(state.workflowPlanSummary || "已生成逐页规划。")}</p>
      </div>
      <div class="assistant-block">
        <h3>页面统计</h3>
        <p>共 ${state.workflowPages.length} 页，其中封面 ${coverCount} 页、内容 ${contentCount} 页、数据 ${dataCount} 页。</p>
      </div>
      <div class="assistant-block">
        <h3>当前主题</h3>
        <p>${escapeHtml(theme.label)}：${escapeHtml(theme.summary)}</p>
      </div>
      <div class="assistant-block">
        <h3>批量生成说明</h3>
        <p>批量逐页生图会按当前万相参数顺序执行，每页固定生成 1 张，生成完成后可继续作为改图底图回流到上方。</p>
      </div>
    </div>
  `;
}

async function saveWorkflowPageResult(pageId, imageUrl, index) {
  const page = state.workflowPages.find((item) => item.id === pageId);
  if (!page) throw new Error("未找到对应的页面结果。");

  const response = await fetch("/api/download", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      imageUrl,
      requestId: page.requestId || `page-${page.pageNumber}`,
      index: index + 1,
    }),
  });
  const data = await response.json();
  if (!response.ok || data.code) throw new Error(data.message || "保存到本地失败。");
  page.savedResults[imageUrl] = data;
  if (!page.detailBackdropUrl || /^https?:/i.test(String(page.detailBackdropUrl || "").trim())) {
    page.detailBackdropUrl = String(data.localUrl || imageUrl || "").trim();
  }
  renderWorkflowPlan();
  return data;
}

function renderWorkflowPlan() {
  renderWorkflowPlanSummary();
  el.workflowBatchBtn.disabled = !state.workflowPages.length || state.workflowRunning;
  el.workflowCopyBtn.disabled = !state.workflowPages.length;

  if (!state.workflowPages.length) {
    el.workflowPlanCards.innerHTML = "";
    return;
  }

  el.workflowPlanCards.innerHTML = "";
  state.workflowPages.forEach((page) => {
    const card = document.createElement("article");
    card.className = "workflow-card";

    const tone = page.status === "error"
      ? "error"
      : page.status === "success"
        ? "success"
        : page.status === "running"
          ? "running"
          : "idle";

    card.innerHTML = `
      <div class="workflow-card-head">
        <div>
          <div class="workflow-page-meta">
            <span class="workflow-page-index">第 ${page.pageNumber} 页</span>
            <span class="workflow-type-chip" data-type="${page.pageType}">${PAGE_TYPE_LABELS[page.pageType] || page.pageType}</span>
          </div>
          <h3>${escapeHtml(page.pageTitle || `第 ${page.pageNumber} 页`)}</h3>
        </div>
        <div class="card-actions">
          <button type="button" class="ghost-btn use-workflow-prompt">替换主提示词</button>
          <button type="button" class="ghost-btn copy-workflow-prompt">复制提示词</button>
          <button type="button" class="primary-btn run-workflow-page">生成本页</button>
        </div>
      </div>
      <div class="workflow-status" data-tone="${tone}">
        ${escapeHtml(
          page.status === "running"
            ? (page.error || "正在生成本页...")
            : page.status === "success"
              ? (page.error || "本页生成完成。")
              : page.status === "error"
                ? (page.error || "本页生成失败。")
                : "尚未生成"
        )}
      </div>
      <div class="workflow-section">
        <strong>原文内容</strong>
        <pre>${escapeHtml(page.pageContent || "")}</pre>
      </div>
      <div class="workflow-section">
        <strong>逐页生图提示词</strong>
        <pre>${escapeHtml(page.pagePrompt || "")}</pre>
      </div>
    `;

    const usePromptBtn = card.querySelector(".use-workflow-prompt");
    const copyPromptBtn = card.querySelector(".copy-workflow-prompt");
    const runPageBtn = card.querySelector(".run-workflow-page");

    usePromptBtn.disabled = page.status === "running";
    copyPromptBtn.disabled = page.status === "running";
    runPageBtn.disabled = page.status === "running" || state.workflowRunning;

    usePromptBtn.addEventListener("click", () => {
      el.prompt.value = page.pagePrompt || "";
      saveSettings();
      updatePayloadPreview();
      setStatus(`已将第 ${page.pageNumber} 页提示词写入主提示词。`, "success");
    });

    copyPromptBtn.addEventListener("click", async () => {
      try {
        await navigator.clipboard.writeText(page.pagePrompt || "");
        setStatus(`第 ${page.pageNumber} 页提示词已复制。`, "success");
      } catch (error) {
        setStatus(error.message || "复制提示词失败。", "error");
      }
    });

    runPageBtn.addEventListener("click", () => {
      generateWorkflowPage(page.id);
    });

    if (page.resultImages.length) {
      const resultsWrap = document.createElement("div");
      resultsWrap.className = "result-images";

      page.resultImages.forEach((imageUrl, index) => {
        const saved = page.savedResults[imageUrl];
        const resultCard = document.createElement("article");
        resultCard.className = "result-card";
        resultCard.innerHTML = `
          <img src="${imageUrl}" alt="第 ${page.pageNumber} 页结果 ${index + 1}" />
          <div class="result-card-foot">
            <div>
              <strong>结果 ${index + 1}</strong>
              <div class="result-status">${saved ? `已保存到：${escapeHtml(saved.fileName)}` : "尚未保存到本地"}</div>
            </div>
            <div class="result-actions">
              <a href="${imageUrl}" target="_blank" rel="noreferrer" class="ghost-btn">打开原图</a>
              <button type="button" class="ghost-btn workflow-use-target">作为改图底图</button>
              <button type="button" class="ghost-btn workflow-save-local">下载到本地</button>
              ${saved ? `<a href="${saved.localUrl}" target="_blank" rel="noreferrer" class="ghost-btn">打开本地文件</a>` : ""}
            </div>
          </div>
        `;

        resultCard.querySelector(".workflow-use-target").addEventListener("click", async () => {
          await addImage({
            name: `第 ${page.pageNumber} 页结果 ${index + 1}（改图底图）`,
            source: imageUrl,
          });
          state.editTargetImageId = state.images[state.images.length - 1]?.id || null;
          syncBboxToggles(true);
          saveSettings();
          renderImages();
          setStatus(`第 ${page.pageNumber} 页结果已加入输入区，并设为改图底图。`, "success");
        });

        resultCard.querySelector(".workflow-save-local").addEventListener("click", async () => {
          setStatus(`正在保存第 ${page.pageNumber} 页结果到 generated-images...`, "running");
          try {
            const data = await saveWorkflowPageResult(page.id, imageUrl, index);
            setStatus(`第 ${page.pageNumber} 页结果已保存到 ${data.savedPath}`, "success");
          } catch (error) {
            setStatus(error.message || "保存本页结果失败。", "error");
          }
        });

        resultsWrap.appendChild(resultCard);
      });

      card.appendChild(resultsWrap);
    }

    el.workflowPlanCards.appendChild(card);
  });
}

function buildWorkflowAssistantPayload() {
  const content = el.workflowContent.value.trim();
  const pageCount = Number(el.workflowPageCount.value || 0);

  if (!content) throw new Error("请先粘贴需要拆分的长文本或讲稿。");
  if (!Number.isInteger(pageCount) || pageCount < 2 || pageCount > 20) {
    throw new Error("目标页数请输入 2 到 20 之间的整数。");
  }

  return {
    model: "qwen3.6-plus",
    input: {
      messages: [
        {
          role: "system",
          content: "你是一名专业的 PPT 内容策划师和视觉脚本编辑。你需要把用户原文拆成指定页数的 PPT 页面规划，并严格返回 JSON。除封面标题外，其余页的 pageContent 尽量直接使用用户原文，不要改写、总结或漏掉关键信息。",
        },
        {
          role: "user",
          content: [
            {
              text: [
                `请把下面内容规划成恰好 ${pageCount} 页 PPT。`,
                "硬性规则：",
                `1. 总页数必须严格等于 ${pageCount}。`,
                "2. 第 1 页必须是 cover。",
                "3. cover 页只放标题或主题概括，允许从原文提炼一个简洁标题。",
                "4. 第 2 页及之后的 pageContent 尽量直接使用用户原文，不要总结、不要洗稿、不要重写。",
                "5. pageType 只能是 cover、content、data；包含明确数字、占比、同比、环比、KPI 或统计结果的页优先使用 data。",
                "6. 拆分优先按章节、主题边界、自然段和逻辑单元进行，不要在句子中间硬切。",
                "7. 输出格式必须是 JSON object，字段包含 summary 和 pagePlan。",
                "8. pagePlan 中每一项必须包含 pageNumber、pageType、pageTitle、pageContent。",
                "输出示例：",
                '{"summary":"...","pagePlan":[{"pageNumber":1,"pageType":"cover","pageTitle":"...","pageContent":"..."},{"pageNumber":2,"pageType":"content","pageTitle":"...","pageContent":"..."}]}',
                "用户原文：",
                content,
              ].join("\n\n"),
            },
          ],
        },
      ],
    },
    parameters: {
      result_format: "message",
      response_format: { type: "json_object" },
      enable_thinking: el.assistantThinking.checked,
    },
  };
}

function pickAssistantPrompt() {
  if (!state.assistantParsed) return "";
  return state.assistantParsed.recommendedPrompt || state.assistantParsed.editPrompt || state.assistantParsed.pagePlan?.[0]?.imagePrompt || "";
}

function buildAssistantPayload() {
  const mode = el.assistantMode.value;
  const request = el.assistantRequest.value.trim();
  const currentPrompt = el.assistantUsePrompt.checked ? el.prompt.value.trim() : "";
  const regionText = el.assistantUseRegions.checked ? buildSlideRegionText() : "";
  const editTargetText = buildEditTargetText();
  const includeImages = el.assistantUseImages.checked;
  if (!request && !currentPrompt && !regionText && !(includeImages && state.images.length)) {
    throw new Error("请填写你的目标，或提供主提示词、区域、图片中的至少一种上下文。");
  }
  const content = [];
  if (includeImages) state.images.forEach((image) => content.push({ image: image.source }));
  content.push({
    text: [
      `当前任务模式：${ASSISTANT_MODE_LABELS[mode]}`,
      request ? `用户目标：${request}` : "用户未额外补充目标，请尽量从现有提示词和版式里推断。",
      currentPrompt ? `当前主提示词：${currentPrompt}` : "当前主提示词：无",
      regionText || "当前未提供 PPT 版式区域。",
      editTargetText || "当前未启用定点改图。",
      includeImages && state.images.length ? `已附带 ${state.images.length} 张参考图，请综合主体、构图、风格和色彩。` : "当前未附带参考图。",
      "请只输出 JSON，不要输出 markdown。JSON 字段必须包含：summary、recommendedPrompt、editPrompt、layoutAdvice、pptNotes、pagePlan。",
      "如果当前任务不是多页草案，pagePlan 返回空数组。recommendedPrompt 和 editPrompt 必须能直接用于图片生成或改图。",
    ].join("\n\n"),
  });
  return {
    model: "qwen3.6-plus",
    input: {
      messages: [
        { role: "system", content: "你是一名资深 PPT 视觉设计师、提示词工程师和版式顾问。重点考虑标题留白、文字可读性、主体位置和演示场景统一性。必须只输出 JSON。" },
        { role: "user", content },
      ],
    },
    parameters: {
      result_format: "message",
      response_format: { type: "json_object" },
      enable_thinking: el.assistantThinking.checked,
    },
  };
}

async function fetchTask(taskId) {
  const response = await fetch("/api/tasks/fetch", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ apiKey: el.apiKey.value.trim(), region: el.region.value, taskId }),
  });
  const data = await response.json();
  syncApiKeyFeedback(response, data);
  if (!response.ok) throw new Error(data.message || "查询任务失败。");
  return data;
}

async function requestGeneration(payload) {
  const response = await fetch("/api/generate", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      apiKey: getDashScopeApiKey(),
      googleApiKey: getGoogleApiKey(),
      region: el.region.value,
      asyncMode: shouldUseAsyncImageGeneration(payload?.model),
      slideAspect: el.slideAspect?.value || "16:9",
      payload,
    }),
  });
  const data = await response.json();
  syncApiKeyFeedback(response, data, { model: payload?.model });
  return { response, data };
}

async function waitForTaskCompletion(taskId, onProgress) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < 15 * 60 * 1000) {
    const data = await fetchTask(taskId);
    onProgress?.(data);
    const status = data.output?.task_status;
    if (status === "SUCCEEDED") return data;
    if (status === "FAILED" || status === "CANCELED" || data.code) {
      throw new Error(data.message || `任务 ${taskId} 执行失败。`);
    }
    await new Promise((resolve) => setTimeout(resolve, 4000));
  }
  throw new Error(`任务 ${taskId} 等待超时。`);
}

function startPolling(taskId) {
  stopPolling();
  state.currentTaskId = taskId;
  state.pollTimer = setInterval(async () => {
    try {
      const data = await fetchTask(taskId);
      renderResponse(data);
      const status = data.output?.task_status;
      if (status === "SUCCEEDED") {
        stopPolling();
        setStatus(`任务 ${taskId} 已完成。`, "success");
      } else if (status === "FAILED" || status === "CANCELED" || data.code) {
        stopPolling();
        setStatus(data.message || `任务 ${taskId} 执行失败。`, "error");
      } else {
        setStatus(`任务 ${taskId} 状态：${status || "查询中"}。`, "running");
      }
    } catch (error) {
      stopPolling();
      setStatus(error.message || "轮询失败。", "error");
    }
  }, 4000);
}

async function sendRequest() {
  saveSettings();
  updatePayloadPreview();
  let payload;
  try {
    payload = buildPayload();
  } catch (error) {
    setStatus(error.message, "error");
    return;
  }
  if (!el.apiKey.value.trim()) return setStatus("请先填写 API Key。", "error");
  if (el.bboxEnabled.checked) {
    const target = getEditTargetImage();
    if (!target) {
      return setStatus("你已启用定点改图，但还没有选择改图底图。", "error");
    }
    if (!target.boxes.length) {
      return setStatus("你已启用定点改图，但当前底图还没有绘制任何 bbox 区域。", "error");
    }
  }
  setStatus("正在调用 DashScope 图片模型...", "running");
  try {
    const { response, data } = await requestGeneration(payload);
    renderResponse(data);
    if (!response.ok || data.code) {
      stopPolling();
      state.currentTaskId = null;
      return setStatus(data.message || "调用失败。", "error");
    }
    if (el.requestMode.value === "async") {
      const taskId = data.output?.task_id;
      if (!taskId) return setStatus("异步任务创建成功，但没有返回 task_id。", "error");
      setStatus(`任务已创建，task_id：${taskId}，开始轮询。`, "running");
      return startPolling(taskId);
    }
    stopPolling();
    state.currentTaskId = null;
    setStatus("同步调用完成。", "success");
  } catch (error) {
    stopPolling();
    state.currentTaskId = null;
    setStatus(error.message || "请求失败。", "error");
  }
}

async function refreshCurrentTask() {
  if (!state.currentTaskId) return setStatus("当前没有可刷新的任务。", "error");
  setStatus(`正在刷新任务 ${state.currentTaskId}...`, "running");
  try {
    const data = await fetchTask(state.currentTaskId);
    renderResponse(data);
    setStatus(`任务 ${state.currentTaskId} 状态：${data.output?.task_status || "未知"}`, "idle");
  } catch (error) {
    setStatus(error.message || "刷新失败。", "error");
  }
}

async function downloadAllResults() {
  if (!state.resultImages.length) return setStatus("当前没有可下载的结果图片。", "error");
  setStatus("正在保存全部结果图片到 generated-images...", "running");
  try {
    for (let index = 0; index < state.resultImages.length; index += 1) {
      const imageUrl = state.resultImages[index];
      if (!state.savedResults[imageUrl]) await saveResultImage(imageUrl, index);
    }
    setStatus("全部结果图片已保存到项目根目录的 generated-images 文件夹。", "success");
  } catch (error) {
    setStatus(error.message || "批量下载失败。", "error");
  }
}

async function handleFiles(event) {
  const files = Array.from(event.target.files || []);
  for (const file of files) {
    const source = await fileToDataUrl(file);
    await addImage({ name: `${file.name}（本地）`, source });
  }
  event.target.value = "";
}

async function handleAddImageUrl() {
  const url = el.imageUrlInput.value.trim();
  if (!url) return setStatus("请输入图片 URL。", "error");
  await addImage({ name: `${url}（URL）`, source: url });
  el.imageUrlInput.value = "";
  setStatus("URL 图片已添加。", "success");
}

async function copyPayload() {
  try {
    await navigator.clipboard.writeText(formatJSON(buildPayload()));
    setStatus("请求 JSON 已复制。", "success");
  } catch (error) {
    setStatus(error.message || "复制失败。", "error");
  }
}

async function sendAssistantRequest() {
  let payload;
  try {
    payload = buildAssistantPayload();
  } catch (error) {
    setStatus(error.message, "error");
    return;
  }
  if (!el.apiKey.value.trim()) return setStatus("请先填写 API Key。", "error");
  setStatus("正在调用 Qwen 提示词助手...", "running");
  el.assistantPreview.textContent = formatJSON(sanitizeForPreview(payload));
  try {
    const response = await fetch("/api/assistant", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ apiKey: el.apiKey.value.trim(), region: el.region.value, payload }),
    });
    const data = await response.json();
    if (!response.ok || data.code) {
      state.assistantParsed = null;
      state.assistantRaw = formatJSON(data);
      el.assistantPreview.textContent = state.assistantRaw;
      renderAssistantResult();
      return setStatus(data.message || "Qwen 助手调用失败。", "error");
    }
    const rawText = extractAssistantMessageText(data);
    state.assistantRaw = rawText || formatJSON(data);
    el.assistantPreview.textContent = state.assistantRaw;
    try {
      state.assistantParsed = normalizeAssistantResult(JSON.parse(rawText));
    } catch {
      state.assistantParsed = null;
    }
    renderAssistantResult();
    setStatus("Qwen 提示词建议已生成。", "success");
  } catch (error) {
    state.assistantParsed = null;
    renderAssistantResult();
    setStatus(error.message || "Qwen 助手请求失败。", "error");
  }
}

async function planWorkflowPages() {
  let payload;
  try {
    payload = buildWorkflowAssistantPayload();
  } catch (error) {
    setStatus(error.message, "error");
    return;
  }
  if (!el.apiKey.value.trim()) return setStatus("请先填写 API Key。", "error");

  setStatus("正在调用 Qwen 拆分页并生成逐页提示词...", "running");
  state.workflowPlanRaw = formatJSON(sanitizeForPreview(payload));

  try {
    const response = await fetch("/api/assistant", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        apiKey: el.apiKey.value.trim(),
        region: el.region.value,
        payload,
      }),
    });
    const data = await response.json();
    if (!response.ok || data.code) {
      state.workflowPlanSummary = "";
      state.workflowPages = [];
      renderWorkflowPlan();
      return setStatus(data.message || "PPT 拆分页调用失败。", "error");
    }

    const rawText = extractAssistantMessageText(data);
    state.workflowPlanRaw = rawText || formatJSON(data);

    let normalized;
    try {
      normalized = normalizeWorkflowPlan(JSON.parse(rawText));
    } catch {
      normalized = { summary: "", pages: [] };
    }

    if (!normalized.pages.length) {
      state.workflowPlanSummary = "";
      state.workflowPages = [];
      renderWorkflowPlan();
      return setStatus("拆分页结果无法解析，请调整原文后重试。", "error");
    }

    state.workflowPlanSummary = normalized.summary || "已根据原文生成逐页规划。";
    state.workflowPages = normalized.pages;
    renderWorkflowPlan();
    setStatus(`拆分页完成，已生成 ${state.workflowPages.length} 页的逐页提示词。`, "success");
  } catch (error) {
    state.workflowPlanSummary = "";
    state.workflowPages = [];
    renderWorkflowPlan();
    setStatus(error.message || "PPT 拆分页请求失败。", "error");
  }
}

async function copyWorkflowPlan() {
  if (!state.workflowPages.length) return setStatus("当前没有可复制的拆分页结果。", "error");
  try {
    const data = state.workflowPages.map((page) => ({
      pageNumber: page.pageNumber,
      pageType: page.pageType,
      pageTitle: page.pageTitle,
      pageContent: page.pageContent,
      pagePrompt: page.pagePrompt,
    }));
    await navigator.clipboard.writeText(formatJSON(data));
    setStatus("拆分页 JSON 已复制。", "success");
  } catch (error) {
    setStatus(error.message || "复制拆分页 JSON 失败。", "error");
  }
}

function clearWorkflowPlan() {
  state.workflowPlanSummary = "";
  state.workflowPlanRaw = "";
  state.workflowPages = [];
  state.workflowPlanLibraryActiveId = "";
  state.workflowRunning = false;
  state.workflowDetailPageId = null;
  el.workflowDetailModal?.classList.add("hidden");
  el.workflowDetailModal?.setAttribute("aria-hidden", "true");
  renderWorkflowPlanLibrary();
  saveSettings();
  renderWorkflowPlan();
  setStatus("已清空拆分页结果。", "success");
}

async function generateWorkflowPage(pageId, options = {}) {
  const { stopOnError = false } = options;
  const page = state.workflowPages.find((item) => item.id === pageId);
  if (!page) return setStatus("未找到对应页面。", "error");
  if (!el.apiKey.value.trim()) return setStatus("请先填写 API Key。", "error");

  page.status = "running";
  page.error = "正在生成本页...";
  renderWorkflowPlan();
  setStatus(`正在生成第 ${page.pageNumber} 页...`, "running");

  try {
    if (workflowPageNeedsDesign(page)) {
      setStatus(`正在整理第 ${page.pageNumber} 页的版式和提示词...`, "running");
      await ensureWorkflowPageDesign(page, {
        force: page.layoutStatus === "error" || !hasWorkflowStructuredLayout(page),
      });
    }

    const payload = buildWorkflowGenerationPayload(page);
    const { response, data } = await requestGeneration(payload);

    if (!response.ok || data.code) {
      page.status = "error";
      page.error = data.message || "本页生成失败。";
      renderWorkflowPlan();
      setStatus(page.error, "error");
      if (stopOnError) throw new Error(page.error);
      return;
    }

    let finalData = data;
    if (el.requestMode.value === "async") {
      page.taskId = data.output?.task_id || "";
      if (!page.taskId) {
        throw new Error(`第 ${page.pageNumber} 页已创建异步任务，但没有返回 task_id。`);
      }
      finalData = await waitForTaskCompletion(page.taskId, (taskData) => {
        page.error = `任务状态：${taskData.output?.task_status || "处理中"}`;
        renderWorkflowPlan();
      });
    }

    page.requestId = finalData.request_id || data.request_id || "";
    page.resultImages = extractResultImages(finalData);
    page.detailBackdropUrl = page.resultImages[0] || "";
    try {
      await ensureWorkflowPageLocalCache(page);
    } catch (cacheError) {
      console.warn("Failed to cache workflow page result locally:", cacheError);
    }
    page.detailBackdropUrl = getWorkflowPreferredImageUrl(page, page.resultImages[0]) || "";
    page.status = "success";
    saveSettings();
    page.error = page.resultImages.length
      ? `本页已生成 ${page.resultImages.length} 张结果图。`
      : "本页已完成，但没有提取到图片结果。";
    renderWorkflowPlan();
    setStatus(`第 ${page.pageNumber} 页生成完成。`, "success");
  } catch (error) {
    page.status = "error";
    page.error = error.message || "本页生成失败。";
    renderWorkflowPlan();
    setStatus(page.error, "error");
    if (stopOnError) throw error;
  }
}

async function generateAllWorkflowPages() {
  if (!state.workflowPages.length) return setStatus("请先完成拆分页。", "error");
  if (state.workflowRunning) return setStatus("当前已经有批量逐页任务在执行。", "error");

  state.workflowRunning = true;
  renderWorkflowPlan();

  try {
    for (let index = 0; index < state.workflowPages.length; index += 1) {
      const page = state.workflowPages[index];
      setStatus(`批量逐页生图中：第 ${index + 1}/${state.workflowPages.length} 页`, "running");
      await generateWorkflowPage(page.id, { stopOnError: true });
    }
    setStatus("批量逐页生图已完成。", "success");
  } catch (error) {
    setStatus(error.message || "批量逐页生图已中断。", "error");
  } finally {
    state.workflowRunning = false;
    renderWorkflowPlan();
  }
}

function applyAssistantPrompt(replace) {
  const prompt = pickAssistantPrompt();
  if (!prompt) return setStatus("当前没有可应用的提示词建议。", "error");
  el.prompt.value = replace ? prompt : `${el.prompt.value.trim()}${el.prompt.value.trim() ? "\n\n" : ""}${prompt}`;
  saveSettings();
  updatePayloadPreview();
  setStatus(replace ? "已替换主提示词。" : "已追加到主提示词。", "success");
}

function createImageEntry({ name, source, width = 0, height = 0 }) {
  return { id: uid(), name, source, width, height, boxes: [], history: [] };
}

function getImageSnapshot(image) {
  return {
    name: image.name,
    source: image.source,
    width: image.width,
    height: image.height,
    boxes: Array.isArray(image.boxes) ? image.boxes.map((box) => [...box]) : [],
  };
}

async function addImage({ name, source }) {
  let width = 0;
  let height = 0;
  try { ({ width, height } = await getImageDimensions(source)); } catch {}
  state.images.push(createImageEntry({ name, source, width, height }));
  renderImages();
}

function buildTemplateRegions(name) {
  return (TEMPLATE_BUILDERS[name] || TEMPLATE_BUILDERS.blank)().map((region) => ({
    id: uid(),
    content: "",
    ...region,
  }));
}

function buildSlideRegionText() {
  if (!state.slideRegions.length) return "";
  const aspect = getCurrentAspectMeta();
  const lines = [
    `PPT版式比例：${aspect.label}（建议输出尺寸 ${aspect.outputWidth}x${aspect.outputHeight}）`,
    "PPT版式区域要求：",
  ];

  state.slideRegions.forEach((region, index) => {
    const x = (region.x * 100).toFixed(1);
    const y = (region.y * 100).toFixed(1);
    const w = (region.w * 100).toFixed(1);
    const h = (region.h * 100).toFixed(1);
    const label = region.label || `区域 ${index + 1}`;
    lines.push(`${index + 1}. ${label}：位置 ${x}%/${y}%，大小 ${w}% x ${h}%`);
    if (region.content?.trim()) {
      lines.push(`   区域文案：${region.content.trim()}`);
    }
  });

  lines.push("请根据这些区域安排主体位置、留白和文字可读性，不要让重要主体压住标题和正文。");
  return lines.join("\n");
}

function buildManualPromptDraft() {
  const theme = getCurrentThemePreset();
  const pageGoal = el.manualPageGoal.value.trim();
  const regionText = buildSlideRegionText();
  const pageType = state.slideRegions.length > 2 ? "content" : "cover";

  return [
    theme.basic,
    theme[pageType] || theme.content,
    pageGoal ? `当前页面目标：${pageGoal}` : "",
    regionText,
    "请把版式秩序、文本留白和视觉层级一起考虑进去，优先保证标题和正文可读性。",
  ].filter(Boolean).join("\n\n");
}

function renderManualLayoutPreview() {
  if (!el.manualLayoutPreview) return;
  const preview = [
    `当前主题：${getCurrentThemePreset().label}`,
    el.manualPageGoal.value.trim() ? `页面目标：${el.manualPageGoal.value.trim()}` : "",
    buildSlideRegionText(),
  ].filter(Boolean).join("\n\n");
  el.manualLayoutPreview.textContent = preview || "// 先画出区域，或者写下这一页的表达目标。";
}

function renderSlideRegionList() {
  if (!state.slideRegions.length) {
    el.slideRegionList.innerHTML = '<p class="hint">当前还没有 PPT 区域，可以先拖拽绘制或应用模板。</p>';
    return;
  }

  el.slideRegionList.innerHTML = "";
  state.slideRegions.forEach((region, index) => {
    const item = document.createElement("div");
    item.className = "region-item";
    item.innerHTML = `
      <div class="region-item-head">
        <strong>区域 ${index + 1}</strong>
        <button type="button" class="ghost-btn">删除</button>
      </div>
      <input type="text" value="${escapeHtml(region.label || `区域 ${index + 1}`)}" placeholder="区域名称" />
      <textarea rows="3" placeholder="填写这个区域预计放什么文案或视觉目标">${escapeHtml(region.content || "")}</textarea>
      <p>位置：${(region.x * 100).toFixed(1)}%，${(region.y * 100).toFixed(1)}%；大小：${(region.w * 100).toFixed(1)}% x ${(region.h * 100).toFixed(1)}%</p>
    `;

    const labelInput = item.querySelector("input");
    const contentInput = item.querySelector("textarea");

    labelInput.addEventListener("input", () => {
      region.label = labelInput.value || `区域 ${index + 1}`;
      renderSlideCanvas();
      renderSlideRegionPreview();
      renderManualLayoutPreview();
      saveSettings();
    });

    contentInput.addEventListener("input", () => {
      region.content = contentInput.value;
      renderSlideRegionPreview();
      renderManualLayoutPreview();
      saveSettings();
    });

    item.querySelector("button").addEventListener("click", () => {
      state.slideRegions = state.slideRegions.filter((entry) => entry.id !== region.id);
      renderSlidePlanner();
      renderManualLayoutPreview();
      updatePayloadPreview();
    });

    el.slideRegionList.appendChild(item);
  });
}

function getOrderedImagesForPayload() {
  const target = getEditTargetImage();
  if (state.activeTabs.main === "revise" && target) {
    const refs = state.images.filter((image) => image.id !== target.id);
    return [...refs, target];
  }
  if (!el.bboxEnabled.checked) {
    return state.images;
  }
  if (!target) {
    throw new Error("启用定点改图时，请先选择一张改图底图。");
  }
  const refs = state.images.filter((image) => image.id !== target.id);
  return [...refs, target];
}

function buildEditTargetText() {
  const target = getEditTargetImage();
  if (!target) return "";

  if (!el.bboxEnabled.checked) {
    if (state.activeTabs.main !== "revise") return "";
    return [
      "改图底稿说明：",
      `当前需要修改的底稿是最后一张图：${target.name}`,
      "其余图片是叠加素材或参考图。请优先保留原有主体关系、构图和关键信息，再完成修改与美化；如果底稿是 PPT，也尽量保持版式稳定。",
    ].join("\n");
  }

  const boxLines = target.boxes.map((box, index) => `区域 ${index + 1}: [${box.join(", ")}]`).join("\n");
  return [
    "定点改图要求：",
    "请以最后一张底图为基础，只修改框选区域，其他主体结构、构图关系、文字排版（如有）和非框选内容尽量保持稳定。",
    `改图底图：${target.name}`,
    boxLines || "当前还没有框选区域。",
  ].join("\n");
}

function renderEditTargetControls() {
  ensureEditTargetSelection();
  const target = getEditTargetImage();
  const historyCount = target?.history?.length || 0;

  if (el.undoTargetBtn) {
    el.undoTargetBtn.disabled = !historyCount;
  }

  if (!state.images.length) {
    el.editTargetSelect.innerHTML = '<option value="">请先上传图片</option>';
    el.editTargetHint.textContent = "上传后左侧就是改图画板。选中的底图会作为本次修改对象，其余图片会当参考图一起提交。";
    saveSettings();
    return;
  }

  el.editTargetSelect.innerHTML = state.images
    .map((image, index) => `<option value="${image.id}">图 ${index + 1} - ${escapeHtml(image.name)}</option>`)
    .join("");

  if (state.editTargetImageId) {
    el.editTargetSelect.value = state.editTargetImageId;
  }

  if (!target) {
    el.editTargetHint.textContent = "请先选择当前要修改的底稿。";
    saveSettings();
    return;
  }

  if (state.activeTabs.main === "revise" && !el.bboxEnabled.checked) {
    el.editTargetHint.textContent = `当前底稿：${target.name}。现在可以直接在右侧写修改提示词提交；如果需要局部改，再打开框选。历史可回退 ${historyCount} 步。`;
  } else if (!el.bboxEnabled.checked) {
    el.editTargetHint.textContent = "启用定点改图后，可在这里指定哪一张图片作为当前底稿。";
  } else {
    el.editTargetHint.textContent = `当前框选底稿：${target.name}。左侧拖框即可指定局部修改区域。历史可回退 ${historyCount} 步。`;
  }

  saveSettings();
}

async function replaceEditTargetWithLatestResult() {
  if (state.activeTabs.main !== "revise" || !el.autoReplaceTarget.checked || !state.resultImages.length) return;
  const target = getEditTargetImage();
  const latest = state.resultImages[0];
  if (!target || !latest || target.source === latest) return;

  target.history = Array.isArray(target.history) ? target.history : [];
  target.history.push(getImageSnapshot(target));

  let width = target.width;
  let height = target.height;
  try { ({ width, height } = await getImageDimensions(latest)); } catch {}

  target.source = latest;
  target.width = width;
  target.height = height;
  target.boxes = [];
  target.name = `${target.name.replace(/（最新结果）$/, "")}（最新结果）`;

  renderImages();
  setStatus("最新结果已自动替换当前底稿，你可以继续下一轮修改。", "success");
}

function undoEditTargetVersion() {
  const target = getEditTargetImage();
  if (!target || !target.history?.length) {
    setStatus("当前没有可回退的底稿历史。", "error");
    return;
  }

  const previous = target.history.pop();
  target.name = previous.name;
  target.source = previous.source;
  target.width = previous.width;
  target.height = previous.height;
  target.boxes = previous.boxes.map((box) => [...box]);

  renderImages();
  setStatus("已回到上一版底稿。", "success");
}

function renderResponse(response) {
  if (response.output?.task_id) state.currentTaskId = response.output.task_id;
  if (response.request_id) state.lastRequestId = response.request_id;
  state.resultImages = extractResultImages(response);
  el.responsePreview.textContent = formatJSON(response);
  renderUsage(response);
  renderResultImages();
  replaceEditTargetWithLatestResult();
}

function bindEvents() {
  [
    el.apiKey, el.requestMode, el.region, el.model, el.prompt, el.presetSize, el.imageCount, el.seed, el.enableSequential,
    el.thinkingMode, el.watermark, el.bboxEnabled, el.customWidth, el.customHeight, el.slideAspect, el.slideTemplate,
    el.includeRegionsInPrompt, el.assistantMode, el.assistantUseImages, el.assistantUseRegions, el.assistantUsePrompt,
    el.assistantThinking, el.assistantRequest, el.workflowPageCount, el.workflowTheme, el.workflowContent, el.manualPageGoal,
  ].forEach((node) => {
    node.addEventListener("input", () => {
      saveSettings();
      updatePayloadPreview();
      if (node === el.manualPageGoal) renderManualLayoutPreview();
    });
    node.addEventListener("change", () => {
      saveSettings();
      updatePayloadPreview();
      if (node === el.bboxEnabled) renderImages();
      if (node === el.slideAspect) {
        renderSlidePlanner();
        refreshWorkflowPagePrompts();
        renderManualLayoutPreview();
      }
      if (node === el.includeRegionsInPrompt) {
        refreshWorkflowPagePrompts();
        renderManualLayoutPreview();
      }
      if (node === el.workflowTheme) {
        renderWorkflowThemePreview();
        refreshWorkflowPagePrompts();
        renderManualLayoutPreview();
      }
      if (node === el.manualPageGoal) renderManualLayoutPreview();
    });
  });

  el.autoReplaceTarget.addEventListener("change", saveSettings);
  el.bboxEnabled.addEventListener("change", () => {
    syncBboxToggles(el.bboxEnabled.checked);
    saveSettings();
    renderEditTargetControls();
    renderImages();
  });
  el.bboxEnabledImagePanel.addEventListener("change", () => {
    syncBboxToggles(el.bboxEnabledImagePanel.checked);
    saveSettings();
    renderEditTargetControls();
    renderImages();
    updatePayloadPreview();
  });
  el.editTargetSelect.addEventListener("change", () => {
    state.editTargetImageId = el.editTargetSelect.value || null;
    saveSettings();
    renderEditTargetControls();
    renderImages();
    updatePayloadPreview();
  });
  el.undoTargetBtn.addEventListener("click", undoEditTargetVersion);
  el.sizeMode.addEventListener("change", () => { saveSettings(); toggleSizeMode(); updatePayloadPreview(); });
  el.fileInput.addEventListener("change", handleFiles);
  el.addImageUrlBtn.addEventListener("click", handleAddImageUrl);
  el.addPaletteBtn.addEventListener("click", addPaletteRow);
  el.sendBtn.addEventListener("click", sendRequest);
  el.refreshTaskBtn.addEventListener("click", refreshCurrentTask);
  el.copyPayloadBtn.addEventListener("click", copyPayload);
  el.downloadAllBtn.addEventListener("click", downloadAllResults);
  el.applyTemplateBtn.addEventListener("click", () => {
    state.slideRegions = buildTemplateRegions(el.slideTemplate.value);
    renderSlidePlanner();
    refreshWorkflowPagePrompts();
    renderManualLayoutPreview();
    updatePayloadPreview();
    setStatus("已应用 PPT 版式模板。", "success");
  });
  el.syncSlideSizeBtn.addEventListener("click", () => {
    const aspect = getCurrentAspectMeta();
    el.sizeMode.value = "custom";
    el.customWidth.value = String(aspect.outputWidth);
    el.customHeight.value = String(aspect.outputHeight);
    toggleSizeMode();
    saveSettings();
    updatePayloadPreview();
    setStatus(`已同步为 ${aspect.outputWidth}x${aspect.outputHeight}。`, "success");
  });
  el.insertRegionsBtn.addEventListener("click", () => {
    const text = buildSlideRegionText();
    if (!text) return setStatus("当前还没有 PPT 区域可插入。", "error");
    const current = el.prompt.value.trim();
    el.prompt.value = current ? `${current}\n\n${text}` : text;
    saveSettings();
    refreshWorkflowPagePrompts();
    renderManualLayoutPreview();
    updatePayloadPreview();
    setStatus("已将区域说明追加到主提示词。", "success");
  });
  el.clearRegionsBtn.addEventListener("click", () => {
    state.slideRegions = [];
    renderSlidePlanner();
    refreshWorkflowPagePrompts();
    renderManualLayoutPreview();
    updatePayloadPreview();
    setStatus("已清空 PPT 区域。", "success");
  });
  el.manualBuildPromptBtn.addEventListener("click", () => {
    const prompt = buildManualPromptDraft();
    if (!prompt) return setStatus("请先填写页面目标或区域。", "error");
    el.prompt.value = prompt;
    saveSettings();
    updatePayloadPreview();
    setStatus("已按当前主题和版式生成主提示词骨架。", "success");
  });
  el.manualAppendPromptBtn.addEventListener("click", () => {
    const prompt = buildManualPromptDraft();
    if (!prompt) return setStatus("请先填写页面目标或区域。", "error");
    el.prompt.value = `${el.prompt.value.trim()}${el.prompt.value.trim() ? "\n\n" : ""}${prompt}`;
    saveSettings();
    updatePayloadPreview();
    setStatus("已把当前版式骨架追加到主提示词。", "success");
  });
  el.manualCopyBriefBtn.addEventListener("click", async () => {
    try {
      await navigator.clipboard.writeText(el.manualLayoutPreview.textContent || "");
      setStatus("排版摘要已复制。", "success");
    } catch (error) {
      setStatus(error.message || "复制排版摘要失败。", "error");
    }
  });
  el.assistantSendBtn.addEventListener("click", sendAssistantRequest);
  el.assistantApplyBtn.addEventListener("click", () => applyAssistantPrompt(true));
  el.assistantAppendBtn.addEventListener("click", () => applyAssistantPrompt(false));
  el.workflowPlanBtn.addEventListener("click", planWorkflowPages);
  el.workflowPlanStopBtn?.addEventListener("click", cancelWorkflowPlanRun);
  el.workflowBatchBtn.addEventListener("click", generateAllWorkflowPages);
  el.workflowCopyBtn.addEventListener("click", copyWorkflowPlan);
  el.workflowClearBtn.addEventListener("click", clearWorkflowPlan);
}

function init() {
  loadSettings();
  syncBboxToggles(el.bboxEnabled.checked);
  toggleSizeMode();
  bindTabNavigation();
  bindEvents();
  setupSlideCanvas();
  renderPalette();
  renderSlidePlanner();
  renderWorkflowThemePreview();
  renderWorkflowPlan();
  renderImages();
  renderAssistantResult();
  renderManualLayoutPreview();
  applyTabState();
  updatePayloadPreview();
  setStatus("页面已就绪，可以开始配置、拆页和改图。");
}

function bindEvents() {
  [
    el.apiKey, el.requestMode, el.region, el.model, el.prompt, el.presetSize, el.imageCount, el.seed, el.enableSequential,
    el.thinkingMode, el.watermark, el.bboxEnabled, el.customWidth, el.customHeight, el.slideAspect, el.slideTemplate,
    el.includeRegionsInPrompt, el.assistantMode, el.assistantUseImages, el.assistantUseRegions, el.assistantUsePrompt,
    el.assistantThinking, el.assistantRequest, el.workflowPageCount, el.workflowTheme, el.workflowContent,
  ].forEach((node) => {
    node.addEventListener("input", () => { saveSettings(); updatePayloadPreview(); });
    node.addEventListener("change", () => {
      saveSettings();
      updatePayloadPreview();
      if (node === el.bboxEnabled) renderImages();
      if (node === el.slideAspect) {
        renderSlidePlanner();
        refreshWorkflowPagePrompts();
      }
      if (node === el.includeRegionsInPrompt) refreshWorkflowPagePrompts();
      if (node === el.workflowTheme) {
        renderWorkflowThemePreview();
        refreshWorkflowPagePrompts();
      }
    });
  });
  el.bboxEnabled.addEventListener("change", () => {
    syncBboxToggles(el.bboxEnabled.checked);
    saveSettings();
    renderEditTargetControls();
    renderImages();
  });
  el.bboxEnabledImagePanel.addEventListener("change", () => {
    syncBboxToggles(el.bboxEnabledImagePanel.checked);
    saveSettings();
    renderEditTargetControls();
    renderImages();
    updatePayloadPreview();
  });
  el.editTargetSelect.addEventListener("change", () => {
    state.editTargetImageId = el.editTargetSelect.value || null;
    saveSettings();
    renderEditTargetControls();
    renderImages();
    updatePayloadPreview();
  });
  el.sizeMode.addEventListener("change", () => { saveSettings(); toggleSizeMode(); updatePayloadPreview(); });
  el.fileInput.addEventListener("change", handleFiles);
  el.addImageUrlBtn.addEventListener("click", handleAddImageUrl);
  el.addPaletteBtn.addEventListener("click", addPaletteRow);
  el.sendBtn.addEventListener("click", sendRequest);
  el.refreshTaskBtn.addEventListener("click", refreshCurrentTask);
  el.copyPayloadBtn.addEventListener("click", copyPayload);
  el.downloadAllBtn.addEventListener("click", downloadAllResults);
  el.applyTemplateBtn.addEventListener("click", () => {
    state.slideRegions = buildTemplateRegions(el.slideTemplate.value);
    renderSlidePlanner();
    refreshWorkflowPagePrompts();
    updatePayloadPreview();
    setStatus("已应用 PPT 版式模板。", "success");
  });
  el.syncSlideSizeBtn.addEventListener("click", () => {
    const aspect = getCurrentAspectMeta();
    el.sizeMode.value = "custom";
    el.customWidth.value = String(aspect.outputWidth);
    el.customHeight.value = String(aspect.outputHeight);
    toggleSizeMode();
    saveSettings();
    updatePayloadPreview();
    setStatus(`已同步为 ${aspect.outputWidth}x${aspect.outputHeight}。`, "success");
  });
  el.insertRegionsBtn.addEventListener("click", () => {
    const text = buildSlideRegionText();
    if (!text) return setStatus("当前还没有 PPT 区域可插入。", "error");
    const current = el.prompt.value.trim();
    el.prompt.value = current ? `${current}\n\n${text}` : text;
    saveSettings();
    refreshWorkflowPagePrompts();
    updatePayloadPreview();
    setStatus("已将区域描述追加到主提示词。", "success");
  });
  el.clearRegionsBtn.addEventListener("click", () => {
    state.slideRegions = [];
    renderSlidePlanner();
    refreshWorkflowPagePrompts();
    updatePayloadPreview();
    setStatus("已清空 PPT 区域。", "success");
  });
  el.assistantSendBtn.addEventListener("click", sendAssistantRequest);
  el.assistantApplyBtn.addEventListener("click", () => applyAssistantPrompt(true));
  el.assistantAppendBtn.addEventListener("click", () => applyAssistantPrompt(false));
  el.workflowPlanBtn.addEventListener("click", planWorkflowPages);
  el.workflowBatchBtn.addEventListener("click", generateAllWorkflowPages);
  el.workflowCopyBtn.addEventListener("click", copyWorkflowPlan);
  el.workflowClearBtn.addEventListener("click", clearWorkflowPlan);
}

function init() {
  loadSettings();
  syncBboxToggles(el.bboxEnabled.checked);
  toggleSizeMode();
  bindTabNavigation();
  bindEvents();
  setupSlideCanvas();
  renderPalette();
  renderSlidePlanner();
  renderWorkflowThemePreview();
  renderWorkflowPlan();
  renderImages();
  renderAssistantResult();
  applyTabState();
  updatePayloadPreview();
  setStatus("页面已就绪，可以开始配置 PPT 版式与图片需求。");
}

function bindEvents() {
  [
    el.apiKey, el.requestMode, el.region, el.model, el.prompt, el.presetSize, el.imageCount, el.seed, el.enableSequential,
    el.thinkingMode, el.watermark, el.bboxEnabled, el.customWidth, el.customHeight, el.slideAspect, el.slideTemplate,
    el.includeRegionsInPrompt, el.assistantMode, el.assistantUseImages, el.assistantUseRegions, el.assistantUsePrompt,
    el.assistantThinking, el.assistantRequest, el.workflowPageCount, el.workflowTheme, el.workflowContent, el.manualPageGoal,
  ].forEach((node) => {
    node.addEventListener("input", () => {
      saveSettings();
      updatePayloadPreview();
      if (node === el.manualPageGoal) renderManualLayoutPreview();
    });
    node.addEventListener("change", () => {
      saveSettings();
      updatePayloadPreview();
      if (node === el.bboxEnabled) renderImages();
      if (node === el.slideAspect) {
        renderSlidePlanner();
        refreshWorkflowPagePrompts();
        renderManualLayoutPreview();
      }
      if (node === el.includeRegionsInPrompt) {
        refreshWorkflowPagePrompts();
        renderManualLayoutPreview();
      }
      if (node === el.workflowTheme) {
        renderWorkflowThemePreview();
        refreshWorkflowPagePrompts();
        renderManualLayoutPreview();
      }
      if (node === el.manualPageGoal) renderManualLayoutPreview();
    });
  });

  el.autoReplaceTarget.addEventListener("change", saveSettings);
  el.bboxEnabled.addEventListener("change", () => {
    syncBboxToggles(el.bboxEnabled.checked);
    saveSettings();
    renderEditTargetControls();
    renderImages();
  });
  el.bboxEnabledImagePanel.addEventListener("change", () => {
    syncBboxToggles(el.bboxEnabledImagePanel.checked);
    saveSettings();
    renderEditTargetControls();
    renderImages();
    updatePayloadPreview();
  });
  el.editTargetSelect.addEventListener("change", () => {
    state.editTargetImageId = el.editTargetSelect.value || null;
    saveSettings();
    renderEditTargetControls();
    renderImages();
    updatePayloadPreview();
  });
  el.undoTargetBtn.addEventListener("click", undoEditTargetVersion);
  el.sizeMode.addEventListener("change", () => { saveSettings(); toggleSizeMode(); updatePayloadPreview(); });
  el.fileInput.addEventListener("change", handleFiles);
  el.addImageUrlBtn.addEventListener("click", handleAddImageUrl);
  el.addPaletteBtn.addEventListener("click", addPaletteRow);
  el.sendBtn.addEventListener("click", sendRequest);
  el.refreshTaskBtn.addEventListener("click", refreshCurrentTask);
  el.copyPayloadBtn.addEventListener("click", copyPayload);
  el.downloadAllBtn.addEventListener("click", downloadAllResults);
  el.applyTemplateBtn.addEventListener("click", () => {
    state.slideRegions = buildTemplateRegions(el.slideTemplate.value);
    renderSlidePlanner();
    refreshWorkflowPagePrompts();
    renderManualLayoutPreview();
    updatePayloadPreview();
    setStatus("已应用 PPT 版式模板。", "success");
  });
  el.syncSlideSizeBtn.addEventListener("click", () => {
    const aspect = getCurrentAspectMeta();
    el.sizeMode.value = "custom";
    el.customWidth.value = String(aspect.outputWidth);
    el.customHeight.value = String(aspect.outputHeight);
    toggleSizeMode();
    saveSettings();
    updatePayloadPreview();
    setStatus(`已同步为 ${aspect.outputWidth}x${aspect.outputHeight}。`, "success");
  });
  el.insertRegionsBtn.addEventListener("click", () => {
    const text = buildSlideRegionText();
    if (!text) return setStatus("当前还没有 PPT 区域可插入。", "error");
    const current = el.prompt.value.trim();
    el.prompt.value = current ? `${current}\n\n${text}` : text;
    saveSettings();
    refreshWorkflowPagePrompts();
    renderManualLayoutPreview();
    updatePayloadPreview();
    setStatus("已将区域说明追加到主提示词。", "success");
  });
  el.clearRegionsBtn.addEventListener("click", () => {
    state.slideRegions = [];
    renderSlidePlanner();
    refreshWorkflowPagePrompts();
    renderManualLayoutPreview();
    updatePayloadPreview();
    setStatus("已清空 PPT 区域。", "success");
  });
  el.manualBuildPromptBtn.addEventListener("click", () => {
    const prompt = buildManualPromptDraft();
    if (!prompt) return setStatus("请先填写页面目标或区域。", "error");
    el.prompt.value = prompt;
    saveSettings();
    updatePayloadPreview();
    setStatus("已按当前主题和版式生成主提示词骨架。", "success");
  });
  el.manualAppendPromptBtn.addEventListener("click", () => {
    const prompt = buildManualPromptDraft();
    if (!prompt) return setStatus("请先填写页面目标或区域。", "error");
    el.prompt.value = `${el.prompt.value.trim()}${el.prompt.value.trim() ? "\n\n" : ""}${prompt}`;
    saveSettings();
    updatePayloadPreview();
    setStatus("已把当前版式骨架追加到主提示词。", "success");
  });
  el.manualCopyBriefBtn.addEventListener("click", async () => {
    try {
      await navigator.clipboard.writeText(el.manualLayoutPreview.textContent || "");
      setStatus("排版摘要已复制。", "success");
    } catch (error) {
      setStatus(error.message || "复制排版摘要失败。", "error");
    }
  });
  el.assistantSendBtn.addEventListener("click", sendAssistantRequest);
  el.assistantApplyBtn.addEventListener("click", () => applyAssistantPrompt(true));
  el.assistantAppendBtn.addEventListener("click", () => applyAssistantPrompt(false));
  el.workflowPlanBtn.addEventListener("click", planWorkflowPages);
  el.workflowBatchBtn.addEventListener("click", generateAllWorkflowPages);
  el.workflowCopyBtn.addEventListener("click", copyWorkflowPlan);
  el.workflowClearBtn.addEventListener("click", clearWorkflowPlan);
}

function init() {
  loadSettings();
  syncBboxToggles(el.bboxEnabled.checked);
  toggleSizeMode();
  bindTabNavigation();
  bindEvents();
  setupSlideCanvas();
  renderPalette();
  renderSlidePlanner();
  renderWorkflowThemePreview();
  renderWorkflowPlan();
  renderImages();
  renderAssistantResult();
  renderManualLayoutPreview();
  applyTabState();
  updatePayloadPreview();
  setStatus("页面已就绪，可以开始配置、拆页和改图。");
}

init();

function getThemeName() {
  return (el.themeName?.value || "").trim();
}

function getFallbackThemeDefinition(themeName = getThemeName()) {
  const resolvedThemeName = themeName || "未命名风格";
  return {
    label: resolvedThemeName,
    summary: `${resolvedThemeName} 的简版风格骨架`,
    basic: `请围绕“${resolvedThemeName}”生成适用于 16:9 演示文稿的高质量视觉画面。整体需要高级、统一、留白稳定，兼顾材质、光影、主视觉锚点与文字可读性，画面必须适合标题、正文和图表排版。`,
    cover: `当前页面是封面页。请使用海报式极简构图，只保留一个强视觉中心和大面积干净留白，为标题预留稳定区域，避免碎片化装饰、复杂图表和干扰性小元素。整体风格围绕“${resolvedThemeName}”。`,
    content: `当前页面是内容页。请围绕“${resolvedThemeName}”使用稳定网格、清晰信息层级和适合正文排版的容器结构，重点照顾标题、要点、解释性文字和视觉主图之间的平衡。`,
    data: `当前页面是数据页。请围绕“${resolvedThemeName}”让图表或数据可视化具备立体材质与视觉锚点，同时保留足够空白给数字、标签和结论，避免平面化、拥挤化表达。`,
  };
}

function normalizeThemeDefinition(parsed, themeName = getThemeName()) {
  const resolvedThemeName = themeName || "未命名风格";
  const fallback = getFallbackThemeDefinition(resolvedThemeName);
  return {
    label: resolvedThemeName,
    summary: `${resolvedThemeName} 的风格模板`,
    basic: String(parsed?.basic || fallback.basic).trim(),
    cover: String(parsed?.cover || fallback.cover).trim(),
    content: String(parsed?.content || fallback.content).trim(),
    data: String(parsed?.data || fallback.data).trim(),
  };
}

function getCurrentThemePreset() {
  return hasCurrentThemeDefinition() ? state.themeDefinition : getFallbackThemeDefinition();
}

function normalizeThemeLibraryEntry(entry) {
  const label = String(entry?.label || entry?.name || "").trim();
  if (!label) return null;
  const definition = normalizeThemeDefinition(entry?.definition || entry, label);
  return {
    label,
    definition,
    raw: String(entry?.raw || formatJSON(definition)),
    confirmed: Boolean(entry?.confirmed),
    updatedAt: Number(entry?.updatedAt || Date.now()),
  };
}

function getThemeLibraryEntry(label = getThemeName()) {
  return state.themeLibrary.find((entry) => entry.label === label) || null;
}

function upsertThemeLibraryEntry({ label, definition, raw, confirmed = false }) {
  const entry = normalizeThemeLibraryEntry({
    label,
    definition,
    raw,
    confirmed,
    updatedAt: Date.now(),
  });
  if (!entry) return;

  const existingIndex = state.themeLibrary.findIndex((item) => item.label === entry.label);
  if (existingIndex >= 0) {
    state.themeLibrary[existingIndex] = {
      ...state.themeLibrary[existingIndex],
      ...entry,
      confirmed: confirmed || state.themeLibrary[existingIndex].confirmed,
    };
  } else {
    state.themeLibrary.unshift(entry);
  }

  state.themeLibrary = state.themeLibrary
    .filter((item, index, list) => list.findIndex((candidate) => candidate.label === item.label) === index)
    .sort((a, b) => b.updatedAt - a.updatedAt)
    .slice(0, 24);
}

function applyThemeLibraryEntry(label) {
  const entry = getThemeLibraryEntry(label);
  if (!entry) throw new Error("未找到这个已生成风格。");

  el.themeName.value = entry.label;
  state.themeDefinition = entry.definition;
  state.themeDefinitionRaw = entry.raw || formatJSON(entry.definition);
  state.themeDefinitionSource = entry.label;
  state.themeConfirmed = Boolean(entry.confirmed);
  state.themeConfirmedSource = entry.confirmed ? entry.label : "";
  el.workflowTheme.value = entry.label;
  if (state.workflowPages.length) {
    resetAllWorkflowPageDesigns("已切换共享风格，请重新整理逐页版式。");
  }
  saveSettings();
  renderThemeLibrary();
  renderThemeStatus();
  syncWorkflowGate();
  renderManualLayoutPreview();
  updatePayloadPreview();
}

function renderThemeLibrary() {
  if (!el.themeLibrarySelect) return;

  const current = getThemeName();
  if (!state.themeLibrary.length) {
    el.themeLibrarySelect.innerHTML = '<option value="">还没有已生成风格</option>';
    el.themeLibrarySelect.disabled = true;
    return;
  }

  el.themeLibrarySelect.disabled = false;
  el.themeLibrarySelect.innerHTML = [
    '<option value="">请选择已生成风格</option>',
    ...state.themeLibrary.map((entry) => {
      const suffix = entry.confirmed ? "（已确认）" : "";
      return `<option value="${escapeHtml(entry.label)}">${escapeHtml(entry.label)}${suffix}</option>`;
    }),
  ].join("");

  if (current && state.themeLibrary.some((entry) => entry.label === current)) {
    el.themeLibrarySelect.value = current;
  }
}

function renderThemeStatus() {
  const themeName = getThemeName();

  if (!themeName) {
    el.themeStatus.textContent = "先输入几个风格词，再生成并确认风格模板。确认后才会进入正文拆分页。";
  } else if (hasConfirmedThemeDefinition(themeName)) {
    el.themeStatus.textContent = `已确认“${themeName}”风格模板。省事流会按 ${getPptOutputDescription()} 输出。`;
  } else if (hasCurrentThemeDefinition(themeName)) {
    el.themeStatus.textContent = `“${themeName}”风格模板已生成，请先确认；确认后才会开放正文拆分页。`;
  } else if (state.themeDefinitionSource && state.themeDefinitionSource !== themeName) {
    el.themeStatus.textContent = `你把风格改成了“${themeName}”。请重新生成并确认新的风格模板。`;
  } else {
    el.themeStatus.textContent = `当前还没有“${themeName}”的风格模板。先点“AI 生成风格模板”。`;
  }

  if (el.themeDefinitionPreview) {
    el.themeDefinitionPreview.textContent = hasCurrentThemeDefinition(themeName)
      ? (state.themeDefinitionRaw || formatJSON(state.themeDefinition))
      : "";
  }

  if (el.workflowTheme) {
    el.workflowTheme.value = themeName;
  }

  renderThemeLibrary();
  renderThemeReviewPanel();
}

function buildWorkflowPlanLibraryLabel(entry) {
  const themeLabel = entry?.themeLabel || getThemeName() || "未命名风格";
  const firstTitle = entry?.pages?.[0]?.pageTitle || "未命名页面";
  const baseLabel = entry?.label || `${themeLabel} / ${firstTitle}`;
  const timeLabel = entry?.updatedAt ? ` · ${formatHistoryTime(entry.updatedAt).slice(5)}` : "";
  return `${baseLabel}${timeLabel}`;
}

function upsertWorkflowPlanLibraryEntry({
  label,
  themeLabel = getThemeName(),
  workflowContent = el.workflowContent.value,
  workflowPageCount = el.workflowPageCount.value,
  workflowPlanSummary = state.workflowPlanSummary,
  pages = state.workflowPages,
} = {}) {
  const normalizedContent = String(workflowContent || "").trim();
  const serializedPages = hydrateWorkflowPages((pages || []).map((page) => serializeWorkflowPage(page)));
  if (!normalizedContent || !serializedPages.length) return null;

  const trimmedThemeLabel = String(themeLabel || "").trim();
  const nextLabel = String(label || buildWorkflowPlanLibraryLabel({ themeLabel: trimmedThemeLabel, pages: serializedPages })).trim();
  const existing = state.workflowPlanLibrary.find((entry) => (
    entry.themeLabel === trimmedThemeLabel
    && entry.workflowContent === normalizedContent
    && Number(entry.workflowPageCount || serializedPages.length) === Number(workflowPageCount || serializedPages.length)
  ));

  const nextEntry = {
    id: existing?.id || uid(),
    label: nextLabel,
    themeLabel: trimmedThemeLabel,
    workflowContent: normalizedContent,
    workflowPageCount: String(workflowPageCount || serializedPages.length),
    workflowPlanSummary: String(workflowPlanSummary || "").trim(),
    pages: serializedPages,
    createdAt: existing?.createdAt || Date.now(),
    updatedAt: Date.now(),
  };

  const existingIndex = state.workflowPlanLibrary.findIndex((entry) => entry.id === nextEntry.id);
  if (existingIndex >= 0) {
    state.workflowPlanLibrary[existingIndex] = nextEntry;
  } else {
    state.workflowPlanLibrary.unshift(nextEntry);
  }

  state.workflowPlanLibrary = state.workflowPlanLibrary
    .slice()
    .sort((a, b) => b.updatedAt - a.updatedAt)
    .slice(0, WORKFLOW_PLAN_LIBRARY_LIMIT);
  state.workflowPlanLibraryActiveId = nextEntry.id;
  renderWorkflowPlanLibrary();
  return nextEntry;
}

function renderWorkflowPlanLibrary() {
  if (!el.workflowPlanLibrarySelect) return;

  if (!state.workflowPlanLibrary.length) {
    el.workflowPlanLibrarySelect.innerHTML = '<option value="">还没有已保存拆分页</option>';
    el.workflowPlanLibrarySelect.disabled = true;
    renderWorkflowPlanHistory();
    return;
  }

  el.workflowPlanLibrarySelect.disabled = false;
  el.workflowPlanLibrarySelect.innerHTML = [
    '<option value="">选择一套已保存拆分页</option>',
    ...state.workflowPlanLibrary.map((entry) => (
      `<option value="${escapeHtml(entry.id)}">${escapeHtml(buildWorkflowPlanLibraryLabel(entry))}</option>`
    )),
  ].join("");

  if (state.workflowPlanLibraryActiveId && state.workflowPlanLibrary.some((entry) => entry.id === state.workflowPlanLibraryActiveId)) {
    el.workflowPlanLibrarySelect.value = state.workflowPlanLibraryActiveId;
  }

  renderWorkflowPlanHistory();
}

function applyWorkflowPlanLibraryEntry(entryId) {
  const entry = state.workflowPlanLibrary.find((item) => item.id === entryId);
  if (!entry) throw new Error("没找到这套已保存拆分页。");

  if (entry.themeLabel && getThemeLibraryEntry(entry.themeLabel)) {
    state.workflowPages = [];
    applyThemeLibraryEntry(entry.themeLabel);
  }

  el.workflowContent.value = entry.workflowContent;
  el.workflowPageCount.value = entry.workflowPageCount || String(entry.pages.length || 4);
  state.workflowPlanSummary = entry.workflowPlanSummary || "";
  state.workflowPages = hydrateWorkflowPages(entry.pages.map((page) => serializeWorkflowPage(page)));
  state.workflowPlanLibraryActiveId = entry.id;
  state.workflowDetailPageId = null;
  renderWorkflowPlanLibrary();
  renderWorkflowPlan();
  saveSettings();
}

function renderWorkflowPlanHistory() {
  if (!el.workflowPlanHistoryPanel) return;
  const activeEntry = state.workflowPlanLibrary.find((entry) => entry.id === state.workflowPlanLibraryActiveId);
  if (!activeEntry) {
    el.workflowPlanHistoryPanel.classList.add("hidden");
    if (el.workflowPlanHistoryMeta) el.workflowPlanHistoryMeta.textContent = "选中后会显示这套方案的主题、时间和封面页摘要。";
    if (el.workflowPlanHistoryLead) el.workflowPlanHistoryLead.textContent = "";
    if (el.workflowPlanHistorySummary) el.workflowPlanHistorySummary.textContent = "";
    if (el.workflowPlanDeleteBtn) el.workflowPlanDeleteBtn.disabled = true;
    return;
  }

  const leadPage = activeEntry.pages[0];
  const pageCount = activeEntry.pages.length;
  const metaParts = [
    activeEntry.themeLabel ? `主题：${activeEntry.themeLabel}` : "",
    `最近保存：${formatHistoryTime(activeEntry.updatedAt)}`,
    activeEntry.createdAt ? `首次生成：${formatHistoryTime(activeEntry.createdAt)}` : "",
  ].filter(Boolean);

  el.workflowPlanHistoryPanel.classList.remove("hidden");
  if (el.workflowPlanHistoryMeta) el.workflowPlanHistoryMeta.textContent = metaParts.join(" · ");
  if (el.workflowPlanHistoryLead) {
    el.workflowPlanHistoryLead.textContent = [
      `共 ${pageCount} 页`,
      leadPage?.pageTitle ? `封面页：${leadPage.pageTitle}` : "",
      activeEntry.label ? `档案名：${activeEntry.label}` : "",
    ].filter(Boolean).join("\n");
  }
  if (el.workflowPlanHistorySummary) {
    el.workflowPlanHistorySummary.textContent = activeEntry.workflowPlanSummary
      || truncateText(activeEntry.workflowContent, 180)
      || "这套拆分页还没有额外摘要。";
  }
  if (el.workflowPlanDeleteBtn) el.workflowPlanDeleteBtn.disabled = false;
}

function deleteWorkflowPlanLibraryEntry(entryId = state.workflowPlanLibraryActiveId) {
  const entry = state.workflowPlanLibrary.find((item) => item.id === entryId);
  if (!entry) throw new Error("没找到要删除的拆分页档案。");

  state.workflowPlanLibrary = state.workflowPlanLibrary.filter((item) => item.id !== entryId);
  if (state.workflowPlanLibraryActiveId === entryId) {
    state.workflowPlanLibraryActiveId = "";
  }
  if (el.workflowPlanLibrarySelect) {
    el.workflowPlanLibrarySelect.value = "";
  }
  renderWorkflowPlanLibrary();
  saveSettings();
}

function openImagePreview({ title, imageUrl }) {
  if (!imageUrl) return;
  state.imagePreview = {
    title: title || "图片预览",
    url: imageUrl,
  };
  if (el.imagePreviewTitle) el.imagePreviewTitle.textContent = state.imagePreview.title;
  if (el.imagePreviewImage) el.imagePreviewImage.src = state.imagePreview.url;
  if (el.imagePreviewOpenNew) el.imagePreviewOpenNew.href = state.imagePreview.url;
  el.imagePreviewModal?.classList.remove("hidden");
  el.imagePreviewModal?.setAttribute("aria-hidden", "false");
}

function closeImagePreview() {
  state.imagePreview = { title: "", url: "" };
  if (el.imagePreviewImage) el.imagePreviewImage.removeAttribute("src");
  if (el.imagePreviewPrompt) el.imagePreviewPrompt.value = "";
  if (el.imagePreviewOpenNew) el.imagePreviewOpenNew.href = "#";
  el.imagePreviewModal?.classList.add("hidden");
  el.imagePreviewModal?.setAttribute("aria-hidden", "true");
}

async function sendImagePreviewToRevise({ submitImmediately = false } = {}) {
  const imageUrl = state.imagePreview?.url;
  if (!imageUrl) {
    setStatus("当前没有可送去改图的图片。", "error");
    return;
  }

  await addImage({
    name: `${state.imagePreview.title || "图片"}（改图底图）`,
    source: imageUrl,
  });
  state.editTargetImageId = state.images[state.images.length - 1]?.id || null;
  state.activeTabs.main = "revise";
  syncBboxToggles(true);
  if (el.imagePreviewPrompt?.value.trim()) {
    el.prompt.value = el.imagePreviewPrompt.value.trim();
  }
  saveSettings();
  renderImages();
  renderEditTargetControls();
  applyTabState();
  updatePayloadPreview();
  closeImagePreview();

  if (submitImmediately) {
    await sendRequest();
  } else {
    setStatus("已把图片送入改图工作台，可以继续框选或直接修改。", "success");
  }
}

function renderThemeReviewPanel() {
  const hasCurrent = hasCurrentThemeDefinition();
  const theme = hasCurrent ? state.themeDefinition : null;

  if (!hasCurrent || !theme) {
    el.themeReviewPanel?.classList.add("hidden");
    if (el.themeBasicPreview) el.themeBasicPreview.textContent = "";
    if (el.themeCoverPreview) el.themeCoverPreview.textContent = "";
    if (el.themeContentPreview) el.themeContentPreview.textContent = "";
    if (el.themeDataPreview) el.themeDataPreview.textContent = "";
    if (el.confirmThemeBtn) {
      el.confirmThemeBtn.disabled = true;
      el.confirmThemeBtn.textContent = "确认这个风格";
    }
    return;
  }

  el.themeReviewPanel.classList.remove("hidden");
  el.themeBasicPreview.textContent = theme.basic;
  el.themeCoverPreview.textContent = theme.cover;
  el.themeContentPreview.textContent = theme.content;
  el.themeDataPreview.textContent = theme.data;
  el.confirmThemeBtn.disabled = hasConfirmedThemeDefinition();
  el.confirmThemeBtn.textContent = hasConfirmedThemeDefinition() ? "已确认这个风格" : "确认这个风格";
}

function syncWorkflowGate() {
  const confirmed = hasConfirmedThemeDefinition();
  if (el.workflowGateHint) {
    el.workflowGateHint.dataset.state = confirmed ? "ready" : "locked";
    el.workflowGateHint.textContent = confirmed
      ? `已确认“${getThemeName()}”，现在可以粘贴正文并拆分页。当前省事流会固定输出 ${getPptOutputDescription()}。`
      : "先在上面输入风格关键词，生成并确认风格模板；确认之后再贴入原始内容并拆分页。";
  }

  [el.workflowPageCount, el.workflowContent, el.workflowPlanBtn].forEach((node) => {
    if (node) node.disabled = !confirmed;
  });
}

function renderWorkflowThemePreview() {
  if (el.workflowThemePreview) {
    el.workflowThemePreview.innerHTML = "";
    el.workflowThemePreview.classList.add("hidden");
  }
}

function getThemeAgentSystemPrompt() {
  return [
    "# Role",
    "你是一位世界顶级的视觉艺术总监、Prompt Engineer 及 UI/UX 专家，擅长把抽象风格词扩展成完整的演示文稿视觉语言系统。",
    "",
    "# Task",
    "请根据用户提供的 [目标风格主题]，按照指定的 JSON 结构，生成一份高质量的演示文稿视觉风格定义文件。",
    "",
    "# Constraints",
    "1. 必须把简单风格词扩展成完整视觉语言系统：风格融合、光影材质、配色方案、网格容器、视觉锚点、渲染质量都要写清楚。",
    "2. Basic 中必须融合 3 种相关高级风格，并明确材质、灯光、色谱、容器细节、3D 主视觉和渲染引擎。",
    "3. Cover 必须强制采用极简海报模式，并写清楚负向约束：禁止出现网格、卡片、图表等干扰元素，只保留一个核心 3D 主视觉和宽敞留白。",
    "4. Content 必须强调基于网格系统的信息布局、容器材质、边缘、阴影和留白。",
    "5. Data 必须把图表物体化，不能是平面图表，必须是符合主题的 3D 转译方式。",
    "6. basic、cover、content、data 必须全部使用简体中文输出。除非用户明确要求英文，否则不要用英文长句。",
    "",
    "## 3. 内容页逻辑 (Content)",
    "使用 Basic 中定义的网格系统，强调容器材质、边缘、阴影和留白。",
    "",
    "## 4. 数据页逻辑 (Data)",
    "将数据图表（饼图、柱状图）进行物体化转译。图表不能是平面的，必须是符合主题的 3D 物体化表达。",
    "",
    "# Output",
    "请只输出 JSON，不要解释，不要 Markdown，不要代码块。",
    "{\"basic\":\"...\",\"cover\":\"...\",\"content\":\"...\",\"data\":\"...\"}",
  ].join("\n");
}

function buildThemeDefinitionPayload(themeName) {
  return {
    model: "qwen3.6-plus",
    input: {
      messages: [
        {
          role: "system",
          content: getThemeAgentSystemPrompt(),
        },
        {
          role: "user",
          content: [
            {
              text: `现在，请生成${themeName}的提示词模板，并生成对应的JSON，不要包含任何解释和说明性内容。`,
            },
          ],
        },
      ],
    },
    parameters: {
      result_format: "message",
      response_format: { type: "json_object" },
      enable_thinking: el.assistantThinking.checked,
    },
  };
}

async function requestThemeDefinition(force = true) {
  const themeName = getThemeName();
  if (!themeName) throw new Error("请先输入风格关键词。");
  if (!el.apiKey.value.trim()) throw new Error("请先填写 API Key。");
  if (hasCurrentThemeDefinition(themeName) && !force) return state.themeDefinition;

  setButtonLoading(el.generateThemeBtn, true, "生成中...");
  setProgressLoading(el.themeProgress, true);
  setStatus(`正在为“${themeName}”生成风格模板...`, "running");
  el.themeStatus.textContent = `正在生成“${themeName}”的风格模板，请稍候...`;

  try {
    const payload = buildThemeDefinitionPayload(themeName);
    const response = await fetch("/api/assistant", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ apiKey: el.apiKey.value.trim(), region: el.region.value, payload }),
    });
    const data = await response.json();
    syncApiKeyFeedback(response, data);
    if (!response.ok || data.code) {
      throw new Error(data.message || "风格模板生成失败。");
    }

    const rawText = extractAssistantMessageText(data);
    let parsed;
    try {
      parsed = JSON.parse(rawText);
    } catch (error) {
      throw new Error(error.message || "风格模板 JSON 解析失败。");
    }

    state.themeDefinition = normalizeThemeDefinition(parsed, themeName);
    state.themeDefinitionRaw = formatJSON(parsed);
    state.themeDefinitionSource = themeName;
    state.themeConfirmed = false;
    state.themeConfirmedSource = "";
    upsertThemeLibraryEntry({
      label: themeName,
      definition: state.themeDefinition,
      raw: state.themeDefinitionRaw,
      confirmed: false,
    });
    saveSettings();
    renderThemeStatus();
    syncWorkflowGate();
    updatePayloadPreview();
    setStatus(`“${themeName}”风格模板已生成，请确认后继续。`, "success");
    return state.themeDefinition;
  } finally {
    setButtonLoading(el.generateThemeBtn, false);
    setProgressLoading(el.themeProgress, false);
  }
}

function requireConfirmedTheme() {
  const themeName = getThemeName();
  if (!themeName) throw new Error("请先输入风格关键词。");
  if (!hasCurrentThemeDefinition(themeName)) throw new Error("请先生成当前风格模板。");
  if (!hasConfirmedThemeDefinition(themeName)) throw new Error("请先确认当前风格模板，再继续下一步。");
  return state.themeDefinition;
}

function inferManualPageType() {
  const goal = (el.manualPageGoal?.value || "").trim();
  if (/数据|占比|同比|环比|kpi|图表|统计|趋势/i.test(goal)) return "data";
  if (state.slideRegions.length <= 2) return "cover";
  return "content";
}

function buildWorkflowPagePrompt(page) {
  const theme = requireConfirmedTheme();
  const aspect = getCurrentAspectMeta();
  const readabilityRules = [
    "文字必须明显偏大，标题优先使用大号中文标题，正文宁可减少碎句也不要做成密集小字。",
    "请优先保证标题、关键数字、结论和正文的远距离可读性，避免把文字压到边缘或复杂纹理上。",
    "除非原文极少，否则不要把正文做成大段连续小字，优先拆成大标题 + 2 到 4 个清晰内容块。",
  ].join("\n");

  return [
    theme.basic,
    theme[page.pageType] || theme.content,
    buildDecorationPromptBlock(getPageDecorationLevel(page)),
    buildNoExtraTextConstraintBlock(),
    `输出要求：${aspect.label} 画幅，目标尺寸 ${aspect.outputWidth}x${aspect.outputHeight}。这是约 4K 的 PPT 画面，必须优先保证页面清晰度、标题醒目度和中文正文可读性。`,
    "除非特殊说明，否则画面中的文字请使用与原文一致的语言。",
    `当前页面类型：${PAGE_TYPE_LABELS[page.pageType] || page.pageType}`,
    page.pageTitle ? `当前页面标题：${page.pageTitle}` : "",
    page.visualFocus ? `视觉焦点：${page.visualFocus}` : "",
    page.layoutSummary ? `版式安排：${page.layoutSummary}` : "",
    page.textHierarchy ? `文字层级：${page.textHierarchy}` : "",
    page.readabilityNotes ? `可读性要求：${page.readabilityNotes}` : readabilityRules,
    `当前页面内容：\n${page.pageContent || page.pageTitle || ""}`,
  ].filter(Boolean).join("\n\n");
}

function buildManualPromptDraft() {
  const theme = requireConfirmedTheme();
  const pageGoal = el.manualPageGoal.value.trim();
  const regionText = buildSlideRegionText();
  const pageType = inferManualPageType();

  return [
    theme.basic,
    theme[pageType] || theme.content,
    buildDecorationPromptBlock(getGlobalDecorationLevel()),
    "如果没有明确提供要上屏的具体文字，不要自行杜撰标题、正文、英文副标题、数字或标签。",
    buildNoExtraTextConstraintBlock({ scope: "decoration" }),
    `输出要求：${getPptOutputDescription()}，请严格按 PPT 画幅构图。`,
    pageGoal ? `当前页面目标：${pageGoal}` : "",
    regionText,
    "请把版式秩序、文本留白和视觉层级一起考虑进去，优先保证标题和正文可读性。",
  ].filter(Boolean).join("\n\n");
}

function renderManualLayoutPreview() {
  if (!el.manualLayoutPreview) return;
  if (!hasConfirmedThemeDefinition()) {
    el.manualLayoutPreview.textContent = "// 先生成并确认风格模板，再整理这一页的提示词概览。";
    return;
  }
  const hasGoal = Boolean(el.manualPageGoal.value.trim());
  const hasRegions = state.slideRegions.length > 0;
  if (!hasGoal && !hasRegions) {
    el.manualLayoutPreview.textContent = "// 先在左侧画出区域，或先写下这一页的页面目标，这里会自动汇总成该页提示词概览。";
    return;
  }
  el.manualLayoutPreview.textContent = buildManualPromptDraft();
}

function renderSlideRegionList() {
  if (!el.slideRegionList) return;
  el.slideRegionList.innerHTML = "";
}

function buildPromptText() {
  const prompt = el.prompt.value.trim();
  if (!prompt) throw new Error("请先填写提示词。");

  const parts = [];
  const shouldInjectTheme = hasConfirmedThemeDefinition() && state.activeTabs.main === "manual";

  if (shouldInjectTheme) {
    const theme = getCurrentThemePreset();
    const themePrefix = state.activeTabs.main === "manual"
      ? [theme.basic, theme[inferManualPageType()] || theme.content]
      : [theme.basic, theme.content];
    const promptAlreadyHasTheme = themePrefix.some((entry) => entry && prompt.includes(entry.slice(0, 20)));
    if (!promptAlreadyHasTheme) {
      parts.push(...themePrefix.filter(Boolean));
    }
  }
  parts.push(prompt);

  if (el.includeRegionsInPrompt.checked && state.slideRegions.length) {
    parts.push(buildSlideRegionText());
  }

  return parts.filter(Boolean).join("\n\n");
}

function buildPayload() {
  const orderedImages = getOrderedImagesForPayload();
  const parameters = buildGenerationParameters({
    orderedImages,
    enableSequential: el.enableSequential.checked,
    enableBbox: el.bboxEnabled.checked,
  });

  if (state.activeTabs.main === "manual") {
    parameters.size = getPptOutputSizeValue();
  }

  const promptText = buildPromptText();
  const editText = buildEditTargetText();
  return {
    model: el.model.value,
    input: {
      messages: [{
        role: "user",
        content: [
          ...orderedImages.map((image) => ({ image: image.source })),
          { text: editText ? `${promptText}\n\n${editText}` : promptText },
        ],
      }],
    },
    parameters,
  };
}

function buildWorkflowGenerationPayload(page) {
  const parameters = buildGenerationParameters({
    n: 1,
    orderedImages: [],
    enableSequential: false,
    enableBbox: false,
  });
  parameters.size = getPptOutputSizeValue();

  return {
    model: el.model.value,
    input: {
      messages: [
        {
          role: "user",
          content: [{ text: page.pagePrompt }],
        },
      ],
    },
    parameters,
  };
}

function buildWorkflowAssistantPayload() {
  const content = el.workflowContent.value.trim();
  const pageCount = Number(el.workflowPageCount.value || 0);
  const theme = requireConfirmedTheme();

  if (!content) throw new Error("请先粘贴需要拆分的长文本或讲稿。");
  if (!Number.isInteger(pageCount) || pageCount < 2 || pageCount > 20) {
    throw new Error("目标页数请输入 2 到 20 之间的整数。");
  }

  return {
    model: "qwen3.6-plus",
    input: {
      messages: [
        {
          role: "system",
          content: "你是一名专业的 PPT 内容策划师和视觉脚本编辑。你需要把用户原文拆成指定页数的 PPT 页面规划，并严格返回 JSON。除了封面标题外，其余页的 pageContent 尽量直接使用用户原文，不要改写、总结或漏掉关键信息。",
        },
        {
          role: "user",
          content: [
            {
              text: [
                `当前视觉主题：${theme.label}`,
                `当前主题风格定义：\n${formatJSON({ basic: theme.basic, cover: theme.cover, content: theme.content, data: theme.data })}`,
                `输出尺寸固定为：${getPptOutputDescription()}`,
                `请把下面内容规划成恰好 ${pageCount} 页 PPT。`,
                "硬性规则：",
                `1. 总页数必须严格等于 ${pageCount}。`,
                "2. 第 1 页必须是 cover。",
                "3. cover 页只放标题或主题概括，允许从原文提炼一个简洁标题。",
                "4. 第 2 页及之后的 pageContent 尽量直接使用用户原文，不要总结，不要洗稿，不要重写。",
                "5. pageType 只能是 cover、content、data；包含明确数字、占比、同比、环比、KPI 或统计结果的页优先使用 data。",
                "6. 拆分优先按章节、主题边界、自然段和逻辑单元进行，不要在句子中间硬切。",
                "7. 输出格式必须是 JSON object，字段包含 summary 和 pagePlan。",
                "8. pagePlan 中每一页都必须包含 pageNumber、pageType、pageTitle、pageContent。",
                "9. 不要输出图片提示词，后续系统会自动把主题 JSON 套进每一页提示词。",
                "10. 不要为每一页重复附加统一区域坐标；省事流由系统在下一步单独处理排版与风格提示词。",
                "输出示例：",
                '{"summary":"...","pagePlan":[{"pageNumber":1,"pageType":"cover","pageTitle":"...","pageContent":"..."},{"pageNumber":2,"pageType":"content","pageTitle":"...","pageContent":"..."}]}',
                "用户原文：",
                content,
              ].join("\n\n"),
            },
          ],
        },
      ],
    },
    parameters: {
      result_format: "message",
      response_format: { type: "json_object" },
      enable_thinking: el.assistantThinking.checked,
    },
  };
}

function buildAssistantPayload() {
  const mode = el.assistantMode.value;
  const request = el.assistantRequest.value.trim();
  const currentPrompt = el.assistantUsePrompt.checked ? el.prompt.value.trim() : "";
  const regionText = el.assistantUseRegions.checked ? buildSlideRegionText() : "";
  const editTargetText = buildEditTargetText();
  const includeImages = el.assistantUseImages.checked;
  const includeTheme = hasConfirmedThemeDefinition();
  const theme = includeTheme ? getCurrentThemePreset() : null;

  if (!request && !currentPrompt && !regionText && !(includeImages && state.images.length)) {
    throw new Error("请填写你的目标，或提供主提示词、区域、图片中的至少一种上下文。");
  }

  const content = [];
  if (includeImages) state.images.forEach((image) => content.push({ image: image.source }));
  content.push({
    text: [
      `当前任务模式：${ASSISTANT_MODE_LABELS[mode]}`,
      includeTheme
        ? `当前主题风格定义：\n${formatJSON({ basic: theme.basic, cover: theme.cover, content: theme.content, data: theme.data })}`
        : "当前未固定全局风格模板，请主要参考用户需求、参考图和当前提示词。",
      request ? `用户目标：${request}` : "用户未额外补充目标，请尽量从现有提示词和版式里推断。",
      currentPrompt ? `当前主提示词：${currentPrompt}` : "当前主提示词：无",
      regionText || "当前未提供 PPT 版式区域。",
      editTargetText || "当前未启用定点改图。",
      includeImages && state.images.length ? `已附带 ${state.images.length} 张参考图，请综合主体、构图、风格和色彩。` : "当前未附带参考图。",
      "请只输出 JSON，不要输出 markdown。JSON 字段必须包含：summary、recommendedPrompt、editPrompt、layoutAdvice、pptNotes、pagePlan。",
      "如果当前任务不是多页草案，pagePlan 返回空数组。recommendedPrompt 和 editPrompt 必须能直接用于图片生成或改图。",
    ].join("\n\n"),
  });

  return {
    model: "qwen3.6-plus",
    input: {
      messages: [
        {
          role: "system",
          content: "你是一名资深 PPT 视觉设计师、提示词工程师和版式顾问。重点考虑标题留白、文字可读性、主体位置、风格统一性，以及用户给出的主题风格 JSON。必须只输出 JSON。",
        },
        { role: "user", content },
      ],
    },
    parameters: {
      result_format: "message",
      response_format: { type: "json_object" },
      enable_thinking: el.assistantThinking.checked,
    },
  };
}

function buildWorkflowPageDesignPayload(page) {
  const theme = requireConfirmedTheme();
  const aspect = getCurrentAspectMeta();

  return {
    model: "qwen3.6-plus",
    input: {
      messages: [
        {
          role: "system",
          content: [
            "你是一位顶级 PPT 页面导演、信息设计师和提示词工程师。",
            "你的任务不是重写原文，而是基于当前页面的内容、页面类型和全局风格模板，为这一页生成更适合出图的排版策略与最终页面提示词。",
            "重点处理四件事：信息主次、视觉重心、留白与版式、字体大小和可读性。",
            "必须优先让标题、关键数字、核心结论和正文在 PPT 场景中清楚可读，避免满屏密集小字。",
            "请只输出 JSON object，不要输出 markdown。",
          ].join("\n"),
        },
        {
          role: "user",
          content: [
            {
              text: [
                `全局风格主题：${theme.label}`,
                `全局风格模板：\n${formatJSON({ basic: theme.basic, cover: theme.cover, content: theme.content, data: theme.data })}`,
                `当前页面类型：${PAGE_TYPE_LABELS[page.pageType] || page.pageType}`,
                `当前页面标题：${page.pageTitle || `第 ${page.pageNumber} 页`}`,
                `当前页面原文：\n${page.pageContent || page.pageTitle || ""}`,
      `输出尺寸固定为 ${aspect.label}，${aspect.outputWidth}x${aspect.outputHeight}，目标是清晰的 4K PPT 画面。`,
                "请根据这一页内容单独思考版式，不要把所有页都套成同一种布局。",
                "请特别注意：",
                "1. 标题区必须明显大，优先保证远距离可读。",
                "2. 正文宁可拆成 2 到 4 个大内容块，也不要做成密集小字墙。",
                "3. 数据页要放大关键数字或核心指标，不要把数据埋在角落里。",
                "4. 封面页要更像海报，正文页要更像信息展示，数据页要更像结论型图表页。",
                "5. 如果这一页原文偏长，请优先保留核心句、关键术语、数据和结论，用更适合 PPT 的大字信息块呈现，不要生成塞满整页的小字段落。",
                "6. 不要重复输出整份全局风格模板，只输出本页真正会用到的排版结果。",
                "返回 JSON，字段必须包含：layoutSummary、textHierarchy、visualFocus、readabilityNotes、pagePrompt。",
                "其中 pagePrompt 必须是可以直接发给生图模型的一整段中文提示词。",
              ].join("\n\n"),
            },
          ],
        },
      ],
    },
    parameters: {
      result_format: "message",
      response_format: { type: "json_object" },
      enable_thinking: el.assistantThinking.checked,
    },
  };
}

function applyWorkflowPageDesign(page, parsed) {
  page.layoutSummary = String(parsed?.layoutSummary || parsed?.layout_summary || "").trim();
  page.textHierarchy = String(parsed?.textHierarchy || parsed?.text_hierarchy || "").trim();
  page.visualFocus = String(parsed?.visualFocus || parsed?.visual_focus || "").trim();
  page.readabilityNotes = String(parsed?.readabilityNotes || parsed?.readability_notes || "").trim();
  page.pagePrompt = String(parsed?.pagePrompt || parsed?.page_prompt || "").trim() || buildWorkflowPagePrompt(page);
  page.layoutStatus = "ready";
  page.layoutError = "";
}

async function ensureWorkflowPageDesign(page, options = {}) {
  const { force = false, signal = null } = options;
  if (!page) throw new Error("未找到对应页面。");
  if (page.layoutStatus === "running" && page.layoutPromise) return page.layoutPromise;
  if (!force && page.pagePrompt && page.layoutStatus === "ready") return;
  if (!el.apiKey.value.trim()) throw new Error("请先填写 API Key。");

  if (signal?.aborted) throw signal.reason || createAbortError("已停止逐页版式整理。");
  page.layoutStatus = "running";
  page.layoutError = "正在整理这一页的版式和提示词...";
  renderWorkflowPlan();
  renderWorkflowDetail();

  page.layoutPromise = (async () => {
    try {
      const payload = buildWorkflowPageDesignPayload(page);
      const response = await fetch("/api/assistant", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        signal,
        body: JSON.stringify({
          apiKey: el.apiKey.value.trim(),
          region: el.region.value,
          payload,
        }),
      });
      const data = await response.json();
      syncApiKeyFeedback(response, data);
      if (!response.ok || data.code) {
        throw new Error(data.message || "逐页版式整理失败。");
      }

      const rawText = extractAssistantMessageText(data);
      let parsed;
      try {
        parsed = JSON.parse(rawText);
      } catch (error) {
        throw new Error(error.message || "逐页版式 JSON 解析失败。");
      }

      applyWorkflowPageDesign(page, parsed);
      upsertWorkflowPlanLibraryEntry();
      saveSettings();
      renderWorkflowPlan();
      renderWorkflowDetail();
    } catch (error) {
      page.layoutStatus = "error";
      page.layoutError = error.message || "逐页版式整理失败。";
      if (!page.pagePrompt) {
        page.pagePrompt = buildWorkflowPagePrompt(page);
      }
      renderWorkflowPlan();
      renderWorkflowDetail();
      throw error;
    } finally {
      page.layoutPromise = null;
    }
  })();

  return page.layoutPromise;
}

async function prepareWorkflowPageLayouts(pages = state.workflowPages) {
  let successCount = 0;
  let errorCount = 0;

  for (let index = 0; index < pages.length; index += 1) {
    const page = pages[index];
    setStatus(`正在整理逐页版式：${index + 1}/${pages.length} · 第 ${page.pageNumber} 页`, "running");
    try {
      await ensureWorkflowPageDesign(page, { force: page.layoutStatus === "error" });
      successCount += 1;
      saveSettings();
    } catch {
      errorCount += 1;
    }
  }

  return { successCount, errorCount };
}

function buildWorkflowPageLayoutText(page) {
  if (page.layoutStatus === "running") {
    return "正在根据这一页内容、页面类型和共享风格模板整理个性化版式...";
  }

  const lines = [
    `页面类型：${PAGE_TYPE_LABELS[page.pageType] || page.pageType}`,
    `输出尺寸：${getPptOutputDescription()}`,
    page.layoutSummary || "这一页还没有生成单独的版式安排。拆分页完成后，系统会自动补做这一页的版式整理。",
    page.textHierarchy ? `文字层级：${page.textHierarchy}` : "",
    page.visualFocus ? `视觉焦点：${page.visualFocus}` : "",
    page.readabilityNotes ? `可读性要求：${page.readabilityNotes}` : "默认要求：标题更大、正文不要密集小字、关键结论和数字优先放大。",
  ];

  return lines.filter(Boolean).join("\n\n");
}

function getWorkflowCardResultImageUrl(pageId) {
  const page = state.workflowPages.find((item) => item.id === pageId);
  const normalize = (input) => {
    const value = String(input || "").trim();
    if (!value || value === "undefined" || value === "null") return "";
    return value;
  };
  const pageNumber = normalize(page?.pageNumber);
  let card = null;

  if (pageId) {
    try {
      const escaped = typeof CSS !== "undefined" && typeof CSS.escape === "function"
        ? CSS.escape(String(pageId))
        : String(pageId).replace(/["\\]/g, "\\$&");
      card = el.workflowPlanCards?.querySelector(`.workflow-card[data-page-id="${escaped}"]`) || null;
    } catch {
      card = null;
    }
  }

  if (!card && pageNumber) {
    const cards = Array.from(el.workflowPlanCards?.querySelectorAll(".workflow-card") || []);
    card = cards.find((node) => {
      const explicit = normalize(node.getAttribute("data-page-number"));
      if (explicit && explicit === pageNumber) return true;
      const meta = normalize(node.querySelector(".workflow-page-index")?.textContent);
      return meta.includes(pageNumber);
    }) || null;
  }

  if (!card) return "";
  return normalize(
    card.getAttribute("data-result-image-url")
      || card.querySelector(".result-images img, .result-card img")?.getAttribute("src")
      || card.querySelector(".result-images img, .result-card img")?.src
  );
}

function hasWorkflowStructuredLayout(page) {
  if (!page || typeof page !== "object") return false;
  return [
    page.layoutSummary,
    page.textHierarchy,
    page.visualFocus,
    page.readabilityNotes,
  ].some((value) => String(value || "").trim());
}

function workflowPageNeedsDesign(page) {
  if (!page) return true;
  return !String(page.pagePrompt || "").trim()
    || page.layoutStatus !== "ready"
    || !hasWorkflowStructuredLayout(page);
}

function getWorkflowArchivedPageSnapshot(page) {
  if (!page) return null;
  const candidates = [];
  const addPages = (pages) => {
    if (!Array.isArray(pages)) return;
    pages.forEach((item) => {
      if (!item || typeof item !== "object") return;
      const sameNumber = Number(item.pageNumber || 0) === Number(page.pageNumber || 0);
      const sameType = String(item.pageType || "").trim() === String(page.pageType || "").trim();
      const sameTitle = String(item.pageTitle || "").trim() === String(page.pageTitle || "").trim();
      if (sameNumber && (sameTitle || sameType)) {
        candidates.push(item);
      }
    });
  };

  const activeEntry = state.workflowPlanLibrary.find((entry) => entry.id === state.workflowPlanLibraryActiveId);
  addPages(activeEntry?.pages);
  state.workflowPlanLibrary.forEach((entry) => {
    if (entry?.id === activeEntry?.id) return;
    addPages(entry?.pages);
  });

  return candidates.find((item) => {
    const hasImages = Array.isArray(item.resultImages) && item.resultImages.length;
    const hasBackdrop = String(item.detailBackdropUrl || "").trim();
    const hasSaved = item.savedResults && typeof item.savedResults === "object" && Object.keys(item.savedResults).length;
    return hasImages || hasBackdrop || hasSaved;
  }) || candidates[0] || null;
}

function getWorkflowSavedLocalImageUrl(page, imageUrl = "") {
  const entries = page?.savedResults && typeof page.savedResults === "object"
    ? Object.entries(page.savedResults)
    : [];
  const target = String(imageUrl || "").trim();
  if (target) {
    const matched = entries.find(([key, value]) => String(key || "").trim() === target && value?.localUrl);
    if (matched?.[1]?.localUrl) return String(matched[1].localUrl).trim();
  }
  return String(entries.find(([, value]) => value?.localUrl)?.[1]?.localUrl || "").trim();
}

function getWorkflowPreferredImageUrl(page, fallbackUrl = "") {
  return String(
    getWorkflowSavedLocalImageUrl(page, fallbackUrl)
    || getWorkflowSavedLocalImageUrl(page)
    || page?.detailBackdropUrl
    || fallbackUrl
    || ""
  ).trim();
}

async function ensureWorkflowPageLocalCache(page) {
  if (!page || !Array.isArray(page.resultImages) || !page.resultImages.length) return "";
  if (!page.savedResults || typeof page.savedResults !== "object") {
    page.savedResults = {};
  }

  let preferred = getWorkflowPreferredImageUrl(page, page.resultImages[0]);
  for (let index = 0; index < page.resultImages.length; index += 1) {
    const remoteUrl = String(page.resultImages[index] || "").trim();
    if (!remoteUrl) continue;
    const existingLocalUrl = getWorkflowSavedLocalImageUrl(page, remoteUrl);
    if (existingLocalUrl) {
      preferred = preferred || existingLocalUrl;
      continue;
    }

    const response = await fetch("/api/download", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        imageUrl: remoteUrl,
        requestId: page.requestId || `page-${page.pageNumber || index + 1}`,
        index: index + 1,
      }),
    });
    const data = await response.json();
    if (!response.ok || data.code) {
      throw new Error(data.message || "自动缓存本页结果图失败。");
    }
    page.savedResults[remoteUrl] = data;
    preferred = preferred || String(data.localUrl || "").trim();
  }

  const nextBackdropUrl = getWorkflowPreferredImageUrl(page, page.resultImages[0]);
  if (nextBackdropUrl) page.detailBackdropUrl = nextBackdropUrl;
  return nextBackdropUrl;
}

function syncWorkflowPageResultBackdrop(page) {
  if (!page) return "";
  const normalize = (input) => {
    const value = String(input || "").trim();
    if (!value || value === "undefined" || value === "null") return "";
    return value;
  };
  const archived = getWorkflowArchivedPageSnapshot(page);
  if (archived?.savedResults && typeof archived.savedResults === "object") {
    page.savedResults = {
      ...(archived.savedResults || {}),
      ...(page.savedResults && typeof page.savedResults === "object" ? page.savedResults : {}),
    };
  }
  if ((!Array.isArray(page.resultImages) || !page.resultImages.length) && Array.isArray(archived?.resultImages) && archived.resultImages.length) {
    page.resultImages = archived.resultImages.filter(Boolean);
  }
  if (!page.detailBackdropUrl && archived?.detailBackdropUrl) {
    page.detailBackdropUrl = String(archived.detailBackdropUrl).trim();
  }
  const resolved = [
    getWorkflowPreferredImageUrl(page, Array.isArray(page.resultImages) ? page.resultImages[0] : ""),
    normalize(page.detailBackdropUrl),
    ...(Array.isArray(page.resultImages) ? page.resultImages.map((item) => normalize(item)) : []),
    getWorkflowCardResultImageUrl(page.id),
  ].find(Boolean) || "";
  if (!resolved) return "";

  page.detailBackdropUrl = resolved;
  if (!Array.isArray(page.resultImages)) page.resultImages = [];
  if (!page.resultImages.includes(resolved)) {
    page.resultImages = [resolved, ...page.resultImages].filter(Boolean);
  }
  return resolved;
}

function openWorkflowDetail(pageId) {
  const page = state.workflowPages.find((item) => item.id === pageId);
  if (page && syncWorkflowPageResultBackdrop(page)) {
    saveSettings();
  }
  state.workflowDetailPageId = pageId;
  renderWorkflowDetail();
  el.workflowDetailModal?.classList.remove("hidden");
  el.workflowDetailModal?.setAttribute("aria-hidden", "false");

  if (page?.resultImages?.length) {
    ensureWorkflowPageLocalCache(page)
      .then((cachedUrl) => {
        if (!cachedUrl) return;
        saveSettings();
        renderWorkflowPlan();
        if (state.workflowDetailPageId === page.id) renderWorkflowDetail();
      })
      .catch(() => {});
  }

  if (page && workflowPageNeedsDesign(page) && page.layoutStatus !== "running") {
    ensureWorkflowPageDesign(page).catch((error) => {
      setStatus(error.message || "逐页版式整理失败。", "error");
    });
  }
}

function closeWorkflowDetail() {
  state.workflowDetailPageId = null;
  el.workflowDetailModal?.classList.add("hidden");
  el.workflowDetailModal?.setAttribute("aria-hidden", "true");
  renderWorkflowDetail();
}

function renderWorkflowDetail() {
  const page = state.workflowPages.find((item) => item.id === state.workflowDetailPageId);
  if (!page) {
    el.workflowDetailModal?.classList.add("hidden");
    el.workflowDetailModal?.setAttribute("aria-hidden", "true");
    return;
  }

  const theme = getCurrentThemePreset();
  el.workflowDetailTitle.textContent = `第 ${page.pageNumber} 页详情`;
  el.workflowDetailMeta.textContent = `${PAGE_TYPE_LABELS[page.pageType] || page.pageType} · ${page.pageTitle || `第 ${page.pageNumber} 页`} · 输出 ${getPptOutputDescription()}`;
  el.workflowDetailContent.textContent = `标题：${page.pageTitle || `第 ${page.pageNumber} 页`}\n\n原文内容：\n${page.pageContent || ""}`;
  el.workflowDetailLayout.textContent = buildWorkflowPageLayoutText(page);
  el.workflowDetailTheme.textContent = [
    `Basic：${theme.basic}`,
    "",
    `本页分类风格：${theme[page.pageType] || theme.content}`,
  ].join("\n");
  el.workflowDetailPrompt.textContent = page.pagePrompt || "// 这一页的最终出图提示词还没准备好。打开详情后会先整理本页版式，再生成最终提示词。";

  el.workflowDetailUseBtn.disabled = page.status === "running" || !page.pagePrompt;
  el.workflowDetailCopyBtn.disabled = page.status === "running" || !page.pagePrompt;
  el.workflowDetailRunBtn.disabled = page.status === "running" || state.workflowRunning || page.layoutStatus === "running";
  el.workflowDetailRunBtn.textContent = page.status === "running" ? "生成中..." : "生成本页";

  el.workflowDetailUseBtn.onclick = () => {
    el.prompt.value = page.pagePrompt || "";
    saveSettings();
    updatePayloadPreview();
    setStatus(`已将第 ${page.pageNumber} 页提示词写入主提示词。`, "success");
  };
  el.workflowDetailCopyBtn.onclick = async () => {
    try {
      await navigator.clipboard.writeText(page.pagePrompt || "");
      setStatus(`第 ${page.pageNumber} 页提示词已复制。`, "success");
    } catch (error) {
      setStatus(error.message || "复制提示词失败。", "error");
    }
  };
  el.workflowDetailRunBtn.onclick = () => generateWorkflowPage(page.id);
  el.workflowDetailCloseBtn.onclick = closeWorkflowDetail;
}

function renderWorkflowPlanSummary() {
  if (!state.workflowPages.length) {
    el.workflowPlanSummary.classList.add("hidden");
    el.workflowPlanSummary.innerHTML = "";
    return;
  }

  const theme = getCurrentThemePreset();
  const coverCount = state.workflowPages.filter((page) => page.pageType === "cover").length;
  const contentCount = state.workflowPages.filter((page) => page.pageType === "content").length;
  const dataCount = state.workflowPages.filter((page) => page.pageType === "data").length;

  el.workflowPlanSummary.classList.remove("hidden");
  el.workflowPlanSummary.innerHTML = `
    <div class="assistant-grid">
      <div class="assistant-block">
        <h3>拆分页摘要</h3>
        <p>${escapeHtml(state.workflowPlanSummary || "已生成逐页规划。")}\n拆分页完成后，系统会自动继续整理每一页的版式安排和最终提示词。</p>
      </div>
      <div class="assistant-block">
        <h3>页面统计</h3>
        <p>共 ${state.workflowPages.length} 页，其中封面 ${coverCount} 页、内容 ${contentCount} 页、数据 ${dataCount} 页。</p>
      </div>
      <div class="assistant-block">
        <h3>已确认的全局风格</h3>
        <p>${escapeHtml(theme.label)}\n${escapeHtml(truncateText(theme.basic, 140))}</p>
      </div>
      <div class="assistant-block">
        <h3>实际输出尺寸</h3>
        <p>${escapeHtml(getPptOutputDescription())}\n省事流的逐页生图会固定按这个比例发送，不再跟“预设 2K / 4K”混用。</p>
      </div>
    </div>
  `;
}

function renderWorkflowPlan() {
  renderWorkflowPlanSummary();
  const hasLayoutRunning = state.workflowPages.some((page) => page.layoutStatus === "running");
  el.workflowBatchBtn.disabled = !state.workflowPages.length || state.workflowRunning || hasLayoutRunning;
  el.workflowCopyBtn.disabled = !state.workflowPages.length;

  if (!state.workflowPages.length) {
    el.workflowPlanCards.innerHTML = "";
    closeWorkflowDetail();
    return;
  }

  el.workflowPlanCards.innerHTML = "";
  state.workflowPages.forEach((page) => {
    const card = document.createElement("article");
    card.className = "workflow-card";
    card.dataset.pageId = String(page.id || "");
    card.dataset.pageNumber = String(page.pageNumber || "");
    card.dataset.resultImageUrl = getWorkflowPreferredImageUrl(
      page,
      Array.isArray(page.resultImages) && page.resultImages.length ? page.resultImages[0] : "",
    );

    const tone = page.status === "error"
      ? "error"
      : page.status === "success"
        ? "success"
        : page.status === "running"
          ? "running"
          : "idle";

    card.innerHTML = `
      <div class="workflow-card-head">
        <div>
          <div class="workflow-page-meta">
            <span class="workflow-page-index">第 ${page.pageNumber} 页</span>
            <span class="workflow-type-chip" data-type="${page.pageType}">${PAGE_TYPE_LABELS[page.pageType] || page.pageType}</span>
          </div>
          <h3>${escapeHtml(
            typeof getWorkflowDisplayTitle === "function"
              ? (getWorkflowDisplayTitle(page) || page.pageTitle || `第 ${page.pageNumber} 页`)
              : (page.pageTitle || `第 ${page.pageNumber} 页`)
          )}</h3>
        </div>
        <div class="card-actions">
          <button type="button" class="ghost-btn view-workflow-page">查看详情</button>
          <button type="button" class="primary-btn run-workflow-page">${page.status === "running" ? "生成中..." : "生成本页"}</button>
        </div>
      </div>
      <div class="workflow-status" data-tone="${tone}">
        ${escapeHtml(
          page.status === "running"
            ? (page.error || "正在生成本页...")
            : page.status === "success"
              ? (page.error || "本页生成完成。")
              : page.status === "error"
                ? (page.error || "本页生成失败。")
                : "尚未生成"
        )}
      </div>
      <div class="workflow-section">
        <strong>最终上屏内容</strong>
        <p class="workflow-card-brief">${escapeHtml(
          typeof buildWorkflowVisibleTextPreview === "function"
            ? buildWorkflowVisibleTextPreview(page)
            : (page.pageContent || "")
        )}</p>
      </div>
      <div class="workflow-section">
        <strong>本页版式状态</strong>
        <p>${escapeHtml(
          page.layoutStatus === "ready"
            ? ((typeof buildWorkflowTextPolicySummary === "function"
              ? `${buildWorkflowTextPolicySummary(page)} `
              : "") + "这一页已经整理好个性化版式和可直接出图的页面提示词。")
            : page.layoutStatus === "running"
              ? (page.layoutError || "正在整理这一页的版式...")
              : page.layoutStatus === "error"
                ? (page.layoutError || "逐页版式整理失败，将回退到基础提示词。")
                : "尚未生成单独版式，拆分页完成后会先自动整理。"
        )}</p>
      </div>
    `;

    card.querySelector(".view-workflow-page").addEventListener("click", () => openWorkflowDetail(page.id));
    const runBtn = card.querySelector(".run-workflow-page");
    runBtn.disabled = page.status === "running" || state.workflowRunning || page.layoutStatus === "running";
    runBtn.addEventListener("click", () => generateWorkflowPage(page.id));

    if (page.resultImages.length) {
      const resultsWrap = document.createElement("div");
      resultsWrap.className = "result-images";

      page.resultImages.forEach((imageUrl, index) => {
        const saved = page.savedResults[imageUrl];
        const displayImageUrl = String(saved?.localUrl || imageUrl || "").trim();
        const resultCard = document.createElement("article");
        resultCard.className = "result-card";
        resultCard.dataset.imageUrl = displayImageUrl;
        resultCard.innerHTML = `
          <img src="${displayImageUrl}" alt="第 ${page.pageNumber} 页结果 ${index + 1}" />
          <div class="result-card-foot">
            <div>
              <strong>结果 ${index + 1}</strong>
              <div class="result-status">${saved ? `已保存到：${escapeHtml(saved.fileName)}` : "尚未保存到本地"}</div>
            </div>
            <div class="result-actions">
              <button type="button" class="ghost-btn preview-image-btn">打开原图</button>
              <button type="button" class="ghost-btn workflow-use-target">作为改图底图</button>
              <button type="button" class="ghost-btn workflow-save-local">下载到本地</button>
              ${saved ? `<a href="${saved.localUrl}" target="_blank" rel="noreferrer" class="ghost-btn">打开本地文件</a>` : ""}
            </div>
          </div>
        `;

        resultCard.querySelector(".preview-image-btn").addEventListener("click", () => {
          openImagePreview({
            title: `第 ${page.pageNumber} 页结果 ${index + 1} 预览`,
            imageUrl: displayImageUrl,
          });
        });

        resultCard.querySelector(".workflow-use-target").addEventListener("click", async () => {
          await addImage({
            name: `第 ${page.pageNumber} 页结果 ${index + 1}（改图底图）`,
            source: displayImageUrl,
          });
          state.editTargetImageId = state.images[state.images.length - 1]?.id || null;
          syncBboxToggles(true);
          saveSettings();
          renderImages();
          setStatus(`第 ${page.pageNumber} 页结果已加入输入区，并设为改图底图。`, "success");
        });

        resultCard.querySelector(".workflow-save-local").addEventListener("click", async () => {
          setStatus(`正在保存第 ${page.pageNumber} 页结果到 generated-images...`, "running");
          try {
            const data = await saveWorkflowPageResult(page.id, imageUrl, index);
            setStatus(`第 ${page.pageNumber} 页结果已保存到 ${data.savedPath}`, "success");
          } catch (error) {
            setStatus(error.message || "保存本页结果失败。", "error");
          }
        });

        resultsWrap.appendChild(resultCard);
      });

      card.appendChild(resultsWrap);
    }

    el.workflowPlanCards.appendChild(card);
  });

  renderWorkflowDetail();
}

async function sendAssistantRequest() {
  if (state.activeTabs.main !== "revise") {
    try {
      requireConfirmedTheme();
    } catch (error) {
      setStatus(error.message, "error");
      return;
    }
  }

  let payload;
  try {
    payload = buildAssistantPayload();
  } catch (error) {
    setStatus(error.message, "error");
    return;
  }

  if (!el.apiKey.value.trim()) return setStatus("请先填写 API Key。", "error");
  setStatus("正在调用 Qwen 提示词助手...", "running");
  el.assistantPreview.textContent = formatJSON(sanitizeForPreview(payload));

  try {
    const response = await fetch("/api/assistant", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ apiKey: el.apiKey.value.trim(), region: el.region.value, payload }),
    });
    const data = await response.json();
    syncApiKeyFeedback(response, data);
    if (!response.ok || data.code) {
      state.assistantParsed = null;
      state.assistantRaw = formatJSON(data);
      el.assistantPreview.textContent = state.assistantRaw;
      renderAssistantResult();
      return setStatus(data.message || "Qwen 助手调用失败。", "error");
    }

    const rawText = extractAssistantMessageText(data);
    state.assistantRaw = rawText || formatJSON(data);
    el.assistantPreview.textContent = state.assistantRaw;
    try {
      state.assistantParsed = normalizeAssistantResult(JSON.parse(rawText));
    } catch {
      state.assistantParsed = null;
    }
    renderAssistantResult();
    setStatus("Qwen 提示词建议已生成。", "success");
  } catch (error) {
    state.assistantParsed = null;
    renderAssistantResult();
    setStatus(error.message || "Qwen 助手请求失败。", "error");
  }
}

async function planWorkflowPages() {
  try {
    requireConfirmedTheme();
  } catch (error) {
    setStatus(error.message, "error");
    return;
  }

  let payload;
  try {
    payload = buildWorkflowAssistantPayload();
  } catch (error) {
    setStatus(error.message, "error");
    return;
  }
  if (!el.apiKey.value.trim()) return setStatus("请先填写 API Key。", "error");

  setButtonLoading(el.workflowPlanBtn, true, "拆分页中...");
  setProgressLoading(el.workflowProgress, true);
  setStatus("正在调用 Qwen 拆分页，并为后续逐页版式整理准备内容...", "running");
  state.workflowPlanRaw = formatJSON(sanitizeForPreview(payload));

  try {
    const response = await fetch("/api/assistant", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        apiKey: el.apiKey.value.trim(),
        region: el.region.value,
        payload,
      }),
    });
    const data = await response.json();
    syncApiKeyFeedback(response, data);
    if (!response.ok || data.code) {
      state.workflowPlanSummary = "";
      state.workflowPages = [];
      state.workflowDetailPageId = null;
      renderWorkflowPlan();
      return setStatus(data.message || "PPT 拆分页调用失败。", "error");
    }

    const rawText = extractAssistantMessageText(data);
    state.workflowPlanRaw = rawText || formatJSON(data);

    let normalized;
    try {
      normalized = normalizeWorkflowPlan(JSON.parse(rawText));
    } catch {
      normalized = { summary: "", pages: [] };
    }

    if (!normalized.pages.length) {
      state.workflowPlanSummary = "";
      state.workflowPages = [];
      state.workflowDetailPageId = null;
      renderWorkflowPlan();
      return setStatus("拆分页结果无法解析，请调整原文后重试。", "error");
    }

    state.workflowPlanSummary = normalized.summary || "已根据原文生成逐页规划。";
    state.workflowPages = normalized.pages;
    state.workflowDetailPageId = null;
    upsertWorkflowPlanLibraryEntry();
    saveSettings();
    renderWorkflowPlan();

    const { successCount, errorCount } = await prepareWorkflowPageLayouts(state.workflowPages);
    upsertWorkflowPlanLibraryEntry();
    saveSettings();
    renderWorkflowPlan();
    setStatus(
      errorCount
        ? `拆分页完成，已整理 ${successCount}/${state.workflowPages.length} 页版式，另有 ${errorCount} 页需要你手动重试。`
        : `拆分页和逐页版式安排已完成，共 ${state.workflowPages.length} 页，可直接查看详情或批量出图。`,
      errorCount ? "error" : "success",
    );
  } catch (error) {
    state.workflowPlanSummary = "";
    state.workflowPages = [];
    state.workflowPlanLibraryActiveId = "";
    state.workflowDetailPageId = null;
    renderWorkflowPlanLibrary();
    renderWorkflowPlan();
    setStatus(error.message || "PPT 拆分页请求失败。", "error");
  } finally {
    setButtonLoading(el.workflowPlanBtn, false);
    setProgressLoading(el.workflowProgress, false);
  }
}

async function sendRequest() {
  saveSettings();
  updatePayloadPreview();

  if (state.activeTabs.main === "manual") {
    try {
      requireConfirmedTheme();
    } catch (error) {
      setStatus(error.message, "error");
      return;
    }
  }

  let payload;
  try {
    payload = buildPayload();
  } catch (error) {
    setStatus(error.message, "error");
    return;
  }
  if (!el.apiKey.value.trim()) return setStatus("请先填写 API Key。", "error");
  if (el.bboxEnabled.checked) {
    const target = getEditTargetImage();
    if (!target) return setStatus("你已启用定点改图，但还没有选择改图底图。", "error");
    if (!target.boxes.length) return setStatus("你已启用定点改图，但当前底图还没有绘制任何 bbox 区域。", "error");
  }

  setStatus("正在调用 DashScope 图片模型...", "running");
  try {
    const { response, data } = await requestGeneration(payload);
    renderResponse(data);
    if (!response.ok || data.code) {
      stopPolling();
      state.currentTaskId = null;
      return setStatus(data.message || "调用失败。", "error");
    }
    if (el.requestMode.value === "async") {
      const taskId = data.output?.task_id;
      if (!taskId) return setStatus("异步任务创建成功，但没有返回 task_id。", "error");
      setStatus(`任务已创建，task_id：${taskId}，开始轮询。`, "running");
      return startPolling(taskId);
    }
    stopPolling();
    state.currentTaskId = null;
    setStatus("同步调用完成。", "success");
  } catch (error) {
    stopPolling();
    state.currentTaskId = null;
    setStatus(error.message || "请求失败。", "error");
  }
}

function bindEvents() {
  [
    el.apiKey, el.googleApiKey, el.requestMode, el.region, el.model, el.prompt, el.presetSize, el.imageCount, el.seed, el.enableSequential,
    el.thinkingMode, el.watermark, el.bboxEnabled, el.customWidth, el.customHeight, el.slideAspect, el.slideTemplate,
    el.includeRegionsInPrompt, el.assistantMode, el.assistantUseImages, el.assistantUseRegions, el.assistantUsePrompt,
    el.assistantThinking, el.assistantRequest, el.workflowPageCount, el.workflowContent, el.manualPageGoal, el.themeName, el.themeDecorationLevel,
  ].forEach((node) => {
    node.addEventListener("input", () => {
      if (node === el.apiKey || node === el.googleApiKey || node === el.model) {
        if (node === el.model) syncImageProviderUi();
        renderApiKeyStatus();
      }
      if (node === el.themeName) {
        el.workflowTheme.value = getThemeName();
        state.themeConfirmed = false;
        state.themeConfirmedSource = "";
        renderThemeStatus();
        syncWorkflowGate();
      }
      if (node === el.themeDecorationLevel) {
        renderManualLayoutPreview();
        renderWorkflowPlan();
        renderWorkflowDetail();
      }
      saveSettings();
      updatePayloadPreview();
      if (node === el.manualPageGoal) renderManualLayoutPreview();
    });
    node.addEventListener("change", () => {
      if (node === el.apiKey || node === el.googleApiKey || node === el.model) {
        if (node === el.model) syncImageProviderUi();
        renderApiKeyStatus();
      }
      if (node === el.themeName) {
        el.workflowTheme.value = getThemeName();
        state.themeConfirmed = false;
        state.themeConfirmedSource = "";
        renderThemeStatus();
        syncWorkflowGate();
      }
      if (node === el.themeDecorationLevel) {
        renderManualLayoutPreview();
        renderWorkflowPlan();
        renderWorkflowDetail();
      }
      saveSettings();
      updatePayloadPreview();
      if (node === el.bboxEnabled) renderImages();
      if (node === el.slideAspect) {
        renderSlidePlanner();
        if (hasConfirmedThemeDefinition() && state.workflowPages.length) {
          refreshWorkflowPagePrompts();
          renderWorkflowPlan();
        }
        renderManualLayoutPreview();
        renderThemeStatus();
        syncWorkflowGate();
      }
      if (node === el.includeRegionsInPrompt) {
        renderManualLayoutPreview();
      }
      if (node === el.manualPageGoal) renderManualLayoutPreview();
    });
  });

  el.testApiKeyBtn?.addEventListener("click", testApiKey);

  el.generateThemeBtn.addEventListener("click", async () => {
    try {
      await requestThemeDefinition(true);
    } catch (error) {
      renderThemeStatus();
      syncWorkflowGate();
      setStatus(error.message || "风格模板生成失败。", "error");
    }
  });

  el.themeReferenceInput?.addEventListener("change", async (event) => {
    try {
      await handleThemeReferenceFiles(event);
    } catch (error) {
      setStatus(error.message || "风格截图读取失败。", "error");
    }
  });
  el.learnThemeBtn?.addEventListener("click", async () => {
    try {
      if (!state.themeReferenceImages.length) {
        throw new Error("请先上传至少一张风格截图。");
      }
      await requestThemeDefinition(true, {
        mode: "reference",
        referenceImages: state.themeReferenceImages,
      });
    } catch (error) {
      renderThemeStatus();
      syncWorkflowGate();
      setStatus(error.message || "截图风格学习失败。", "error");
    }
  });

  el.confirmThemeBtn?.addEventListener("click", () => {
    if (!hasCurrentThemeDefinition()) {
      return setStatus("请先生成当前风格模板。", "error");
    }
    state.themeConfirmed = true;
    state.themeConfirmedSource = getThemeName();
    upsertThemeLibraryEntry({
      label: getThemeName(),
      definition: state.themeDefinition,
      raw: state.themeDefinitionRaw,
      confirmed: true,
    });
    saveSettings();
    if (state.workflowPages.length) {
      refreshWorkflowPagePrompts();
      renderWorkflowPlan();
    }
    renderThemeStatus();
    syncWorkflowGate();
    renderManualLayoutPreview();
    updatePayloadPreview();
    setStatus(`已确认“${getThemeName()}”风格模板。`, "success");
  });

  el.themeLibrarySelect?.addEventListener("change", () => {
    if (!el.themeLibrarySelect.value) {
      return;
    }
    try {
      applyThemeLibraryEntry(el.themeLibrarySelect.value);
      setStatus(`已切换到“${el.themeLibrarySelect.value}”风格。`, "success");
    } catch (error) {
      setStatus(error.message || "切换已生成风格失败。", "error");
    }
  });

  el.workflowPlanLibrarySelect?.addEventListener("change", () => {
    if (!el.workflowPlanLibrarySelect.value) return;
    try {
      applyWorkflowPlanLibraryEntry(el.workflowPlanLibrarySelect.value);
      setStatus("已载入这套拆分页，可以直接查看详情或继续生成。", "success");
    } catch (error) {
      setStatus(error.message || "载入已保存拆分页失败。", "error");
    }
  });

  el.workflowPlanDeleteBtn?.addEventListener("click", () => {
    if (!state.workflowPlanLibraryActiveId) {
      return setStatus("请先选中一套已保存拆分页。", "error");
    }
    try {
      deleteWorkflowPlanLibraryEntry(state.workflowPlanLibraryActiveId);
      setStatus("这套拆分页档案已经从资料库删除。", "success");
    } catch (error) {
      setStatus(error.message || "删除拆分页档案失败。", "error");
    }
  });

  el.autoReplaceTarget.addEventListener("change", saveSettings);
  el.bboxEnabled.addEventListener("change", () => {
    syncBboxToggles(el.bboxEnabled.checked);
    saveSettings();
    renderEditTargetControls();
    renderImages();
  });
  el.bboxEnabledImagePanel.addEventListener("change", () => {
    syncBboxToggles(el.bboxEnabledImagePanel.checked);
    saveSettings();
    renderEditTargetControls();
    renderImages();
    updatePayloadPreview();
  });
  el.editTargetSelect.addEventListener("change", () => {
    state.editTargetImageId = el.editTargetSelect.value || null;
    saveSettings();
    renderEditTargetControls();
    renderImages();
    updatePayloadPreview();
  });
  el.undoTargetBtn.addEventListener("click", undoEditTargetVersion);
  el.sizeMode.addEventListener("change", () => { saveSettings(); toggleSizeMode(); updatePayloadPreview(); });
  el.fileInput.addEventListener("change", handleFiles);
  el.addImageUrlBtn.addEventListener("click", handleAddImageUrl);
  el.addPaletteBtn.addEventListener("click", addPaletteRow);
  el.sendBtn.addEventListener("click", sendRequest);
  el.refreshTaskBtn.addEventListener("click", refreshCurrentTask);
  el.copyPayloadBtn.addEventListener("click", copyPayload);
  el.downloadAllBtn.addEventListener("click", downloadAllResults);
  el.applyTemplateBtn.addEventListener("click", () => {
    state.slideRegions = buildTemplateRegions(el.slideTemplate.value);
    renderSlidePlanner();
    renderManualLayoutPreview();
    updatePayloadPreview();
    setStatus("已应用 PPT 版式模板。", "success");
  });
  el.syncSlideSizeBtn.addEventListener("click", () => {
    const aspect = getCurrentAspectMeta();
    el.sizeMode.value = "custom";
    el.customWidth.value = String(aspect.outputWidth);
    el.customHeight.value = String(aspect.outputHeight);
    toggleSizeMode();
    saveSettings();
    updatePayloadPreview();
    renderThemeStatus();
    syncWorkflowGate();
    setStatus(`已同步为 ${aspect.outputWidth}x${aspect.outputHeight}。`, "success");
  });
  el.insertRegionsBtn.addEventListener("click", () => {
    const text = buildSlideRegionText();
    if (!text) return setStatus("当前还没有 PPT 区域可插入。", "error");
    const current = el.prompt.value.trim();
    el.prompt.value = current ? `${current}\n\n${text}` : text;
    saveSettings();
    renderManualLayoutPreview();
    updatePayloadPreview();
    setStatus("已将区域说明追加到主提示词。", "success");
  });
  el.clearRegionsBtn.addEventListener("click", () => {
    state.slideRegions = [];
    renderSlidePlanner();
    renderManualLayoutPreview();
    updatePayloadPreview();
    setStatus("已清空 PPT 区域。", "success");
  });
  el.manualBuildPromptBtn.addEventListener("click", () => {
    try {
      requireConfirmedTheme();
    } catch (error) {
      return setStatus(error.message, "error");
    }
    if (!el.manualPageGoal.value.trim() && !state.slideRegions.length) {
      return setStatus("请先写页面目标，或先在左侧画出区域。", "error");
    }
    renderManualLayoutPreview();
    updatePayloadPreview();
    setStatus("该页提示词概览已更新，确认后再写入主提示词即可。", "success");
  });
  el.manualAppendPromptBtn.addEventListener("click", () => {
    try {
      requireConfirmedTheme();
    } catch (error) {
      return setStatus(error.message, "error");
    }
    if (!el.manualPageGoal.value.trim() && !state.slideRegions.length) {
      return setStatus("请先写页面目标，或先在左侧画出区域。", "error");
    }
    const prompt = buildManualPromptDraft();
    if (!prompt) return setStatus("请先填写页面目标或区域。", "error");
    el.prompt.value = prompt;
    saveSettings();
    renderManualLayoutPreview();
    updatePayloadPreview();
    setStatus("已将该页提示词概览写入主提示词。", "success");
  });
  el.manualCopyBriefBtn.addEventListener("click", async () => {
    try {
      await navigator.clipboard.writeText(el.manualLayoutPreview.textContent || "");
      setStatus("排版摘要已复制。", "success");
    } catch (error) {
      setStatus(error.message || "复制排版摘要失败。", "error");
    }
  });
  el.assistantSendBtn.addEventListener("click", sendAssistantRequest);
  el.assistantApplyBtn.addEventListener("click", () => applyAssistantPrompt(true));
  el.assistantAppendBtn.addEventListener("click", () => applyAssistantPrompt(false));
  el.workflowPlanBtn.addEventListener("click", planWorkflowPages);
  el.workflowBatchBtn.addEventListener("click", generateAllWorkflowPages);
  el.workflowCopyBtn.addEventListener("click", copyWorkflowPlan);
  el.workflowClearBtn.addEventListener("click", () => {
    clearWorkflowPlan();
    closeWorkflowDetail();
  });
  el.workflowDetailBackdrop?.addEventListener("click", closeWorkflowDetail);
  el.imagePreviewBackdrop?.addEventListener("click", closeImagePreview);
  el.imagePreviewCloseBtn?.addEventListener("click", closeImagePreview);
  el.imagePreviewUseTargetBtn?.addEventListener("click", () => {
    sendImagePreviewToRevise({ submitImmediately: false }).catch((error) => {
      setStatus(error.message || "设为改图底图失败。", "error");
    });
  });
  el.imagePreviewOpenReviseBtn?.addEventListener("click", () => {
    sendImagePreviewToRevise({ submitImmediately: false }).catch((error) => {
      setStatus(error.message || "打开改图工作台失败。", "error");
    });
  });
  el.imagePreviewEditBtn?.addEventListener("click", () => {
    sendImagePreviewToRevise({ submitImmediately: true }).catch((error) => {
      setStatus(error.message || "直接提交修改失败。", "error");
    });
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      if (state.imagePreview.url) closeImagePreview();
      if (state.workflowDetailPageId) closeWorkflowDetail();
    }
  });
}

async function requestThemeDefinition(force = true, options = {}) {
  let themeName = getThemeName();
  const referenceImages = Array.isArray(options.referenceImages) && options.referenceImages.length
    ? options.referenceImages
    : state.themeReferenceImages;

  if (!themeName) {
    if (!referenceImages.length) throw new Error("请先输入风格关键词。");
    themeName = buildAutoThemeReferenceLabel();
    if (el.themeName) el.themeName.value = themeName;
    if (el.workflowTheme) el.workflowTheme.value = themeName;
  }
  if (!el.apiKey.value.trim()) throw new Error("请先填写 API Key。");
  if (hasCurrentThemeDefinition(themeName) && !force) return state.themeDefinition;

  const triggerBtn = options.mode === "reference" ? el.learnThemeBtn : el.generateThemeBtn;
  setButtonLoading(triggerBtn, true, options.mode === "reference" ? "学习中..." : "生成中...");
  setProgressLoading(el.themeProgress, true);
  setStatus(
    options.mode === "reference"
      ? `正在根据截图学习“${themeName}”风格...`
      : `正在为“${themeName}”生成风格模板...`,
    "running",
  );
  if (el.themeStatus) {
    el.themeStatus.textContent = options.mode === "reference"
      ? `正在学习“${themeName}”的截图风格，请稍候...`
      : `正在生成“${themeName}”的风格模板，请稍候...`;
  }

  try {
    const payload = buildThemeDefinitionPayload(themeName, {
      mode: options.mode || "keyword",
      referenceImages,
    });
    const response = await fetch("/api/assistant", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ apiKey: el.apiKey.value.trim(), region: el.region.value, payload }),
    });
    const data = await response.json();
    syncApiKeyFeedback(response, data);
    if (!response.ok || data.code) throw new Error(data.message || "风格模板生成失败。");

    const rawText = extractAssistantMessageText(data);
    let parsed;
    try {
      parsed = JSON.parse(rawText);
    } catch (error) {
      throw new Error(error.message || "风格模板 JSON 解析失败。");
    }

    state.themeDefinition = normalizeThemeDefinition(parsed, themeName);
    state.themeDefinitionRaw = formatJSON(parsed);
    state.themeDefinitionSource = themeName;
    state.themeConfirmed = false;
    state.themeConfirmedSource = "";
    upsertThemeLibraryEntry({
      label: themeName,
      definition: state.themeDefinition,
      raw: state.themeDefinitionRaw,
      confirmed: false,
    });
    saveSettings();
    renderThemeStatus();
    renderThemeReferenceAssets();
    syncWorkflowGate();
    updatePayloadPreview();
    setStatus(
      options.mode === "reference"
        ? `已根据截图学习出“${themeName}”风格，请确认后继续。`
        : `“${themeName}”风格模板已生成，请确认后继续。`,
      "success",
    );
    return state.themeDefinition;
  } finally {
    setButtonLoading(triggerBtn, false);
    setProgressLoading(el.themeProgress, false);
  }
}

async function sendImagePreviewToRevise({ submitImmediately = false } = {}) {
  const imageUrl = state.imagePreview?.url;
  if (!imageUrl) {
    setStatus("当前没有可送去改图的图片。", "error");
    return;
  }

  await addImage({
    name: `${state.imagePreview.title || "图片"}（改图底图）`,
    source: imageUrl,
  });
  state.editTargetImageId = state.images[state.images.length - 1]?.id || null;
  state.activeTabs.main = "revise";
  syncBboxToggles(true);
  if (el.imagePreviewPrompt?.value.trim()) {
    el.prompt.value = el.imagePreviewPrompt.value.trim();
  }
  saveSettings();
  renderImages();
  renderEditTargetControls();
  closeImagePreview();
  if (typeof closeWorkflowDetail === "function") closeWorkflowDetail();
  applyTabState();
  updatePayloadPreview();
  document.querySelector('[data-main-panel="revise"]')?.scrollIntoView({ behavior: "smooth", block: "start" });

  if (submitImmediately) {
    await sendRequest();
  } else {
    setStatus("已把图片送入改图工作台，可以继续框选或直接修改。", "success");
  }
}

async function init() {
  enforceRegionOptions();
  loadSettings();
  await loadPersistentLibrary();
  renderThemeStatus();
  renderThemeReferenceAssets();
  renderWorkflowPlanLibrary();
  syncWorkflowGate();
  syncWorkflowPlanStopButton();
  syncBboxToggles(el.bboxEnabled.checked);
  toggleSizeMode();
  bindTabNavigation();
  bindEvents();
  setupSlideCanvas();
  renderPalette();
  renderSlidePlanner();
  renderWorkflowPlan();
  renderWorkflowDetail();
  renderImages();
  renderAssistantResult();
  renderManualLayoutPreview();
  applyTabState();
  updatePayloadPreview();
  setStatus("页面已就绪，可以开始配置、拆页和改图。");
}

function truncateText(value, limit = 220) {
  const text = String(value || "").trim();
  if (!text) return "";
  return text.length > limit ? `${text.slice(0, limit)}…` : text;
}

function hasCurrentThemeDefinition(themeName = getThemeName()) {
  return Boolean(themeName && state.themeDefinition && state.themeDefinitionSource === themeName);
}

function hasConfirmedThemeDefinition(themeName = getThemeName()) {
  return Boolean(
    themeName
    && hasCurrentThemeDefinition(themeName)
    && state.themeConfirmed
    && state.themeConfirmedSource === themeName,
  );
}

function getPptOutputSizeValue() {
  const aspect = getCurrentAspectMeta();
  return `${aspect.outputWidth}*${aspect.outputHeight}`;
}

function getPptOutputDescription() {
  const aspect = getCurrentAspectMeta();
  return `${aspect.label}（${aspect.outputWidth}x${aspect.outputHeight}）`;
}

function setButtonLoading(button, loading, loadingLabel) {
  if (!button) return;
  if (!button.dataset.idleLabel) {
    button.dataset.idleLabel = button.textContent.trim();
  }
  button.disabled = loading;
  button.classList.toggle("button-loading", loading);
  button.textContent = loading ? (loadingLabel || button.dataset.idleLabel) : button.dataset.idleLabel;
}

function setProgressLoading(node, loading) {
  if (!node) return;
  node.classList.toggle("hidden", !loading);
}

function simplifyStaticCopy() {
  document.querySelector(".hero h1")?.replaceChildren("PPT Image Studio");
  document.querySelector(".settings-panel h2")?.replaceChildren("连接");
  document.querySelector(".flow-nav-panel h2")?.replaceChildren("模式");
  document.querySelector(".theme-flow-panel h2")?.replaceChildren("风格");
  document.querySelector('.flow-panel[data-main-panel="smart"] .panel-head h2')?.replaceChildren("省事流");
  document.querySelector('.flow-panel[data-main-panel="manual"] .panel-head h2')?.replaceChildren("手动流");
  document.querySelector('.flow-panel[data-main-panel="revise"] .panel-head h2')?.replaceChildren("改图流");
  document.querySelector('[data-main-tab="smart"]')?.replaceChildren("01 省事流");
  document.querySelector('[data-main-tab="manual"]')?.replaceChildren("02 手动流");
  document.querySelector('[data-main-tab="revise"]')?.replaceChildren("03 改图流");
  document.querySelector(".file-context-panel .panel-head h2")?.replaceChildren("附件");
  document.querySelector("#requestLogBar .request-log-head strong")?.replaceChildren("记录");
  document.querySelector("#clearContextFilesBtn")?.replaceChildren("清空");
  document.querySelector("#clearRequestLogBtn")?.replaceChildren("清空");
  document.querySelector("#generateThemeBtn")?.replaceChildren("生成风格");
  document.querySelector("#learnThemeBtn")?.replaceChildren("截图学风格");
  document.querySelector("#confirmThemeBtn")?.replaceChildren("确认风格");
  document.querySelector("#testApiKeyBtn")?.replaceChildren("测试");
  document.querySelector("#workflowPlanBtn")?.replaceChildren("拆分页面");
  document.querySelector("#workflowBatchBtn")?.replaceChildren("批量生成");
  document.querySelector("#workflowClearBtn")?.replaceChildren("清空");
  document.querySelector("#workflowDetailConfirmContentBtn")?.replaceChildren("确认内容");
  const detailCard = document.querySelector("#workflowDetailSuggestedContent")?.closest(".theme-review-card");
  if (detailCard) {
    const strongs = detailCard.querySelectorAll("strong");
    if (strongs[0]) strongs[0].replaceChildren("上屏内容确认");
    if (strongs[1]) strongs[1].remove();
    const suggestedTitle = document.querySelector("#workflowDetailSuggestedContent")?.previousElementSibling;
    if (suggestedTitle && suggestedTitle.tagName === "STRONG") suggestedTitle.replaceChildren("系统建议上屏");
    const confirmedTitle = document.querySelector("#workflowDetailConfirmedContent")?.previousElementSibling;
    if (confirmedTitle && confirmedTitle.tagName === "STRONG") confirmedTitle.replaceChildren("已确认上屏预览");
  }
}

function renderApiKeyStatus(message, tone = "idle") {
  if (!el.apiKeyStatus) return;
  const key = el.apiKey.value.trim();
  const defaultMessage = key ? `已保存 · ${maskApiKey(key)}` : "未填写 API Key";
  el.apiKeyStatus.textContent = message || defaultMessage;
  el.apiKeyStatus.dataset.tone = tone;
}

function syncApiKeyFeedback(response, data) {
  const message = String(data?.message || "");
  if (response?.status === 401 || /invalid api-?key/i.test(message)) {
    renderApiKeyStatus("当前 Key 无效", "error");
    return;
  }

  if (response?.ok && el.apiKey.value.trim()) {
    renderApiKeyStatus();
  }
}

async function testApiKey() {
  const apiKey = el.apiKey.value.trim();
  if (!apiKey) {
    renderApiKeyStatus("请先填写 API Key", "error");
    setStatus("请先填写 API Key", "error");
    return;
  }

  saveSettings();

  const payload = {
    model: "qwen3.6-plus",
    input: {
      messages: [
        {
          role: "user",
          content: [{ text: "请只回复 OK" }],
        },
      ],
    },
    parameters: {
      result_format: "message",
      enable_thinking: false,
    },
  };

  setButtonLoading(el.testApiKeyBtn, true, "测试中...");
  renderApiKeyStatus(`测试中 · ${maskApiKey(apiKey)}`, "running");
  setStatus("正在测试 API Key", "running");

  try {
    const response = await fetch("/api/assistant", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        apiKey,
        region: el.region.value,
        payload,
      }),
    });
    const data = await response.json();
    syncApiKeyFeedback(response, data);

    if (!response.ok || data.code) {
      const requestId = data?.request_id ? ` · ${data.request_id}` : "";
      const message = data?.message || "API Key 不可用";
      renderApiKeyStatus(`不可用 · ${message}${requestId}`, "error");
      setStatus(message, "error");
      return;
    }

    renderApiKeyStatus(`可用 · ${maskApiKey(apiKey)}`, "success");
    setStatus("API Key 可用", "success");
  } catch (error) {
    renderApiKeyStatus(`测试失败 · ${error.message || "网络失败"}`, "error");
    setStatus(error.message || "API Key 测试失败", "error");
  } finally {
    setButtonLoading(el.testApiKeyBtn, false);
  }
}

function renderThemeStatus() {
  const themeName = getThemeName();
  const aspect = getCurrentAspectMeta();
  const outputLabel = `${aspect.label} · ${aspect.outputWidth}×${aspect.outputHeight}`;

  if (!themeName) {
    el.themeStatus.textContent = "未生成";
  } else if (hasConfirmedThemeDefinition(themeName)) {
    el.themeStatus.textContent = `已确认 · ${themeName} · ${outputLabel}`;
  } else if (hasCurrentThemeDefinition(themeName)) {
    el.themeStatus.textContent = `待确认 · ${themeName}`;
  } else if (state.themeDefinitionSource && state.themeDefinitionSource !== themeName) {
    el.themeStatus.textContent = `需重生成 · ${themeName}`;
  } else {
    el.themeStatus.textContent = `未生成 · ${themeName}`;
  }

  if (el.themeDefinitionPreview) {
    el.themeDefinitionPreview.textContent = hasCurrentThemeDefinition(themeName)
      ? (state.themeDefinitionRaw || formatJSON(state.themeDefinition))
      : "";
  }

  if (el.workflowTheme) {
    el.workflowTheme.value = themeName;
  }

  renderThemeLibrary();
  renderThemeReviewPanel();
}

function renderWorkflowPlanSummary() {
  if (!state.workflowPages.length) {
    el.workflowPlanSummary.classList.add("hidden");
    el.workflowPlanSummary.innerHTML = "";
    return;
  }

  const coverCount = state.workflowPages.filter((page) => page.pageType === "cover").length;
  const contentCount = state.workflowPages.filter((page) => page.pageType === "content").length;
  const dataCount = state.workflowPages.filter((page) => page.pageType === "data").length;
  const themeLabel = getThemeName() || "未命名风格";
  const aspect = getCurrentAspectMeta();
  const decorationLabel = getDecorationLevelLabel(getGlobalDecorationLevel());

  el.workflowPlanSummary.classList.remove("hidden");
  el.workflowPlanSummary.innerHTML = `
    <div class="metric-strip">
      <span class="metric-pill">共 ${state.workflowPages.length} 页</span>
      <span class="metric-pill">封面 ${coverCount}</span>
      <span class="metric-pill">内容 ${contentCount}</span>
      <span class="metric-pill">数据 ${dataCount}</span>
      <span class="metric-pill">${escapeHtml(themeLabel)}</span>
      <span class="metric-pill">装饰 ${escapeHtml(decorationLabel)}</span>
      <span class="metric-pill">${aspect.label} · ${aspect.outputWidth}×${aspect.outputHeight}</span>
    </div>
  `;
}

function applyTabState() {
  const main = state.activeTabs.main || TAB_DEFAULTS.main;

  document.querySelectorAll("[data-main-tab]").forEach((button) => {
    const active = button.dataset.mainTab === main;
    button.classList.toggle("active", active);
    button.setAttribute("aria-selected", active ? "true" : "false");
  });

  document.querySelectorAll("[data-main-panel]").forEach((panel) => {
    panel.hidden = panel.dataset.mainPanel !== main;
  });

  if (main === "manual") {
    el.manualComposeHost.appendChild(el.craftStudioPanel);
    el.craftStudioPanel.hidden = false;
    el.craftStudioPanel.classList.remove("revise-simple");
    el.manualHelperPanel.hidden = false;
    if (el.assistantPanel) el.assistantPanel.hidden = false;
    if (el.downloadAllBtn) el.downloadAllBtn.hidden = false;
    if (el.debugPanel) el.debugPanel.hidden = true;
    if (el.usageBox) {
      el.usageBox.classList.add("hidden");
      el.usageBox.innerHTML = "";
    }
    el.studioSectionLabel.textContent = "02";
    el.studioTitle.textContent = "提示词 / 结果";
    el.studioHint.textContent = "";
    el.promptLabel.textContent = "主提示词";
    el.sendBtn.textContent = "生成";
    el.prompt.rows = 7;
    renderSlideCanvas();
  } else if (main === "revise") {
    el.reviseComposeHost.appendChild(el.craftStudioPanel);
    el.craftStudioPanel.hidden = false;
    el.craftStudioPanel.classList.add("revise-simple");
    el.manualHelperPanel.hidden = true;
    if (el.assistantPanel) el.assistantPanel.hidden = true;
    if (el.downloadAllBtn) el.downloadAllBtn.hidden = true;
    if (el.debugPanel) el.debugPanel.hidden = true;
    if (el.usageBox) {
      el.usageBox.classList.add("hidden");
      el.usageBox.innerHTML = "";
    }
    el.studioSectionLabel.textContent = "03";
    el.studioTitle.textContent = "提示词 / 提交";
    el.studioHint.textContent = "";
    el.promptLabel.textContent = "修改需求";
    el.sendBtn.textContent = "提交";
    el.prompt.rows = 10;
  } else {
    el.craftStudioPanel.hidden = true;
  }

  renderApiKeyStatus();
  renderManualLayoutPreview();
  renderEditTargetControls();
}

async function ensureWorkflowPageDesign(page, options = {}) {
  const { force = false, signal = null } = options;
  if (!page) throw new Error("未找到对应页面。");
  if (page.layoutStatus === "running" && page.layoutPromise) return page.layoutPromise;
  if (!force && page.pagePrompt && page.layoutStatus === "ready") return;
  if (!el.apiKey.value.trim()) throw new Error("请先填写 API Key。");
  if (signal?.aborted) throw signal.reason || createAbortError("已停止逐页版式整理。");

  page.layoutStatus = "running";
  page.layoutError = "正在整理这一页的版式和提示词...";
  renderWorkflowPlan();
  renderWorkflowDetail();

  page.layoutPromise = (async () => {
    try {
      const payload = buildWorkflowPageDesignPayload(page);
      const response = await fetch("/api/assistant", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        signal,
        body: JSON.stringify({
          apiKey: el.apiKey.value.trim(),
          region: el.region.value,
          payload,
        }),
      });
      const data = await response.json();
      syncApiKeyFeedback(response, data);
      if (!response.ok || data.code) {
        throw new Error(data.message || "逐页版式整理失败。");
      }

      const rawText = extractAssistantMessageText(data);
      let parsed;
      try {
        parsed = JSON.parse(rawText);
      } catch (error) {
        throw new Error(error.message || "逐页版式 JSON 解析失败。");
      }

      applyWorkflowPageDesign(page, parsed);
      upsertWorkflowPlanLibraryEntry();
      saveSettings();
      renderWorkflowPlan();
      renderWorkflowDetail();
    } catch (error) {
      if (isAbortError(error)) {
        page.layoutStatus = page.pagePrompt ? "ready" : "idle";
        page.layoutError = error.message || "已停止逐页版式整理。";
      } else {
        page.layoutStatus = "error";
        page.layoutError = error.message || "逐页版式整理失败。";
      }
      if (!page.pagePrompt) {
        page.pagePrompt = buildWorkflowPagePrompt(page);
      }
      renderWorkflowPlan();
      renderWorkflowDetail();
      throw error;
    } finally {
      page.layoutPromise = null;
    }
  })();

  return page.layoutPromise;
}

async function prepareWorkflowPageLayouts(pages = state.workflowPages, options = {}) {
  const { signal = null } = options;
  let successCount = 0;
  let errorCount = 0;

  for (let index = 0; index < pages.length; index += 1) {
    if (signal?.aborted) throw signal.reason || createAbortError("已停止逐页版式整理。");
    const page = pages[index];
    setStatus(`正在整理逐页版式：${index + 1}/${pages.length} · 第 ${page.pageNumber} 页`, "running");
    try {
      await ensureWorkflowPageDesign(page, { force: page.layoutStatus === "error", signal });
      successCount += 1;
      saveSettings();
    } catch (error) {
      if (isAbortError(error)) throw error;
      errorCount += 1;
    }
  }

  return { successCount, errorCount };
}

async function planWorkflowPages() {
  try {
    requireConfirmedTheme();
  } catch (error) {
    setStatus(error.message, "error");
    return;
  }

  let payload;
  try {
    payload = buildWorkflowAssistantPayload();
  } catch (error) {
    setStatus(error.message, "error");
    return;
  }
  if (!el.apiKey.value.trim()) {
    return setStatus("请先填写 API Key。", "error");
  }

  const controller = beginWorkflowPlanRun();
  setButtonLoading(el.workflowPlanBtn, true, "拆分页中...");
  setProgressLoading(el.workflowProgress, true);
  setStatus("正在调用 Qwen 拆分页，并为后续逐页版式整理准备内容...", "running");
  state.workflowPlanRaw = formatJSON(sanitizeForPreview(payload));

  try {
    const response = await fetch("/api/assistant", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      signal: controller.signal,
      body: JSON.stringify({
        apiKey: el.apiKey.value.trim(),
        region: el.region.value,
        payload,
      }),
    });
    const data = await response.json();
    syncApiKeyFeedback(response, data);
    if (!response.ok || data.code) {
      state.workflowPlanSummary = "";
      state.workflowPages = [];
      state.workflowDetailPageId = null;
      renderWorkflowPlan();
      return setStatus(data.message || "PPT 拆分页调用失败。", "error");
    }

    const rawText = extractAssistantMessageText(data);
    state.workflowPlanRaw = rawText || formatJSON(data);

    let normalized;
    try {
      normalized = normalizeWorkflowPlan(JSON.parse(rawText));
    } catch {
      normalized = { summary: "", pages: [] };
    }

    if (!normalized.pages.length) {
      state.workflowPlanSummary = "";
      state.workflowPages = [];
      state.workflowDetailPageId = null;
      renderWorkflowPlan();
      return setStatus("拆分页结果无法解析，请调整原文后重试。", "error");
    }

    state.workflowPlanSummary = normalized.summary || "已根据原文生成逐页规划。";
    state.workflowPages = normalized.pages;
    state.workflowDetailPageId = null;
    upsertWorkflowPlanLibraryEntry();
    saveSettings();
    renderWorkflowPlan();

    const { successCount, errorCount } = await prepareWorkflowPageLayouts(state.workflowPages, { signal: controller.signal });
    upsertWorkflowPlanLibraryEntry();
    saveSettings();
    renderWorkflowPlan();
    setStatus(
      errorCount
        ? `拆分页完成，已整理 ${successCount}/${state.workflowPages.length} 页版式，另有 ${errorCount} 页需要你手动重试。`
        : `拆分页和逐页版式安排已完成，共 ${state.workflowPages.length} 页，可直接查看详情或批量出图。`,
      errorCount ? "error" : "success",
    );
  } catch (error) {
    if (isAbortError(error)) {
      upsertWorkflowPlanLibraryEntry();
      saveSettings();
      renderWorkflowPlan();
      setStatus(
        state.workflowPages.length
          ? "已停止拆分页，当前已经拿到的页面和版式结果已保留。"
          : "已停止拆分页。",
        "idle",
      );
      return;
    }

    state.workflowPlanSummary = "";
    state.workflowPages = [];
    state.workflowPlanLibraryActiveId = "";
    state.workflowDetailPageId = null;
    renderWorkflowPlanLibrary();
    renderWorkflowPlan();
    setStatus(error.message || "PPT 拆分页请求失败。", "error");
  } finally {
    endWorkflowPlanRun(controller);
    setButtonLoading(el.workflowPlanBtn, false);
    setProgressLoading(el.workflowProgress, false);
  }
}

async function init() {
  enforceRegionOptions();
  ensureGoogleApiKeyField();
  loadSettings();
  await loadPersistentLibrary();
  syncImageProviderUi();
  renderApiKeyStatus();
  simplifyStaticCopy();
  renderThemeStatus();
  renderWorkflowPlanLibrary();
  syncWorkflowGate();
  syncBboxToggles(el.bboxEnabled.checked);
  toggleSizeMode();
  bindTabNavigation();
  bindEvents();
  setupSlideCanvas();
  renderPalette();
  renderSlidePlanner();
  renderWorkflowPlan();
  renderWorkflowDetail();
  renderImages();
  renderAssistantResult();
  renderManualLayoutPreview();
  applyTabState();
  updatePayloadPreview();
  setStatus("已就绪", "success");
}
