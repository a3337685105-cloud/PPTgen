const STORAGE_KEY = "ppt-studio-v2-mainline";
const DEFAULT_REGION = "beijing";
const PPT_MODEL = "nano-banana-2";
const GEMINI_WORKFLOW_MODELS = new Set([
  "gemini-3.1-flash-image-preview",
  "gemini-3-pro-image-preview",
  "gemini-2.5-flash-image",
]);
const GRSAI_WORKFLOW_MODELS = new Set([
  "nano-banana-2",
  "nano-banana-pro",
  "gemini-3.1-pro",
]);
const OPENAI_WORKFLOW_MODELS = new Set([
  "gpt-image-2",
]);
const EDIT_MODEL = "wan2.7-image-pro";
const MAX_REVISE_BOXES = 2;

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
  styleMode: { business: "商务汇报感", academic: "学术研究感", creative: "创意表达感" },
  layoutVariety: { uniform: "版式尽量统一", balanced: "版式适度变化", diverse: "版式明显变化" },
  detailLevel: { minimal: "细节尽量克制", polished: "细节精致均衡", rich: "细节明显丰富" },
  visualDensity: { airy: "留白明显更多", balanced: "留白信息均衡", dense: "信息明显更满" },
  compositionFocus: { imageLead: "画面更抢眼", balanced: "图文同等重要", textLead: "文字信息优先" },
  dataNarrative: { clean: "数据清楚直给", balanced: "数据图表增强", expressive: "数据更有视觉冲击" },
  pageMood: { steady: "稳重专业", modern: "现代清晰", dramatic: "强烈冲击" },
};

const AI_PROCESSING_MODE_LABELS = {
  strict: "Strict · 原汁原味",
  balanced: "Balanced · 适度润色",
  creative: "Creative · 深度扩写",
};

const PAGE_TYPE_META = {
  cover: { label: "封面", short: "封" },
  catalog: { label: "目录", short: "目" },
  chapter: { label: "章节", short: "章" },
  content: { label: "内容", short: "内" },
  data: { label: "数据", short: "数" },
};

const SPLIT_PRESETS = [
  {
    id: "balanced",
    label: "平衡标准",
    text: [
      "优先保证逻辑完整和单页单主题。",
      "普通内容页尽量控制在 50-150 字，关键转折允许适度少字。",
      "遇到时间线、对比关系、分类结构时优先拆成独立页。",
    ].join("\n"),
  },
  {
    id: "concise",
    label: "简洁讲解",
    text: [
      "优先少字与结论感。",
      "尽量把每页压缩成明确观点、短结论和少量支撑说明。",
      "内容过多时宁可拆页，也不要堆成密集信息墙。",
    ].join("\n"),
  },
  {
    id: "research",
    label: "研究细节",
    text: [
      "允许更多背景、数据和技术脉络。",
      "对方法、指标、机制、趋势要保留更多上下文。",
      "优先把数据、图表和引用拆成更清晰的独立页。",
    ].join("\n"),
  },
];

const ASPECT_META = {
  "16:9": { width: 1600, height: 900, outputSize: "3840*2160" },
  "4:3": { width: 1400, height: 1050, outputSize: "3840*2880" },
  "1:1": { width: 1200, height: 1200, outputSize: "3840*3840" },
};

const WORKFLOW_PROJECTS_VERSION = 1;
const THEME_PROMPT_SECTIONS = [
  { key: "basic", label: "基础风格" },
  { key: "cover", label: "封面模块" },
  { key: "catalog", label: "目录模块" },
  { key: "chapter", label: "章节模块" },
  { key: "content", label: "内容模块" },
  { key: "data", label: "数据模块" },
  { key: "modelPrompt", label: "总纲" },
];

const state = {
  activeTab: "smart",
  smartStep: "split",
  settings: {
    apiKey: "",
    googleApiKey: "",
    workflowImageModel: PPT_MODEL,
    enableGeminiGoogleSearch: false,
    grsaiHost: "domestic",
    region: DEFAULT_REGION,
    slideAspect: "16:9",
    outputSize: "2K",
    seed: "",
  },
  serverConfig: {
    loaded: false,
    configuredKeys: {
      dashscope: false,
      hostedImage: false,
    },
  },
  workspaceZoom: 100,
  themeName: "",
  decorationLevel: "medium",
  preferences: { ...DEFAULT_PREFERENCES },
  themeDefinition: null,
  themePromptTrace: null,
  selectedThemePromptSection: "basic",
  themeConfirmed: false,
  workflowContent: "",
  workflowPageCount: 8,
  aiProcessingMode: "balanced",
  workflowEnableExpansion: false,
  workflowTargetChars: 0,
  workflowMaxChars: 200,
  splitPresetId: "",
  splitTemplateText: "",
  parsedFiles: [],
  workflowJobId: "",
  workflowJob: null,
  workflowPollTimer: null,
  selectedPageId: "",
  pageDrafts: {},
  workflowProjectsIndex: [],
  workflowProjectSnapshots: {},
  selectedHistoryProjectId: "",
  pageDrawing: {
    tool: "pen",
    color: "#22d3ee",
    width: 6,
    active: false,
    pointerId: null,
    startX: 0,
    startY: 0,
    snapshot: null,
  },
  revise: {
    images: [],
    selectedImageId: "",
    prompt: "",
    results: [],
    drawing: null,
  },
};

const el = {};
const activeRequests = new Map();
const CANCEL_LABELS = {
  theme: "已取消生成风格。",
  split: "已取消拆分。",
  reprepare: "已取消重新整理。",
  repolish: "已取消 AI 一键重润。",
  pageGenerate: "已取消当前页生成。",
  batchGenerate: "已取消批量生成。",
  revise: "已取消改图。",
  testApi: "已取消 Key 测试。",
};

function cacheElements() {
  [
    "statusBar",
    "chainDescription",
    "workspaceZoomRange",
    "workspaceZoomValue",
    "themeName",
    "themeDecorationLevel",
    "prefStyleMode",
    "prefLayoutVariety",
    "prefDetailLevel",
    "prefVisualDensity",
    "prefCompositionFocus",
    "prefDataNarrative",
    "prefPageMood",
    "preferenceSummary",
    "generateThemeBtn",
    "cancelThemeBtn",
    "confirmThemeBtn",
    "goSplitBtn",
    "themeStatus",
    "themeSummaryPreview",
    "themePromptTabs",
    "themeModelPrompt",
    "quickApiKey",
    "quickGoogleApiKey",
    "quickGrsaiHost",
    "quickTestApiKeyBtn",
      "workflowPageCount",
    "workflowContent",
    "splitTemplateInput",
    "aiProcessingMode",
    "workflowEnableExpansion",
    "workflowTargetChars",
    "workflowMaxChars",
    "splitPresetToolbar",
    "pickReferenceFilesBtn",
    "referenceFilesInput",
    "referenceFilesList",
    "runSplitBtn",
    "cancelSplitBtn",
    "backToThemeBtn",
    "backToSplitBtn",
    "workflowSummary",
    "workflowStats",
    "workflowDiagnostics",
    "workflowPromptTrace",
    "workflowRibbonMeta",
    "workflowPageList",
    "pageMetaHint",
    "pageOnscreenPreview",
    "pageOnscreenEditor",
    "pageVisualElementsBlock",
    "pageVisualElementsDisplay",
    "repreparePageBtn",
    "aiRepolishPageBtn",
    "cancelRepreparePageBtn",
    "batchGenerateReadyBtn",
    "cancelBatchGenerateBtn",
    "uploadOverlayBtn",
    "overlayFileInput",
    "clearOverlayBtn",
    "slideStage",
    "slideFrame",
    "slideBaseImage",
    "slideEmptyState",
    "overlayLayer",
    "generateCurrentPageBtn",
    "cancelGenerateCurrentPageBtn",
    "pageExtraPrompt",
    "pagePromptTrace",
    "pageResultStrip",
    "viewCurrentPageLargeBtn",
    "copyPagePromptBtn",
    "exportWorkflowPptBtn",
    "historySummary",
    "historyProjectList",
    "historyProjectMeta",
    "historyPageGrid",
    "restoreHistoryProjectBtn",
    "pageImageModal",
    "pageImageModalImg",
    "closePageImageModalBtn",
    "revisePrevBtn",
    "reviseNextBtn",
    "reviseImportBtn",
    "reviseDeleteBtn",
    "reviseFileInput",
    "reviseImageName",
    "reviseImageCounter",
    "reviseBaseImage",
    "reviseCanvas",
    "reviseStage",
    "reviseEmptyState",
    "reviseThumbStrip",
    "revisePrompt",
    "sendReviseBtn",
    "cancelReviseBtn",
    "reviseResultStrip",
    "apiKey",
    "googleApiKey",
    "workflowImageModel",
    "enableGeminiGoogleSearch",
    "geminiGoogleSearchField",
    "grsaiHost",
    "region",
    "slideAspect",
    "outputSize",
    "seed",
    "testApiKeyBtn",
    "cancelTestApiKeyBtn",
  ].forEach((id) => {
    el[id] = document.getElementById(id);
  });
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;")
    .replaceAll("'", "&#39;");
}

function uid() {
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function safeJsonParse(text) {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

function setStatus(message, tone = "idle") {
  if (el.statusBar) {
    el.statusBar.textContent = message;
    el.statusBar.dataset.tone = tone;
  }
  const toast = el.statusToast || document.getElementById("statusToast");
  if (!toast || !message) return;
  toast.textContent = message;
  toast.dataset.tone = tone;
  toast.hidden = false;
  clearTimeout(setStatus.timer);
  const delay = tone === "error" ? 5200 : tone === "running" ? 0 : 2800;
  if (delay > 0) {
    setStatus.timer = setTimeout(() => {
      toast.hidden = true;
    }, delay);
  }
}

function setButtonLoading(button, loading, runningText) {
  if (!button) return;
  button.disabled = loading;
  if (loading) {
    button.dataset.idleText = button.textContent;
    button.textContent = runningText || "处理中...";
  } else if (button.dataset.idleText) {
    button.textContent = button.dataset.idleText;
  }
}

function startCancelableAction(key, button, cancelButton, runningText) {
  const previous = activeRequests.get(key);
  if (previous) {
    previous.controller.abort();
    activeRequests.delete(key);
  }
  const controller = new AbortController();
  activeRequests.set(key, { controller, button, cancelButton });
  setButtonLoading(button, true, runningText);
  if (cancelButton) {
    cancelButton.hidden = false;
    cancelButton.disabled = false;
  }
  return controller.signal;
}

function finishCancelableAction(key) {
  const active = activeRequests.get(key);
  if (!active) return;
  if (active.cancelButton) {
    active.cancelButton.hidden = true;
    active.cancelButton.disabled = true;
  }
  setButtonLoading(active.button, false);
  activeRequests.delete(key);
}

function cancelAction(key, message = CANCEL_LABELS[key] || "已取消当前请求。") {
  const active = activeRequests.get(key);
  if (!active) return false;
  active.controller.abort();
  setStatus(message, "idle");
  return true;
}

function isAbortError(error) {
  return error?.name === "AbortError" || /abort/i.test(String(error?.message || ""));
}

function isMissingWorkflowJobError(error) {
  return String(error?.message || "").includes("找不到对应的工作流任务");
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function syncGeminiSearchControls() {
  const geminiSelected = usingGeminiWorkflowModel();
  if (el.geminiGoogleSearchField) {
    el.geminiGoogleSearchField.hidden = !geminiSelected;
  }
  if (!geminiSelected) {
    state.settings.enableGeminiGoogleSearch = false;
    if (el.enableGeminiGoogleSearch) {
      el.enableGeminiGoogleSearch.checked = false;
    }
  }
}

function buildProjectTitle(job) {
  const firstPageTitle = job?.pages?.find((page) => page.pageTitle)?.pageTitle || "";
  const themeName = job?.themeDefinition?.themeName || state.themeName || "";
  return String(firstPageTitle || themeName || `项目 ${new Date().toLocaleString("zh-CN")}`).trim();
}

function buildProjectIndexEntry(job) {
  const title = buildProjectTitle(job);
  const pages = Array.isArray(job?.pages) ? job.pages : [];
  return {
    version: WORKFLOW_PROJECTS_VERSION,
    jobId: job.id,
    title,
    createdAt: job.createdAt || new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    totalPages: pages.length,
    themeName: job?.themeDefinition?.themeName || state.themeName || "",
    pageSummaries: pages.map((page) => {
      const resultImages = Array.isArray(page.resultImages)
        ? Array.from(new Set(page.resultImages.filter(Boolean)))
        : (page.baseImage ? [page.baseImage] : []);
      return {
        pageId: page.id,
        pageNumber: page.pageNumber,
        pageTitle: page.pageTitle || "",
        generated: Boolean(page.generated || resultImages.length),
        baseImage: page.baseImage || resultImages[0] || "",
        resultImages,
      };
    }),
  };
}

function buildProjectSnapshot(job) {
  if (!job?.id) return null;
  return {
    version: WORKFLOW_PROJECTS_VERSION,
    jobId: job.id,
    savedAt: new Date().toISOString(),
    themeName: state.themeName,
    decorationLevel: state.decorationLevel,
    preferences: { ...state.preferences },
    workflowContent: state.workflowContent,
    workflowPageCount: state.workflowPageCount,
    aiProcessingMode: state.aiProcessingMode,
    workflowEnableExpansion: state.workflowEnableExpansion,
    workflowTargetChars: state.workflowTargetChars,
    workflowMaxChars: state.workflowMaxChars,
    parsedFiles: Array.isArray(state.parsedFiles) ? state.parsedFiles : [],
    workflowJob: sanitizeRecoveredWorkflowJob(JSON.parse(JSON.stringify(job))),
    selectedPageId: state.selectedPageId || "",
    pageDrafts: serializePageDraftsForStorage(),
  };
}

function syncActiveProjectSnapshot() {
  if (!state.workflowJob?.id) return;
  const entry = buildProjectIndexEntry(state.workflowJob);
  state.workflowProjectSnapshots[state.workflowJob.id] = buildProjectSnapshot(state.workflowJob);
  state.workflowProjectsIndex = [
    entry,
    ...state.workflowProjectsIndex.filter((item) => item?.jobId && item.jobId !== state.workflowJob.id),
  ];
  if (!state.selectedHistoryProjectId) {
    state.selectedHistoryProjectId = state.workflowJob.id;
  }
}

function serializePageDraftsForStorage() {
  return Object.fromEntries(Object.entries(state.pageDrafts).map(([pageId, draft]) => [
    pageId,
    {
      onscreenTitle: draft.onscreenTitle || "",
      onscreenBody: draft.onscreenBody || "",
      onscreenContent: draft.onscreenContent || "",
      sourceOnscreenTitle: draft.sourceOnscreenTitle || "",
      sourceOnscreenContent: draft.sourceOnscreenContent || "",
      extraPrompt: draft.extraPrompt || "",
      overlays: (draft.overlays || []).filter((item) => /^https?:\/\//i.test(item.src) || item.src.startsWith("/generated-images/")),
    },
  ]));
}

function saveState() {
  syncActiveProjectSnapshot();
  const draftToStore = serializePageDraftsForStorage();
  const payload = {
    activeTab: state.activeTab,
    smartStep: state.smartStep,
    settings: state.settings,
    workspaceZoom: state.workspaceZoom,
    themeName: state.themeName,
    decorationLevel: state.decorationLevel,
    preferences: state.preferences,
    themeDefinition: state.themeDefinition,
    themePromptTrace: state.themePromptTrace,
    selectedThemePromptSection: state.selectedThemePromptSection,
    themeConfirmed: state.themeConfirmed,
    workflowContent: state.workflowContent,
    workflowPageCount: state.workflowPageCount,
    aiProcessingMode: state.aiProcessingMode,
    workflowEnableExpansion: state.workflowEnableExpansion,
    workflowTargetChars: state.workflowTargetChars,
    workflowMaxChars: state.workflowMaxChars,
    splitPresetId: state.splitPresetId,
    splitTemplateText: state.splitTemplateText,
    parsedFiles: state.parsedFiles,
    workflowJobId: state.workflowJobId,
    workflowJob: state.workflowJob,
    selectedPageId: state.selectedPageId,
    pageDrafts: draftToStore,
    workflowProjectsIndex: state.workflowProjectsIndex,
    workflowProjectSnapshots: state.workflowProjectSnapshots,
    selectedHistoryProjectId: state.selectedHistoryProjectId,
  };
  localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
}

function loadState() {
  const parsed = safeJsonParse(localStorage.getItem(STORAGE_KEY) || "");
  if (!parsed || typeof parsed !== "object") return;
  state.activeTab = ["smart", "history", "revise", "settings"].includes(parsed.activeTab) ? parsed.activeTab : "smart";
  state.smartStep = ["split", "theme", "pages"].includes(parsed.smartStep) ? parsed.smartStep : "split";
  if (state.smartStep === "theme" && !parsed.workflowJob && !parsed.themeDefinition) {
    state.smartStep = "split";
  }
  state.settings = { ...state.settings, ...(parsed.settings || {}) };
  state.workspaceZoom = clamp(Number(parsed.workspaceZoom || 100), 50, 140);
  state.themeName = String(parsed.themeName || "");
  state.decorationLevel = String(parsed.decorationLevel || "medium");
  state.preferences = { ...DEFAULT_PREFERENCES, ...(parsed.preferences || {}) };
  state.themeDefinition = parsed.themeDefinition || null;
  state.themePromptTrace = parsed.themePromptTrace || null;
  state.selectedThemePromptSection = String(parsed.selectedThemePromptSection || "basic");
  state.themeConfirmed = Boolean(parsed.themeConfirmed);
  state.workflowContent = String(parsed.workflowContent || "");
  state.workflowPageCount = clamp(Number(parsed.workflowPageCount || 8), 2, 120);
  state.aiProcessingMode = ["strict", "balanced", "creative"].includes(parsed.aiProcessingMode) ? parsed.aiProcessingMode : "balanced";
  state.workflowEnableExpansion = Boolean(parsed.workflowEnableExpansion);
  state.workflowTargetChars = clamp(Number(parsed.workflowTargetChars || 0), 0, 300);
  state.workflowMaxChars = clamp(Number(parsed.workflowMaxChars || 200), 0, 400);
  state.splitPresetId = "";
  state.splitTemplateText = "";
  state.parsedFiles = Array.isArray(parsed.parsedFiles) ? parsed.parsedFiles : [];
  state.workflowJobId = String(parsed.workflowJobId || "");
  state.workflowJob = parsed.workflowJob || null;
  state.selectedPageId = String(parsed.selectedPageId || "");
  state.pageDrafts = parsed.pageDrafts && typeof parsed.pageDrafts === "object" ? parsed.pageDrafts : {};
  state.workflowProjectsIndex = Array.isArray(parsed.workflowProjectsIndex) ? parsed.workflowProjectsIndex : [];
  state.workflowProjectSnapshots = parsed.workflowProjectSnapshots && typeof parsed.workflowProjectSnapshots === "object"
    ? parsed.workflowProjectSnapshots
    : {};
  state.selectedHistoryProjectId = String(parsed.selectedHistoryProjectId || "");
}

function applyStateToUi() {
  syncWorkflowModelOptions();
  el.workspaceZoomRange.value = String(state.workspaceZoom);
  applyWorkspaceZoom(state.workspaceZoom);
  el.themeName.value = state.themeName;
  el.themeDecorationLevel.value = state.decorationLevel;
  el.prefStyleMode.value = state.preferences.styleMode;
  el.prefLayoutVariety.value = state.preferences.layoutVariety;
  el.prefDetailLevel.value = state.preferences.detailLevel;
  el.prefVisualDensity.value = state.preferences.visualDensity;
  el.prefCompositionFocus.value = state.preferences.compositionFocus;
  el.prefDataNarrative.value = state.preferences.dataNarrative;
  el.prefPageMood.value = state.preferences.pageMood;
  el.workflowPageCount.value = String(state.workflowPageCount);
  el.aiProcessingMode.value = state.aiProcessingMode;
  if (el.workflowEnableExpansion) el.workflowEnableExpansion.checked = Boolean(state.workflowEnableExpansion);
  if (el.workflowTargetChars) el.workflowTargetChars.value = state.workflowTargetChars ? String(state.workflowTargetChars) : "";
  if (el.workflowMaxChars) el.workflowMaxChars.value = state.workflowMaxChars ? String(state.workflowMaxChars) : "";
  el.workflowContent.value = state.workflowContent;
  if (el.splitTemplateInput) el.splitTemplateInput.value = state.splitTemplateText;
  syncSplitExpansionControls();
  el.apiKey.value = state.settings.apiKey || "";
  el.googleApiKey.value = state.settings.googleApiKey || "";
  if (el.quickApiKey) el.quickApiKey.value = state.settings.apiKey || "";
  if (el.quickGoogleApiKey) el.quickGoogleApiKey.value = state.settings.googleApiKey || "";
  if (el.quickGrsaiHost) el.quickGrsaiHost.value = state.settings.grsaiHost || "domestic";
  el.workflowImageModel.value = state.settings.workflowImageModel || PPT_MODEL;
  if (el.enableGeminiGoogleSearch) {
    el.enableGeminiGoogleSearch.checked = Boolean(state.settings.enableGeminiGoogleSearch);
  }
  if (el.grsaiHost) el.grsaiHost.value = state.settings.grsaiHost || "domestic";
  el.region.value = state.settings.region || DEFAULT_REGION;
  el.slideAspect.value = state.settings.slideAspect || "16:9";
  el.outputSize.value = state.settings.outputSize || "2K";
  el.seed.value = state.settings.seed || "";
  el.revisePrompt.value = state.revise.prompt || "";
  syncGeminiSearchControls();
}

function syncSplitExpansionControls() {
  const enabled = Boolean(state.workflowEnableExpansion);
  if (el.workflowEnableExpansion) {
    el.workflowEnableExpansion.checked = enabled;
  }
  if (el.workflowTargetChars) {
    el.workflowTargetChars.disabled = !enabled;
    el.workflowTargetChars.placeholder = enabled ? "例如 180" : "勾选扩写后启用";
    el.workflowTargetChars.closest(".field")?.classList.toggle("is-disabled", !enabled);
  }
}

function applyWorkspaceZoom(value) {
  state.workspaceZoom = clamp(Number(value || 100), 50, 140);
  document.documentElement.style.setProperty("--workspace-zoom", (state.workspaceZoom / 100).toFixed(2));
  el.workspaceZoomValue.textContent = `${state.workspaceZoom}%`;
}

function getCurrentPreferences() {
  return {
    styleMode: el.prefStyleMode.value,
    layoutVariety: el.prefLayoutVariety.value,
    detailLevel: el.prefDetailLevel.value,
    visualDensity: el.prefVisualDensity.value,
    compositionFocus: el.prefCompositionFocus.value,
    dataNarrative: el.prefDataNarrative.value,
    pageMood: el.prefPageMood.value,
  };
}

function renderPreferenceSummary() {
  const current = getCurrentPreferences();
  state.preferences = current;
  el.preferenceSummary.textContent = [
    "当前默认：",
    PREFERENCE_LABELS.styleMode[current.styleMode],
    PREFERENCE_LABELS.layoutVariety[current.layoutVariety],
    PREFERENCE_LABELS.detailLevel[current.detailLevel],
    PREFERENCE_LABELS.visualDensity[current.visualDensity],
    PREFERENCE_LABELS.compositionFocus[current.compositionFocus],
    PREFERENCE_LABELS.dataNarrative[current.dataNarrative],
    PREFERENCE_LABELS.pageMood[current.pageMood],
  ].join("、");
}

function getPageTypeMeta(type) {
  const normalized = String(type || "").trim().toLowerCase();
  return PAGE_TYPE_META[normalized] || PAGE_TYPE_META.content;
}

function switchTab(tab) {
  if (tab !== "smart") {
    closeCurrentPageLargeImage();
  }
  state.activeTab = tab;
  document.querySelectorAll(".sidebar-tab").forEach((button) => {
    const active = button.dataset.tab === tab;
    button.classList.toggle("is-active", active);
  });
  document.querySelectorAll(".view").forEach((view) => {
    view.classList.toggle("is-active", view.dataset.view === tab);
  });
  if (tab === "history") {
    renderHistoryProjects();
  }
  saveState();
}

function switchSmartStep(step) {
  state.smartStep = step;
  document.querySelectorAll(".ribbon-step").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.step === step);
  });
  document.querySelectorAll(".smart-stage").forEach((panel) => {
    panel.classList.toggle("is-active", panel.dataset.stepPanel === step);
  });
  const meta = {
    split: "先把内容拆分和字数倾向定下来，再匹配稳定的成熟风格。",
    theme: "AI 会根据拆分后的页面结构自动选择风格基底，用户输入作为修饰。",
    pages: state.workflowJob
      ? `${state.workflowJob.readyToGeneratePages || 0} 页已可直接生成，${state.workflowJob.preparedPages || 0}/${state.workflowJob.totalPages || 0} 页已完成准备。`
      : "左侧看进度，中间改上屏内容，右侧直接生成当前页。",
  };
  if (el.workflowRibbonMeta) el.workflowRibbonMeta.textContent = meta[step];
  saveState();
}

function getSelectedPage() {
  if (!state.workflowJob?.pages?.length) return null;
  return state.workflowJob.pages.find((page) => page.id === state.selectedPageId) || state.workflowJob.pages[0];
}

function ensureSelectedPage() {
  const selected = getSelectedPage();
  if (selected) {
    state.selectedPageId = selected.id;
    ensurePageDraft(selected);
  } else {
    state.selectedPageId = "";
  }
}

function clearWorkflowSession({ toSplit = false } = {}) {
  stopWorkflowPolling();
  state.workflowJobId = "";
  state.workflowJob = null;
  state.selectedPageId = "";
  state.pageDrafts = {};
  if (toSplit) {
    switchSmartStep("split");
  } else {
    renderPagesWorkbench();
  }
  saveState();
}

function getCurrentWorkflowImageModel() {
  return String(state.settings.workflowImageModel || PPT_MODEL).trim() || PPT_MODEL;
}

function usingGeminiWorkflowModel() {
  return GEMINI_WORKFLOW_MODELS.has(getCurrentWorkflowImageModel());
}

function usingHostedWorkflowModel() {
  const model = getCurrentWorkflowImageModel();
  return GEMINI_WORKFLOW_MODELS.has(model) || GRSAI_WORKFLOW_MODELS.has(model) || OPENAI_WORKFLOW_MODELS.has(model);
}

function hasDashScopeApiKey() {
  return Boolean(state.settings.apiKey || state.serverConfig?.configuredKeys?.dashscope);
}

function hasHostedImageApiKey() {
  return Boolean(state.settings.googleApiKey || state.serverConfig?.configuredKeys?.hostedImage);
}

function syncWorkflowModelOptions() {
  if (!el.workflowImageModel) return;
  const currentValue = state.settings.workflowImageModel || el.workflowImageModel.value || PPT_MODEL;
  el.workflowImageModel.innerHTML = [
    `<option value="nano-banana-2">Grsai Nano Banana 2</option>`,
    `<option value="nano-banana-pro">Grsai Nano Banana Pro</option>`,
    `<option value="gemini-3.1-pro">Grsai Gemini 3.1 Pro</option>`,
    `<option value="gpt-image-2">OpenAI GPT Image 2</option>`,
    `<option value="gemini-3.1-flash-image-preview">Nano Banana 2</option>`,
    `<option value="gemini-3-pro-image-preview">Nano Banana Pro</option>`,
    `<option value="gemini-2.5-flash-image">Nano Banana</option>`,
    `<option value="wan2.7-image-pro">Wan 2.7</option>`,
  ].join("");
  el.workflowImageModel.value = currentValue;
  if (el.workflowImageModel.value !== currentValue) {
    el.workflowImageModel.value = PPT_MODEL;
  }
}

function sanitizeRecoveredWorkflowJob(job) {
  if (!job || !Array.isArray(job.pages)) return job;
  job.pages = job.pages.map((page) => {
    const normalizedPage = { ...page };
    if (normalizedPage.generationStatus === "running") {
      normalizedPage.generationStatus = normalizedPage.generated ? "done" : "idle";
    }
    if (!normalizedPage.onscreenContentText && normalizedPage.onscreenContent) {
      normalizedPage.onscreenContentText = normalizedPage.onscreenContent;
    }
    if (!normalizedPage.onscreenContent && normalizedPage.onscreenContentText) {
      normalizedPage.onscreenContent = normalizedPage.onscreenContentText;
    }
    normalizedPage.visualElementsPrompt = String(normalizedPage.visualElementsPrompt || "");
    normalizedPage.visualElementsDisplay = String(normalizedPage.visualElementsDisplay || normalizedPage.visualElementsPrompt || "");
    normalizedPage.resultImages = Array.isArray(normalizedPage.resultImages)
      ? Array.from(new Set(normalizedPage.resultImages.filter(Boolean)))
      : [];
    if (normalizedPage.baseImage && !normalizedPage.resultImages.includes(normalizedPage.baseImage)) {
      normalizedPage.resultImages.unshift(normalizedPage.baseImage);
    }
    normalizedPage.generated = Boolean(normalizedPage.generated || normalizedPage.resultImages.length);
    return normalizedPage;
  });
  return job;
}

function getWorkflowGenerationSize() {
  const aspectMeta = getAspectMeta();
  return usingHostedWorkflowModel()
    ? state.settings.outputSize
    : (aspectMeta?.outputSize || ASPECT_META["16:9"].outputSize);
}

function getPageGenerateRequestKey(pageId) {
  return `pageGenerate:${pageId}`;
}

function isPageGenerating(pageId) {
  return activeRequests.has(getPageGenerateRequestKey(pageId));
}

function ensurePageDraft(page) {
  if (!page) return null;
  const canonicalContent = formatOnscreenPreview(page.onscreenContentText || page.onscreenContent || page.pageContent || "");
  const canonicalSplit = splitOnscreenContentForEditor(canonicalContent, page.pageTitle || "");
  if (!state.pageDrafts[page.id]) {
    state.pageDrafts[page.id] = {
      onscreenTitle: canonicalSplit.title,
      onscreenBody: canonicalSplit.body,
      onscreenContent: composeOnscreenContentFromEditors(canonicalSplit.title, canonicalSplit.body),
      sourceOnscreenTitle: canonicalSplit.title,
      sourceOnscreenContent: composeOnscreenContentFromEditors(canonicalSplit.title, canonicalSplit.body),
      extraPrompt: page.extraPrompt || "",
      overlays: [],
      drawingLayer: "",
    };
  }
  if (typeof state.pageDrafts[page.id].onscreenTitle !== "string") {
    state.pageDrafts[page.id].onscreenTitle = "";
  }
  if (typeof state.pageDrafts[page.id].onscreenBody !== "string") {
    state.pageDrafts[page.id].onscreenBody = "";
  }
  if (typeof state.pageDrafts[page.id].sourceOnscreenTitle !== "string") {
    state.pageDrafts[page.id].sourceOnscreenTitle = "";
  }
  if (typeof state.pageDrafts[page.id].sourceOnscreenContent !== "string") {
    state.pageDrafts[page.id].sourceOnscreenContent = "";
  }
  if (!state.pageDrafts[page.id].onscreenContent && page.onscreenContent) {
    state.pageDrafts[page.id].onscreenContent = page.onscreenContent;
  }
  if (!state.pageDrafts[page.id].onscreenTitle && !state.pageDrafts[page.id].onscreenBody && state.pageDrafts[page.id].onscreenContent) {
    const localSplit = splitOnscreenContentForEditor(state.pageDrafts[page.id].onscreenContent, page.pageTitle || "");
    state.pageDrafts[page.id].onscreenTitle = localSplit.title;
    state.pageDrafts[page.id].onscreenBody = localSplit.body;
  }
  if (!state.pageDrafts[page.id].extraPrompt && page.extraPrompt) {
    state.pageDrafts[page.id].extraPrompt = page.extraPrompt;
  }
  if (!Array.isArray(state.pageDrafts[page.id].overlays)) {
    state.pageDrafts[page.id].overlays = [];
  }
  if (typeof state.pageDrafts[page.id].drawingLayer !== "string") {
    state.pageDrafts[page.id].drawingLayer = "";
  }
  return state.pageDrafts[page.id];
}

function syncPageDraftFromPage(page, { force = false } = {}) {
  const draft = ensurePageDraft(page);
  if (!page || !draft) return draft;

  const serverContent = formatOnscreenPreview(page.onscreenContentText || page.onscreenContent || page.pageContent || "");
  const serverSplit = splitOnscreenContentForEditor(serverContent, page.pageTitle || "");
  const serverComposite = composeOnscreenContentFromEditors(serverSplit.title, serverSplit.body);
  const previousServerContent = formatOnscreenPreview(draft.sourceOnscreenContent || "");
  const currentDraftContent = formatOnscreenPreview(
    composeOnscreenContentFromEditors(draft.onscreenTitle || "", draft.onscreenBody || "")
    || draft.onscreenContent
    || ""
  );
  const shouldAdoptServer =
    force
    || !currentDraftContent
    || !previousServerContent
    || currentDraftContent === previousServerContent;

  if (shouldAdoptServer) {
    draft.onscreenTitle = serverSplit.title;
    draft.onscreenBody = serverSplit.body;
    draft.onscreenContent = serverComposite;
  }

  draft.sourceOnscreenTitle = serverSplit.title;
  draft.sourceOnscreenContent = serverComposite;
  return draft;
}

function updateThemeView() {
  el.confirmThemeBtn.disabled = !state.themeDefinition;
  el.goSplitBtn.disabled = !state.themeConfirmed && !state.workflowJob;
  el.themeSummaryPreview.textContent = state.themeDefinition?.displaySummaryZh || "风格摘要会显示在这里。";
  el.themeModelPrompt.textContent = state.themePromptTrace
    ? stringifyTrace(state.themePromptTrace)
    : (state.themeDefinition?.modelPrompt || "还没有生成模型总纲。");
}

function renderThemePromptModules() {
  if (!el.themePromptTabs || !el.themeModelPrompt) return;
  const themeDefinition = state.themeDefinition || null;
  if (!themeDefinition) {
    el.themePromptTabs.innerHTML = "";
    el.themeModelPrompt.textContent = "还没有生成主题模块提示词。";
    return;
  }
  const availableSections = THEME_PROMPT_SECTIONS
    .map((section) => ({
      ...section,
      value: String(themeDefinition?.[section.key] || "").trim(),
    }))
    .filter((section) => section.value);
  if (!availableSections.length) {
    el.themePromptTabs.innerHTML = "";
    el.themeModelPrompt.textContent = "还没有可展示的主题模块提示词。";
    return;
  }
  if (!availableSections.some((section) => section.key === state.selectedThemePromptSection)) {
    state.selectedThemePromptSection = availableSections[0].key;
  }
  el.themePromptTabs.innerHTML = availableSections.map((section) => `
    <button
      class="template-chip theme-prompt-chip ${section.key === state.selectedThemePromptSection ? "is-active" : ""}"
      type="button"
      data-theme-prompt-section="${section.key}"
    >
      ${escapeHtml(section.label)}
    </button>
  `).join("");
  el.themePromptTabs.querySelectorAll("[data-theme-prompt-section]").forEach((button) => {
    button.addEventListener("click", () => {
      state.selectedThemePromptSection = button.dataset.themePromptSection || availableSections[0].key;
      renderThemePromptModules();
      saveState();
    });
  });
  const activeSection = availableSections.find((section) => section.key === state.selectedThemePromptSection) || availableSections[0];
  el.themeModelPrompt.textContent = activeSection.value;
}

const baseUpdateThemeView = updateThemeView;
updateThemeView = function updateThemeViewWithModuleTabs() {
  baseUpdateThemeView();
  renderThemePromptModules();
};

function renderSplitPresets() {
  if (!el.splitPresetToolbar || !el.splitTemplateInput) {
    state.splitPresetId = "";
    state.splitTemplateText = "";
    return;
  }
  el.splitPresetToolbar.innerHTML = SPLIT_PRESETS.map((preset) => `
    <button class="template-chip ${preset.id === state.splitPresetId ? "is-active" : ""}" type="button" data-preset-id="${preset.id}">
      ${escapeHtml(preset.label)}
    </button>
  `).join("");
  el.splitPresetToolbar.querySelectorAll("[data-preset-id]").forEach((button) => {
    button.addEventListener("click", () => {
      const preset = SPLIT_PRESETS.find((item) => item.id === button.dataset.presetId);
      if (!preset) return;
      state.splitPresetId = preset.id;
      state.splitTemplateText = preset.text;
      el.splitTemplateInput.value = preset.text;
      renderSplitPresets();
      saveState();
    });
  });
}

function renderHistoryProjectsLegacy() {
  if (!el.historyProjectList || !el.historySummary || !el.historyPageGrid || !el.historyProjectMeta) return;
  const projects = Array.isArray(state.workflowProjectsIndex) ? state.workflowProjectsIndex : [];
  if (!projects.length) {
    el.historySummary.textContent = "还没有历史项目。完成一次拆分后会自动收录到这里。";
    el.historyProjectList.innerHTML = `<div class="inline-hint">暂无历史生图项目</div>`;
    el.historyProjectMeta.textContent = "";
    el.historyPageGrid.innerHTML = "";
    if (el.restoreHistoryProjectBtn) el.restoreHistoryProjectBtn.disabled = true;
    return;
  }

  const selectedProjectId = state.selectedHistoryProjectId && state.workflowProjectSnapshots[state.selectedHistoryProjectId]
    ? state.selectedHistoryProjectId
    : projects[0].jobId;
  state.selectedHistoryProjectId = selectedProjectId;
  const selectedProject = projects.find((item) => item.jobId === selectedProjectId) || projects[0];
  const snapshot = state.workflowProjectSnapshots[selectedProject.jobId];

  el.historySummary.textContent = `共 ${projects.length} 个项目，按每次拆分归档。`;
  el.historyProjectList.innerHTML = projects.map((project) => `
    <div class="page-item history-project-item ${project.jobId === selectedProjectId ? "is-active" : ""}" data-history-project-id="${escapeHtml(project.jobId)}">
      <div class="history-project-title">
        <strong>${escapeHtml(project.title || "未命名项目")}</strong>
        <span class="status-pill ${project.jobId === state.workflowJobId ? "generated" : "idle"}">${project.jobId === state.workflowJobId ? "当前项目" : "历史项目"}</span>
      </div>
      <div class="page-meta">
        <span class="meta-pill">${escapeHtml(project.themeName || "未命名主题")}</span>
        <span class="meta-pill">${escapeHtml(`${project.totalPages || 0} 页`)}</span>
      </div>
      <div class="inline-hint">${escapeHtml(new Date(project.updatedAt || project.createdAt || Date.now()).toLocaleString("zh-CN"))}</div>
    </div>
  `).join("");

  el.historyProjectList.querySelectorAll("[data-history-project-id]").forEach((node) => {
    node.addEventListener("click", () => {
      state.selectedHistoryProjectId = node.dataset.historyProjectId;
      renderHistoryProjects();
      saveState();
    });
  });

  el.historyProjectMeta.textContent = selectedProject
    ? `${selectedProject.title || "未命名项目"} · ${selectedProject.themeName || "未命名主题"} · ${new Date(selectedProject.updatedAt || selectedProject.createdAt || Date.now()).toLocaleString("zh-CN")}`
    : "";

  const pageCards = Array.isArray(selectedProject?.pageSummaries) ? selectedProject.pageSummaries : [];
  el.historyPageGrid.innerHTML = pageCards.length
    ? pageCards.map((page) => `
      <div class="file-item history-page-card">
        ${page.baseImage ? `<img src="${escapeHtml(page.baseImage)}" alt="${escapeHtml(page.pageTitle || `第${page.pageNumber}页`)}" />` : `<div class="inline-summary">该页还没有生成结果</div>`}
        <strong>第${page.pageNumber}页 · ${escapeHtml(page.pageTitle || "未命名")}</strong>
        <span>${page.generated ? "已生成" : "未生成"}</span>
      </div>
    `).join("")
    : `<div class="inline-hint">这个项目还没有页面快照。</div>`;

  if (el.restoreHistoryProjectBtn) {
    el.restoreHistoryProjectBtn.disabled = !snapshot;
  }
}

function restoreHistoryProject() {
  const projectId = state.selectedHistoryProjectId;
  const snapshot = state.workflowProjectSnapshots[projectId];
  if (!snapshot) {
    setStatus("没有找到可恢复的历史项目。", "error");
    return;
  }
  stopWorkflowPolling();
  state.themeName = snapshot.themeName || "";
  state.decorationLevel = snapshot.decorationLevel || "medium";
  state.preferences = { ...DEFAULT_PREFERENCES, ...(snapshot.preferences || {}) };
  state.workflowContent = snapshot.workflowContent || "";
  state.workflowPageCount = clamp(Number(snapshot.workflowPageCount || 8), 2, 120);
  state.aiProcessingMode = ["strict", "balanced", "creative"].includes(snapshot.aiProcessingMode) ? snapshot.aiProcessingMode : "balanced";
  state.workflowEnableExpansion = Boolean(snapshot.workflowEnableExpansion);
  state.workflowTargetChars = clamp(Number(snapshot.workflowTargetChars || 0), 0, 300);
  state.workflowMaxChars = clamp(Number(snapshot.workflowMaxChars || 200), 0, 400);
  state.parsedFiles = Array.isArray(snapshot.parsedFiles) ? snapshot.parsedFiles : [];
  state.workflowJob = sanitizeRecoveredWorkflowJob(snapshot.workflowJob || null);
  state.workflowJobId = snapshot.jobId || state.workflowJob?.id || "";
  state.selectedPageId = snapshot.selectedPageId || state.workflowJob?.pages?.[0]?.id || "";
  state.pageDrafts = snapshot.pageDrafts && typeof snapshot.pageDrafts === "object" ? snapshot.pageDrafts : {};
  state.smartStep = state.workflowJob ? "pages" : "split";
  state.themeDefinition = state.workflowJob?.themeDefinition || state.themeDefinition;
  state.themePromptTrace = state.workflowJob?.promptTrace?.themeCore || state.themePromptTrace;
  state.themeConfirmed = Boolean(state.themeDefinition);
  applyStateToUi();
  renderPreferenceSummary();
  renderReferenceFiles();
  updateThemeView();
  ensureSelectedPage();
  renderHistoryProjects();
  switchTab("smart");
  switchSmartStep(state.smartStep);
  renderPagesWorkbench();
  setStatus(`已恢复项目：${buildProjectTitle(state.workflowJob || { pages: [] })}`, "success");
  saveState();
}

function collectHistoryImagesFromSummary(page, snapshot) {
  const directImages = Array.isArray(page?.resultImages) ? page.resultImages.filter(Boolean) : [];
  if (directImages.length) return Array.from(new Set(directImages));
  const snapshotPage = Array.isArray(snapshot?.workflowJob?.pages)
    ? snapshot.workflowJob.pages.find((item) => item.id === page?.pageId)
    : null;
  const snapshotImages = Array.isArray(snapshotPage?.resultImages) ? snapshotPage.resultImages.filter(Boolean) : [];
  const fallbackImages = [snapshotPage?.baseImage, page?.baseImage].filter(Boolean);
  return Array.from(new Set([...snapshotImages, ...fallbackImages]));
}

function renderHistoryPageImages(page, snapshot) {
  const historyImages = collectHistoryImagesFromSummary(page, snapshot);
  if (!historyImages.length) {
    return '<div class="inline-summary">该页还没有生成结果</div>';
  }
  return `
    <div class="history-image-stack">
      <img class="history-image-primary" src="${escapeHtml(historyImages[0])}" alt="${escapeHtml(page.pageTitle || `第${page.pageNumber}页`)}" />
      <div class="history-image-thumbs">
        ${historyImages.map((src, index) => `
          <img
            src="${escapeHtml(src)}"
            alt="${escapeHtml(`${page.pageTitle || `第${page.pageNumber}页`} - ${index + 1}`)}"
            title="${escapeHtml(`第${index + 1}次生成`)}"
          />
        `).join("")}
      </div>
    </div>
  `;
}

function renderHistoryProjects() {
  if (!el.historyProjectList || !el.historySummary || !el.historyPageGrid || !el.historyProjectMeta) return;
  const projects = Array.isArray(state.workflowProjectsIndex) ? state.workflowProjectsIndex : [];
  if (!projects.length) {
    el.historySummary.textContent = "还没有历史项目。完成一次拆分后会自动收录到这里。";
    el.historyProjectList.innerHTML = `<div class="inline-hint">暂无历史生图项目</div>`;
    el.historyProjectMeta.textContent = "";
    el.historyPageGrid.innerHTML = "";
    if (el.restoreHistoryProjectBtn) el.restoreHistoryProjectBtn.disabled = true;
    return;
  }

  const selectedProjectId = state.selectedHistoryProjectId && state.workflowProjectSnapshots[state.selectedHistoryProjectId]
    ? state.selectedHistoryProjectId
    : projects[0].jobId;
  state.selectedHistoryProjectId = selectedProjectId;
  const selectedProject = projects.find((item) => item.jobId === selectedProjectId) || projects[0];
  const snapshot = state.workflowProjectSnapshots[selectedProject.jobId];

  el.historySummary.textContent = `共 ${projects.length} 个项目，按每次拆分归档。`;
  el.historyProjectList.innerHTML = projects.map((project) => `
    <div class="page-item history-project-item ${project.jobId === selectedProjectId ? "is-active" : ""}" data-history-project-id="${escapeHtml(project.jobId)}">
      <div class="history-project-title">
        <strong>${escapeHtml(project.title || "未命名项目")}</strong>
        <span class="status-pill ${project.jobId === state.workflowJobId ? "generated" : "idle"}">${project.jobId === state.workflowJobId ? "当前项目" : "历史项目"}</span>
      </div>
      <div class="page-meta">
        <span class="meta-pill">${escapeHtml(project.themeName || "未命名主题")}</span>
        <span class="meta-pill">${escapeHtml(`${project.totalPages || 0} 页`)}</span>
      </div>
      <div class="inline-hint">${escapeHtml(new Date(project.updatedAt || project.createdAt || Date.now()).toLocaleString("zh-CN"))}</div>
    </div>
  `).join("");

  el.historyProjectList.querySelectorAll("[data-history-project-id]").forEach((node) => {
    node.addEventListener("click", () => {
      state.selectedHistoryProjectId = node.dataset.historyProjectId;
      renderHistoryProjects();
      saveState();
    });
  });

  el.historyProjectMeta.textContent = selectedProject
    ? `${selectedProject.title || "未命名项目"} · ${selectedProject.themeName || "未命名主题"} · ${new Date(selectedProject.updatedAt || selectedProject.createdAt || Date.now()).toLocaleString("zh-CN")}`
    : "";

  const snapshotPages = Array.isArray(snapshot?.workflowJob?.pages) ? snapshot.workflowJob.pages : [];
  const pageCards = Array.isArray(selectedProject?.pageSummaries) && selectedProject.pageSummaries.length
    ? selectedProject.pageSummaries
    : snapshotPages.map((page) => ({
      pageId: page.id,
      pageNumber: page.pageNumber,
      pageTitle: page.pageTitle || "",
      generated: Boolean(page.generated || (page.resultImages || []).length || page.baseImage),
      baseImage: page.baseImage || "",
      resultImages: Array.isArray(page.resultImages) ? page.resultImages : [],
    }));

  el.historyPageGrid.innerHTML = pageCards.length
    ? pageCards.map((page) => {
      const historyCount = collectHistoryImagesFromSummary(page, snapshot).length;
      return `
        <div class="file-item history-page-card">
          ${renderHistoryPageImages(page, snapshot)}
          <strong>第${page.pageNumber}页 · ${escapeHtml(page.pageTitle || "未命名")}</strong>
          <span>${page.generated ? `已生成 ${historyCount} 张` : "未生成"}</span>
        </div>
      `;
    }).join("")
    : `<div class="inline-hint">这个项目还没有页面快照。</div>`;

  if (el.restoreHistoryProjectBtn) {
    el.restoreHistoryProjectBtn.disabled = !snapshot;
  }
}

function getNextWorkflowPage(currentPageId) {
  const pages = Array.isArray(state.workflowJob?.pages) ? state.workflowJob.pages : [];
  const currentIndex = pages.findIndex((page) => page.id === currentPageId);
  if (currentIndex < 0) return null;
  return pages[currentIndex + 1] || null;
}

function openCurrentPageLargeImage() {
  const page = getSelectedPage();
  if (!page?.baseImage || !el.pageImageModal || !el.pageImageModalImg) return;
  el.pageImageModalImg.src = page.baseImage;
  el.pageImageModal.hidden = false;
  document.body.style.overflow = "hidden";
}

function closeCurrentPageLargeImage() {
  if (!el.pageImageModal || !el.pageImageModalImg) return;
  el.pageImageModal.hidden = true;
  el.pageImageModalImg.src = "";
  document.body.style.overflow = "";
}

function buildWorkflowExportPayload() {
  const pages = Array.isArray(state.workflowJob?.pages) ? state.workflowJob.pages : [];
  return {
    projectTitle: buildProjectTitle(state.workflowJob || { pages: [] }),
    slideAspect: state.settings.slideAspect || "16:9",
    pages: pages.map((page, index) => {
      const draft = ensurePageDraft(page);
      const fallbackContent = formatOnscreenPreview(page.onscreenContentText || page.onscreenContent || page.pageContent || "");
      const onscreenTitle = normalizeDisplayText(draft?.onscreenTitle || page.pageTitle || "").split("\n")[0].trim();
      const onscreenBody = formatOnscreenPreview(draft?.onscreenBody || fallbackContent);
      return {
        pageId: page.id,
        pageNumber: page.pageNumber || index + 1,
        pageTitle: page.pageTitle || "",
        onscreenTitle,
        onscreenBody,
        onscreenContent: composeOnscreenContentFromEditors(onscreenTitle, onscreenBody),
        imageUrl: page.baseImage || "",
      };
    }),
  };
}

async function exportWorkflowPpt() {
  if (!state.workflowJob?.pages?.length) {
    setStatus("请先完成拆分，至少生成出页面结构后再导出 PPT。", "error");
    return;
  }

  const button = el.exportWorkflowPptBtn;
  const originalLabel = button?.textContent || "一键导出 PPT";
  if (button) {
    button.disabled = true;
    button.textContent = "导出中...";
  }

  setStatus("正在导出 PPT...", "running");
  try {
    const response = await fetch("/api/export-workflow-ppt", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(buildWorkflowExportPayload()),
    });
    const data = await response.json();
    if (!response.ok || data.code) {
      throw new Error(data.message || "导出 PPT 失败。");
    }

    const anchor = document.createElement("a");
    anchor.href = data.downloadUrl;
    anchor.download = data.fileName || "";
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    setStatus(`PPT 已导出：${data.fileName || "output.pptx"}`, "success");
  } catch (error) {
    setStatus(error.message || "导出 PPT 失败。", "error");
  } finally {
    if (button) {
      button.textContent = originalLabel;
    }
    syncCurrentPageGenerateUi();
  }
}

function stringifyTrace(trace) {
  if (!trace) return "还没有 prompt trace。";
  return JSON.stringify(trace, null, 2);
}

function getFinalPromptFromPage(page) {
  return String(page?.promptTrace?.finalImage?.prompt || "").trim();
}

async function copyTextToClipboard(text) {
  const value = String(text || "");
  if (!value) return false;
  if (navigator.clipboard?.writeText && window.isSecureContext) {
    await navigator.clipboard.writeText(value);
    return true;
  }
  const textarea = document.createElement("textarea");
  textarea.value = value;
  textarea.setAttribute("readonly", "");
  textarea.style.position = "fixed";
  textarea.style.left = "-9999px";
  document.body.appendChild(textarea);
  textarea.select();
  const ok = document.execCommand("copy");
  textarea.remove();
  return ok;
}

function getAspectMeta() {
  return ASPECT_META[state.settings.slideAspect] || ASPECT_META["16:9"];
}

function renderArtboard() {
  const page = getSelectedPage();
  const draft = page ? ensurePageDraft(page) : null;
  const baseImage = page?.baseImage || "";
  if (!page) {
    el.slideBaseImage.hidden = true;
    el.slideBaseImage.src = "";
    el.slideEmptyState.hidden = false;
    el.overlayLayer.innerHTML = "";
    renderPageDrawingLayer();
    return;
  }
  const hasVisualContent = Boolean(baseImage || draft?.drawingLayer || draft?.overlays?.length);
  if (baseImage) {
    el.slideBaseImage.src = baseImage;
    el.slideBaseImage.hidden = false;
  } else {
    el.slideBaseImage.hidden = true;
    el.slideBaseImage.src = "";
  }
  el.slideEmptyState.hidden = hasVisualContent;

  const overlays = draft?.overlays || [];
  renderPageDrawingLayer();
  el.overlayLayer.innerHTML = overlays.map((item) => `
    <div class="overlay-item" data-overlay-id="${item.id}" style="left:${item.x}%;top:${item.y}%;width:${item.w}%;height:${item.h}%;">
      <img src="${escapeHtml(item.src)}" alt="overlay" />
      <button class="overlay-resize-handle" type="button" data-overlay-resize-id="${item.id}" aria-label="调整补图大小"></button>
    </div>
  `).join("");

  el.overlayLayer.querySelectorAll(".overlay-item").forEach((node) => {
    const overlayId = node.dataset.overlayId;
    node.addEventListener("pointerdown", (event) => beginOverlayDrag(event, overlayId));
  });
  el.overlayLayer.querySelectorAll("[data-overlay-resize-id]").forEach((node) => {
    const overlayId = node.dataset.overlayResizeId;
    node.addEventListener("pointerdown", (event) => beginOverlayResize(event, overlayId));
  });
}

function resizePageDrawCanvas(forceRedraw = false) {
  if (!el.pageDrawCanvas || !el.slideStage) return;
  const rect = el.slideStage.getBoundingClientRect();
  const width = Math.max(1, Math.round(rect.width));
  const height = Math.max(1, Math.round(rect.height));
  if (el.pageDrawCanvas.width !== width || el.pageDrawCanvas.height !== height) {
    el.pageDrawCanvas.width = width;
    el.pageDrawCanvas.height = height;
    forceRedraw = true;
  }
  if (forceRedraw) {
    renderPageDrawingLayer();
  }
}

function renderPageDrawingLayer() {
  if (!el.pageDrawCanvas) return;
  const ctx = el.pageDrawCanvas.getContext("2d");
  if (!ctx) return;
  resizePageDrawCanvas(false);
  ctx.clearRect(0, 0, el.pageDrawCanvas.width, el.pageDrawCanvas.height);
  const page = getSelectedPage();
  const draft = ensurePageDraft(page);
  const layer = draft?.drawingLayer || "";
  if (!page || !layer) return;
  const renderKey = `${page.id}:${layer.slice(0, 48)}`;
  el.pageDrawCanvas.dataset.renderKey = renderKey;
  loadImage(layer)
    .then((image) => {
      if (el.pageDrawCanvas?.dataset.renderKey !== renderKey) return;
      const currentCtx = el.pageDrawCanvas.getContext("2d");
      if (!currentCtx) return;
      currentCtx.clearRect(0, 0, el.pageDrawCanvas.width, el.pageDrawCanvas.height);
      currentCtx.drawImage(image, 0, 0, el.pageDrawCanvas.width, el.pageDrawCanvas.height);
    })
    .catch(() => {});
}

function updatePageDrawToolbar() {
  if (!el.pageDrawPenBtn || !el.pageDrawRectBtn) return;
  const tool = state.pageDrawing?.tool || "pen";
  el.pageDrawPenBtn.classList.toggle("is-active", tool === "pen");
  el.pageDrawRectBtn.classList.toggle("is-active", tool === "rect");
  if (el.pageDrawColorInput && state.pageDrawing?.color) {
    el.pageDrawColorInput.value = state.pageDrawing.color;
  }
}

function savePageDrawingLayer() {
  const page = getSelectedPage();
  const draft = ensurePageDraft(page);
  if (!page || !draft || !el.pageDrawCanvas) return;
  draft.drawingLayer = el.pageDrawCanvas.toDataURL("image/png");
  saveState();
  renderArtboard();
}

function clearPageDrawingLayer() {
  const page = getSelectedPage();
  const draft = ensurePageDraft(page);
  if (!page || !draft) return;
  draft.drawingLayer = "";
  renderPageDrawingLayer();
  renderArtboard();
  saveState();
}

function setupPageDrawingInteractions() {
  if (!el.pageDrawCanvas || el.pageDrawCanvas.dataset.bound === "true") return;
  const drawingState = state.pageDrawing;
  const getPoint = (event) => {
    const rect = el.pageDrawCanvas.getBoundingClientRect();
    if (!rect.width || !rect.height) return null;
    return {
      x: ((event.clientX - rect.left) / rect.width) * el.pageDrawCanvas.width,
      y: ((event.clientY - rect.top) / rect.height) * el.pageDrawCanvas.height,
    };
  };
  const getCtx = () => el.pageDrawCanvas.getContext("2d");
  const applyStrokeStyle = (ctx) => {
    ctx.strokeStyle = state.pageDrawing.color || "#22d3ee";
    ctx.lineWidth = state.pageDrawing.width || 6;
    ctx.setLineDash([]);
    ctx.lineJoin = "round";
    ctx.lineCap = "round";
  };
  const applyRectStyle = (ctx) => {
    const color = state.pageDrawing.color || "#22d3ee";
    ctx.strokeStyle = color;
    ctx.fillStyle = `${color}26`;
    ctx.lineWidth = Math.max(4, state.pageDrawing.width || 6);
    ctx.setLineDash([10, 8]);
    ctx.lineJoin = "round";
    ctx.lineCap = "round";
  };

  const pointerDown = (event) => {
    if (event.button !== 0) return;
    const page = getSelectedPage();
    if (!page) return;
    resizePageDrawCanvas(false);
    const point = getPoint(event);
    const ctx = getCtx();
    if (!point || !ctx) return;
    drawingState.active = true;
    drawingState.pointerId = event.pointerId;
    drawingState.startX = point.x;
    drawingState.startY = point.y;
    drawingState.snapshot = ctx.getImageData(0, 0, el.pageDrawCanvas.width, el.pageDrawCanvas.height);
    if ((drawingState.tool || "pen") === "pen") {
      applyStrokeStyle(ctx);
      ctx.beginPath();
      ctx.moveTo(point.x, point.y);
    }
    el.pageDrawCanvas.setPointerCapture?.(event.pointerId);
  };

  const pointerMove = (event) => {
    if (!drawingState.active || drawingState.pointerId !== event.pointerId) return;
    const point = getPoint(event);
    const ctx = getCtx();
    if (!point || !ctx) return;
    const tool = drawingState.tool || "pen";
    if (tool === "pen") {
      applyStrokeStyle(ctx);
      ctx.lineTo(point.x, point.y);
      ctx.stroke();
      return;
    }
    if (tool === "rect" && drawingState.snapshot) {
      ctx.putImageData(drawingState.snapshot, 0, 0);
      applyRectStyle(ctx);
      const x = Math.min(drawingState.startX, point.x);
      const y = Math.min(drawingState.startY, point.y);
      const w = Math.abs(point.x - drawingState.startX);
      const h = Math.abs(point.y - drawingState.startY);
      ctx.fillRect(x, y, w, h);
      ctx.strokeRect(x, y, w, h);
      ctx.setLineDash([]);
    }
  };

  const pointerUp = (event) => {
    if (!drawingState.active) return;
    if (drawingState.pointerId != null && event.pointerId != null && drawingState.pointerId !== event.pointerId) return;
    drawingState.active = false;
    drawingState.pointerId = null;
    drawingState.snapshot = null;
    savePageDrawingLayer();
  };

  el.pageDrawCanvas.addEventListener("pointerdown", pointerDown);
  el.pageDrawCanvas.addEventListener("pointermove", pointerMove);
  el.pageDrawCanvas.addEventListener("pointerup", pointerUp);
  el.pageDrawCanvas.addEventListener("pointerleave", pointerUp);
  el.pageDrawPenBtn?.addEventListener("click", () => {
    state.pageDrawing.tool = "pen";
    updatePageDrawToolbar();
  });
  el.pageDrawRectBtn?.addEventListener("click", () => {
    state.pageDrawing.tool = "rect";
    updatePageDrawToolbar();
  });
  el.clearPageDrawingBtn?.addEventListener("click", clearPageDrawingLayer);
  el.pageDrawColorInput?.addEventListener("input", () => {
    state.pageDrawing.color = el.pageDrawColorInput.value || "#22d3ee";
    saveState();
  });
  el.pageDrawCanvas.dataset.bound = "true";
  updatePageDrawToolbar();
}

function beginOverlayDrag(event, overlayId) {
  if (event.target?.closest?.("[data-overlay-resize-id]")) return;
  const page = getSelectedPage();
  const draft = ensurePageDraft(page);
  const overlay = draft.overlays.find((item) => item.id === overlayId);
  if (!overlay) return;
  event.preventDefault();
  const stageRect = el.slideStage.getBoundingClientRect();
  const startX = event.clientX;
  const startY = event.clientY;
  const startLeft = overlay.x;
  const startTop = overlay.y;
  const move = (moveEvent) => {
    const dx = ((moveEvent.clientX - startX) / stageRect.width) * 100;
    const dy = ((moveEvent.clientY - startY) / stageRect.height) * 100;
    overlay.x = clamp(startLeft + dx, 0, Math.max(0, 100 - overlay.w));
    overlay.y = clamp(startTop + dy, 0, Math.max(0, 100 - overlay.h));
    renderArtboard();
  };
  const up = () => {
    window.removeEventListener("pointermove", move);
    window.removeEventListener("pointerup", up);
    saveState();
  };
  window.addEventListener("pointermove", move);
  window.addEventListener("pointerup", up);
}

function beginOverlayResize(event, overlayId) {
  event.preventDefault();
  event.stopPropagation();
  const page = getSelectedPage();
  const draft = ensurePageDraft(page);
  const overlay = draft?.overlays?.find((item) => item.id === overlayId);
  if (!overlay || !el.slideStage) return;

  const stageRect = el.slideStage.getBoundingClientRect();
  const startX = event.clientX;
  const startY = event.clientY;
  const startWidth = overlay.w;
  const startHeight = overlay.h;
  const maxWidth = 100 - overlay.x;
  const maxHeight = 100 - overlay.y;
  const minSize = 8;

  const move = (moveEvent) => {
    const dw = ((moveEvent.clientX - startX) / stageRect.width) * 100;
    const dh = ((moveEvent.clientY - startY) / stageRect.height) * 100;
    overlay.w = clamp(startWidth + dw, minSize, Math.max(minSize, maxWidth));
    overlay.h = clamp(startHeight + dh, minSize, Math.max(minSize, maxHeight));
    renderArtboard();
  };

  const up = () => {
    window.removeEventListener("pointermove", move);
    window.removeEventListener("pointerup", up);
    saveState();
  };

  window.addEventListener("pointermove", move);
  window.addEventListener("pointerup", up);
}

async function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error(`读取失败：${file.name}`));
    reader.readAsDataURL(file);
  });
}

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    if (/^https?:\/\//i.test(String(src || ""))) {
      image.crossOrigin = "anonymous";
    }
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("图片加载失败。"));
    image.src = src;
  });
}

async function exportCurrentArtboard() {
  const page = getSelectedPage();
  if (!page) return "";
  const draft = ensurePageDraft(page);
  if (!page.baseImage && !draft.drawingLayer && !draft.overlays.length) return "";
  try {
    const meta = getAspectMeta();
    const canvas = document.createElement("canvas");
    canvas.width = meta.width;
    canvas.height = meta.height;
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    if (page.baseImage) {
      const base = await loadImage(page.baseImage);
      ctx.drawImage(base, 0, 0, canvas.width, canvas.height);
    }
    if (draft.drawingLayer) {
      const drawing = await loadImage(draft.drawingLayer);
      ctx.drawImage(drawing, 0, 0, canvas.width, canvas.height);
    }
    for (const overlay of draft.overlays) {
      const image = await loadImage(overlay.src);
      ctx.drawImage(
        image,
        (overlay.x / 100) * canvas.width,
        (overlay.y / 100) * canvas.height,
        (overlay.w / 100) * canvas.width,
        (overlay.h / 100) * canvas.height,
      );
    }
    return canvas.toDataURL("image/png");
  } catch (error) {
    setStatus("当前底图包含跨域图片，已忽略画布叠加继续生成。", "error");
    return "";
  }
}

async function apiJson(url, options = {}) {
  const response = await fetch(url, options);
  const data = await response.json();
  if (!response.ok || data.code) {
    throw new Error(data.message || "请求失败。");
  }
  return data;
}

async function refreshServerConfig() {
  try {
    const response = await fetch("/api/health");
    if (!response.ok) return;
    const data = await response.json();
    state.serverConfig = {
      loaded: true,
      configuredKeys: {
        dashscope: Boolean(data?.configuredKeys?.dashscope),
        hostedImage: Boolean(data?.configuredKeys?.hostedImage),
      },
      workflowModels: data?.workflowModels || null,
    };
    syncCurrentPageGenerateUi();
  } catch (_error) {
    // Local env configuration is optional; keep UI usable if health probing fails.
  }
}

async function ensureServerConfigReady() {
  if (state.serverConfig?.loaded) return;
  await refreshServerConfig();
}

async function handleReferenceFiles(event) {
  const files = Array.from(event.target.files || []);
  if (!files.length) return;
  const formData = new FormData();
  files.forEach((file) => formData.append("files", file));
  setStatus(`正在解析 ${files.length} 个参考文件...`, "running");
  try {
    const response = await fetch("/api/files/parse", { method: "POST", body: formData });
    const data = await response.json();
    if (!response.ok || data.code) throw new Error(data.message || "文件解析失败。");
    state.parsedFiles = [...state.parsedFiles, ...(data.files || [])];
    renderReferenceFiles();
    setStatus("参考文件已加入拆分材料。", "success");
    saveState();
  } catch (error) {
    setStatus(error.message || "文件解析失败。", "error");
  } finally {
    event.target.value = "";
  }
}

function stopWorkflowPolling() {
  if (state.workflowPollTimer) {
    clearInterval(state.workflowPollTimer);
    state.workflowPollTimer = null;
  }
}

async function handleOverlayFiles(event) {
  const page = getSelectedPage();
  const draft = ensurePageDraft(page);
  const files = Array.from(event.target.files || []);
  if (!page || !files.length) return;
  for (const file of files) {
    try {
      const src = await fileToDataUrl(file);
      draft.overlays.push({
        id: uid(),
        src,
        x: 8,
        y: 8,
        w: 26,
        h: 26,
      });
    } catch (error) {
      setStatus(error.message || "补充图片上传失败。", "error");
    }
  }
  renderArtboard();
  saveState();
  event.target.value = "";
}

function clearCurrentOverlays() {
  const page = getSelectedPage();
  const draft = ensurePageDraft(page);
  if (!draft) return;
  draft.overlays = [];
  renderArtboard();
  saveState();
}

function getCurrentReviseImage() {
  return state.revise.images.find((item) => item.id === state.revise.selectedImageId) || null;
}

function ensureReviseSelection() {
  const current = getCurrentReviseImage();
  if (!current) {
    state.revise.selectedImageId = state.revise.images[0]?.id || "";
  }
}

function renderRevise() {
  ensureReviseSelection();
  const image = getCurrentReviseImage();
  el.reviseImageName.textContent = image?.name || "请先导入底图";
  const index = state.revise.images.findIndex((item) => item.id === state.revise.selectedImageId);
  el.reviseImageCounter.textContent = image ? `${index + 1} / ${state.revise.images.length}` : "0 / 0";
  el.reviseBaseImage.hidden = !image;
  el.reviseEmptyState.hidden = Boolean(image);
  if (image) {
    el.reviseBaseImage.src = image.src;
    drawReviseCanvas();
  } else {
    const ctx = el.reviseCanvas.getContext("2d");
    ctx.clearRect(0, 0, el.reviseCanvas.width, el.reviseCanvas.height);
  }

  el.reviseThumbStrip.innerHTML = state.revise.images.map((item, idx) => `
    <button class="thumb-item ${item.id === state.revise.selectedImageId ? "is-active" : ""}" type="button" data-image-id="${item.id}">
      <img src="${escapeHtml(item.src)}" alt="${escapeHtml(item.name)}" />
      <strong>${escapeHtml(`${idx + 1}. ${item.name}`)}</strong>
    </button>
  `).join("");
  el.reviseThumbStrip.querySelectorAll("[data-image-id]").forEach((button) => {
    button.addEventListener("click", () => {
      state.revise.selectedImageId = button.dataset.imageId;
      renderRevise();
    });
  });

  el.reviseResultStrip.innerHTML = state.revise.results.map((src) => `
    <div class="result-item"><img src="${escapeHtml(src)}" alt="改图结果" /></div>
  `).join("");
}

function fitCanvasToStage() {
  const rect = el.reviseStage.getBoundingClientRect();
  el.reviseCanvas.width = rect.width;
  el.reviseCanvas.height = rect.height;
}

function drawReviseCanvas(tempBox = null) {
  fitCanvasToStage();
  const ctx = el.reviseCanvas.getContext("2d");
  ctx.clearRect(0, 0, el.reviseCanvas.width, el.reviseCanvas.height);
  const image = getCurrentReviseImage();
  if (!image?.naturalWidth || !image?.naturalHeight) return;
  const scaleX = el.reviseCanvas.width / image.naturalWidth;
  const scaleY = el.reviseCanvas.height / image.naturalHeight;
  [...(image.boxes || []), ...(tempBox ? [tempBox] : [])].forEach((box, index) => {
    const [x1, y1, x2, y2] = box;
    ctx.strokeStyle = "#d92d20";
    ctx.fillStyle = "rgba(217, 45, 32, 0.14)";
    ctx.lineWidth = 2;
    ctx.strokeRect(x1 * scaleX, y1 * scaleY, (x2 - x1) * scaleX, (y2 - y1) * scaleY);
    ctx.fillRect(x1 * scaleX, y1 * scaleY, (x2 - x1) * scaleX, (y2 - y1) * scaleY);
    ctx.fillStyle = "#ffffff";
    ctx.font = "600 12px sans-serif";
    ctx.fillText(`区域 ${index + 1}`, x1 * scaleX + 6, y1 * scaleY + 18);
  });
}

function stagePointToNatural(clientX, clientY) {
  const image = getCurrentReviseImage();
  if (!image) return null;
  const rect = el.reviseStage.getBoundingClientRect();
  const x = clamp(((clientX - rect.left) / rect.width) * image.naturalWidth, 0, image.naturalWidth);
  const y = clamp(((clientY - rect.top) / rect.height) * image.naturalHeight, 0, image.naturalHeight);
  return { x, y };
}

function setupReviseCanvasInteractions() {
  const startDrawing = (event) => {
    const image = getCurrentReviseImage();
    if (!image) return;
    const point = stagePointToNatural(event.clientX, event.clientY);
    if (!point) return;
    state.revise.drawing = { start: point };
  };
  const moveDrawing = (event) => {
    const image = getCurrentReviseImage();
    if (!image || !state.revise.drawing?.start) return;
    const current = stagePointToNatural(event.clientX, event.clientY);
    if (!current) return;
    const box = normalizeBox([
      state.revise.drawing.start.x,
      state.revise.drawing.start.y,
      current.x,
      current.y,
    ]);
    drawReviseCanvas(box);
  };
  const endDrawing = (event) => {
    const image = getCurrentReviseImage();
    if (!image || !state.revise.drawing?.start) return;
    const current = stagePointToNatural(event.clientX, event.clientY);
    const box = current ? normalizeBox([
      state.revise.drawing.start.x,
      state.revise.drawing.start.y,
      current.x,
      current.y,
    ]) : null;
    state.revise.drawing = null;
    if (box && image.boxes.length < MAX_REVISE_BOXES) {
      image.boxes.push(box);
    }
    drawReviseCanvas();
  };
  el.reviseCanvas.addEventListener("pointerdown", startDrawing);
  el.reviseCanvas.addEventListener("pointermove", moveDrawing);
  el.reviseCanvas.addEventListener("dblclick", () => {
    const image = getCurrentReviseImage();
    if (!image) return;
    image.boxes = [];
    drawReviseCanvas();
    setStatus("已清空当前图片的框选区域。", "success");
  });
  window.addEventListener("pointerup", endDrawing);
}

function normalizeBox(box) {
  const [ax, ay, bx, by] = box;
  return [
    Math.round(Math.min(ax, bx)),
    Math.round(Math.min(ay, by)),
    Math.round(Math.max(ax, bx)),
    Math.round(Math.max(ay, by)),
  ];
}

async function handleReviseFiles(event) {
  const files = Array.from(event.target.files || []);
  if (!files.length) return;
  for (const file of files) {
    try {
      const src = await fileToDataUrl(file);
      const img = await loadImage(src);
      state.revise.images.push({
        id: uid(),
        name: file.name,
        src,
        naturalWidth: img.naturalWidth,
        naturalHeight: img.naturalHeight,
        boxes: [],
      });
    } catch (error) {
      setStatus(error.message || "导入图片失败。", "error");
    }
  }
  ensureReviseSelection();
  renderRevise();
  event.target.value = "";
}

function stepReviseImage(delta) {
  if (state.revise.images.length <= 1) return;
  const currentIndex = state.revise.images.findIndex((item) => item.id === state.revise.selectedImageId);
  const nextIndex = (Math.max(0, currentIndex) + delta + state.revise.images.length) % state.revise.images.length;
  state.revise.selectedImageId = state.revise.images[nextIndex].id;
  renderRevise();
}

function deleteCurrentReviseImage() {
  if (!state.revise.selectedImageId) return;
  state.revise.images = state.revise.images.filter((item) => item.id !== state.revise.selectedImageId);
  ensureReviseSelection();
  renderRevise();
}

function normalizeDisplayText(text) {
  return String(text || "").trim();
}

function formatJobStats(job) {
  if (!job) return "";
  return [
    `总页数 ${job.totalPages || 0}`,
    `已准备 ${job.preparedPages || 0}`,
    `可生成 ${job.readyToGeneratePages || 0}`,
    `失败 ${job.failedPages || 0}`,
  ].join(" · ");
}

function renderReferenceFiles() {
  if (!state.parsedFiles.length) {
    el.referenceFilesList.innerHTML = `<div class="inline-hint">暂无参考文件</div>`;
    return;
  }
  el.referenceFilesList.innerHTML = state.parsedFiles.map((file) => `
    <div class="file-item">
      <strong>${escapeHtml(file.name)}</strong>
      <div class="file-meta">
        <span class="meta-pill">${escapeHtml(file.category || "unknown")}</span>
        <span class="meta-pill">${escapeHtml(file.parseStatus || "unknown")}</span>
      </div>
      ${file.previewUrl ? `<img class="file-preview-image" src="${escapeHtml(file.previewUrl)}" alt="${escapeHtml(file.name)}" />` : ""}
      ${file.previewText ? `<details class="trace-details"><summary>预览文本</summary><pre>${escapeHtml(file.previewText)}</pre></details>` : ""}
    </div>
  `).join("");
}

async function confirmTheme() {
  if (!state.themeDefinition) return;
  state.themeConfirmed = true;
  if (state.workflowJobId && state.workflowJob) {
    try {
      const data = await apiJson("/api/workflow/theme/apply", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          jobId: state.workflowJobId,
          themeDefinition: state.themeDefinition,
          preferences: state.preferences,
          decorationLevel: state.decorationLevel,
          promptTrace: { themeCore: state.themePromptTrace },
        }),
      });
      state.workflowJob = sanitizeRecoveredWorkflowJob(data.job);
    } catch (error) {
      setStatus(error.message || "应用风格失败。", "error");
      return;
    }
  }
  updateThemeView();
  switchSmartStep(state.workflowJob ? "pages" : "split");
  setStatus(state.workflowJob ? "风格已应用到当前项目，可以逐页确认并生成。" : "风格已确认，继续输入文本并拆分。", "success");
  saveState();
}

function renderOnscreenPreview(value) {
  if (!el.pageOnscreenPreview) return;
  const lines = formatOnscreenPreview(value).split("\n").map((line) => line.trim()).filter(Boolean);
  if (!lines.length) {
    el.pageOnscreenPreview.innerHTML = `<div class="inline-hint">当前页还没有可展示的上屏内容。</div>`;
    return;
  }
  el.pageOnscreenPreview.innerHTML = lines.map((line, index) => {
    if (index === 0) {
      return `<div class="onscreen-preview-title">${escapeHtml(line)}</div>`;
    }
    if (/^(\d+[\.\u3001]|[一二三四五六七八九十]+[、.])/.test(line)) {
      return `<div class="onscreen-preview-bullet">${escapeHtml(line)}</div>`;
    }
    return `<div class="onscreen-preview-line">${escapeHtml(line)}</div>`;
  }).join("");
}

function buildThemePagePlanSummary() {
  const pages = Array.isArray(state.workflowJob?.pages) ? state.workflowJob.pages : [];
  return pages.slice(0, 24).map((page) => {
    const content = formatOnscreenPreview(page.onscreenContentText || page.onscreenContent || page.pageContent || "");
    return `${page.pageNumber}. [${page.pageType || "content"}] ${page.pageTitle || ""}: ${content.slice(0, 180)}`;
  }).join("\n");
}

async function generateTheme() {
  state.themeName = el.themeName.value.trim();
  state.decorationLevel = el.themeDecorationLevel.value;
  state.preferences = getCurrentPreferences();
  await ensureServerConfigReady();
  if (!hasDashScopeApiKey()) {
    setStatus("请先填写 DashScope / Qwen API Key。", "error");
    switchTab("settings");
    return;
  }
  const signal = startCancelableAction("theme", el.generateThemeBtn, el.cancelThemeBtn, "生成中...");
  setStatus("正在根据内容匹配风格主题...", "running");
  try {
    const data = await apiJson("/api/workflow/theme", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      signal,
      body: JSON.stringify({
        apiKey: state.settings.apiKey,
        region: state.settings.region,
        themeName: state.themeName || "AI 自动匹配成熟风格",
        decorationLevel: state.decorationLevel,
        preferences: state.preferences,
        referenceFiles: state.parsedFiles,
        workflowJobId: state.workflowJobId,
        contentContext: state.workflowContent,
        pagePlanSummary: buildThemePagePlanSummary(),
      }),
    });
    state.themeDefinition = data.themeDefinition;
    state.themePromptTrace = data.promptTrace || null;
    state.themeConfirmed = false;
    updateThemeView();
    setStatus("风格模板已生成，确认后即可使用。", "success");
    saveState();
  } catch (error) {
    if (isAbortError(error)) return;
    setStatus(error.message || "生成风格失败。", "error");
  } finally {
    finishCancelableAction("theme");
  }
}

async function sendRevise() {
  const image = getCurrentReviseImage();
  state.revise.prompt = el.revisePrompt.value.trim();
  await ensureServerConfigReady();
  if (!hasDashScopeApiKey()) {
    setStatus("改图需要 DashScope / Qwen API Key。", "error");
    switchTab("settings");
    return;
  }
  if (!image) {
    setStatus("请先导入一张底图。", "error");
    return;
  }
  if (!state.revise.prompt) {
    setStatus("请先输入改图提示词。", "error");
    return;
  }

  const payload = {
    model: EDIT_MODEL,
    input: {
      messages: [
        {
          role: "user",
          content: [
            { text: state.revise.prompt },
            { image: image.src },
          ],
        },
      ],
    },
    parameters: {
      size: getAspectMeta().outputSize,
      n: 1,
    },
  };
  if (state.settings.seed) payload.parameters.seed = Number(state.settings.seed);
  if (Array.isArray(image.boxes) && image.boxes.length) {
    payload.parameters.bbox_list = [image.boxes];
  }

  const signal = startCancelableAction("revise", el.sendReviseBtn, el.cancelReviseBtn, "改图中...");
  setStatus("正在调用 Wan 改图...", "running");
  try {
    const response = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      signal,
      body: JSON.stringify({
        workflowRouteVersion: 2,
        apiKey: state.settings.apiKey,
        googleApiKey: state.settings.googleApiKey,
        region: state.settings.region,
        slideAspect: state.settings.slideAspect,
        payload,
      }),
    });
    const data = await response.json();
    if (!response.ok || data.code) throw new Error(data.message || "改图失败。");
    const results = ((data.output?.choices || [])[0]?.message?.content || [])
      .filter((item) => item.type === "image" && item.image)
      .map((item) => item.image);
    if (!results.length) throw new Error("接口返回成功，但没有图片结果。");
    state.revise.results = results;
    const current = getCurrentReviseImage();
    if (current && results[0]) {
      current.src = results[0];
      current.boxes = [];
      current.naturalWidth = 0;
      current.naturalHeight = 0;
      el.reviseBaseImage.onload = () => {
        current.naturalWidth = el.reviseBaseImage.naturalWidth;
        current.naturalHeight = el.reviseBaseImage.naturalHeight;
        drawReviseCanvas();
      };
    }
    renderRevise();
    setStatus("改图完成。", "success");
  } catch (error) {
    if (isAbortError(error)) return;
    setStatus(error.message || "改图失败。", "error");
  } finally {
    finishCancelableAction("revise");
  }
}

function bindEvents() {
  document.querySelectorAll(".sidebar-tab").forEach((button) => {
    button.addEventListener("click", () => switchTab(button.dataset.tab));
  });
  document.querySelectorAll(".ribbon-step").forEach((button) => {
    button.addEventListener("click", () => {
      const next = button.dataset.step;
      if (next === "pages" && !state.workflowJob) return;
      switchSmartStep(next);
    });
  });

  el.workspaceZoomRange.addEventListener("input", () => {
    applyWorkspaceZoom(el.workspaceZoomRange.value);
    saveState();
  });
  ["prefStyleMode", "prefLayoutVariety", "prefDetailLevel", "prefVisualDensity", "prefCompositionFocus", "prefDataNarrative", "prefPageMood"].forEach((id) => {
    el[id].addEventListener("change", () => {
      renderPreferenceSummary();
      saveState();
    });
  });
  el.themeName.addEventListener("input", () => {
    state.themeName = el.themeName.value;
    saveState();
  });
  el.themeDecorationLevel.addEventListener("change", () => {
    state.decorationLevel = el.themeDecorationLevel.value;
    saveState();
  });
  el.generateThemeBtn.addEventListener("click", generateTheme);
  el.cancelThemeBtn.addEventListener("click", () => cancelAction("theme"));
  el.confirmThemeBtn.addEventListener("click", confirmTheme);
  el.goSplitBtn.addEventListener("click", () => switchSmartStep(state.workflowJob ? "pages" : "split"));
  el.backToThemeBtn.addEventListener("click", () => switchSmartStep("theme"));
  el.backToSplitBtn.addEventListener("click", () => switchSmartStep("split"));
  el.pickReferenceFilesBtn.addEventListener("click", () => el.referenceFilesInput.click());
  el.referenceFilesInput.addEventListener("change", handleReferenceFiles);
  el.runSplitBtn.addEventListener("click", runSplit);
  el.cancelSplitBtn.addEventListener("click", () => cancelAction("split"));
  el.workflowContent.addEventListener("input", () => {
    state.workflowContent = el.workflowContent.value;
    saveState();
  });
  el.workflowPageCount.addEventListener("change", () => {
    state.workflowPageCount = clamp(Number(el.workflowPageCount.value || 8), 2, 120);
    saveState();
  });
  el.aiProcessingMode.addEventListener("change", () => {
    state.aiProcessingMode = el.aiProcessingMode.value || "balanced";
    saveState();
  });
  el.workflowEnableExpansion?.addEventListener("change", () => {
    state.workflowEnableExpansion = Boolean(el.workflowEnableExpansion.checked);
    syncSplitExpansionControls();
    saveState();
  });
  el.splitTemplateInput?.addEventListener("input", () => {
    state.splitTemplateText = el.splitTemplateInput.value;
    saveState();
  });
  el.pageOnscreenEditor.addEventListener("input", () => {
    const page = getSelectedPage();
    const draft = ensurePageDraft(page);
    if (!draft) return;
    draft.onscreenContent = el.pageOnscreenEditor.value;
    renderOnscreenPreview(draft.onscreenContent);
    saveState();
  });
  el.pageExtraPrompt.addEventListener("input", () => {
    const page = getSelectedPage();
    const draft = ensurePageDraft(page);
    if (!draft) return;
    draft.extraPrompt = el.pageExtraPrompt.value;
    saveState();
  });
  el.repreparePageBtn.addEventListener("click", reprepareCurrentPage);
  el.aiRepolishPageBtn.addEventListener("click", aiRepolishCurrentPage);
  el.cancelRepreparePageBtn.addEventListener("click", () => cancelAction("repolish") || cancelAction("reprepare"));
  el.batchGenerateReadyBtn.addEventListener("click", batchGenerateReadyPages);
  el.cancelBatchGenerateBtn.addEventListener("click", () => cancelAction("batchGenerate"));
  el.uploadOverlayBtn.addEventListener("click", () => el.overlayFileInput.click());
  el.overlayFileInput.addEventListener("change", handleOverlayFiles);
  el.clearOverlayBtn.addEventListener("click", clearCurrentOverlays);
  el.generateCurrentPageBtn.addEventListener("click", generateCurrentPage);
  el.copyPagePromptBtn?.addEventListener("click", copyCurrentPagePrompt);
  el.viewCurrentPageLargeBtn?.addEventListener("click", openCurrentPageLargeImage);
  el.exportWorkflowPptBtn?.addEventListener("click", exportWorkflowPpt);
  el.cancelGenerateCurrentPageBtn.addEventListener("click", () => {
    const page = getSelectedPage();
    if (!page) return;
    cancelAction(getPageGenerateRequestKey(page.id), "已取消当前页生成。");
  });
  el.restoreHistoryProjectBtn?.addEventListener("click", restoreHistoryProject);
  el.closePageImageModalBtn?.addEventListener("click", closeCurrentPageLargeImage);
  document.querySelectorAll("[data-close-page-image-modal]").forEach((node) => {
    node.addEventListener("click", closeCurrentPageLargeImage);
  });

  el.reviseImportBtn.addEventListener("click", () => el.reviseFileInput.click());
  el.reviseFileInput.addEventListener("change", handleReviseFiles);
  el.revisePrevBtn.addEventListener("click", () => stepReviseImage(-1));
  el.reviseNextBtn.addEventListener("click", () => stepReviseImage(1));
  el.reviseDeleteBtn.addEventListener("click", deleteCurrentReviseImage);
  el.revisePrompt.addEventListener("input", () => {
    state.revise.prompt = el.revisePrompt.value;
  });
  el.sendReviseBtn.addEventListener("click", sendRevise);
  el.cancelReviseBtn.addEventListener("click", () => cancelAction("revise"));

  el.apiKey.addEventListener("input", () => {
    state.settings.apiKey = el.apiKey.value.trim();
    if (el.quickApiKey) el.quickApiKey.value = state.settings.apiKey;
    saveState();
  });
  el.googleApiKey.addEventListener("input", () => {
    state.settings.googleApiKey = el.googleApiKey.value.trim();
    if (el.quickGoogleApiKey) el.quickGoogleApiKey.value = state.settings.googleApiKey;
    saveState();
  });
  el.quickApiKey?.addEventListener("input", () => {
    state.settings.apiKey = el.quickApiKey.value.trim();
    if (el.apiKey) el.apiKey.value = state.settings.apiKey;
    saveState();
  });
  el.quickGoogleApiKey?.addEventListener("input", () => {
    state.settings.googleApiKey = el.quickGoogleApiKey.value.trim();
    if (el.googleApiKey) el.googleApiKey.value = state.settings.googleApiKey;
    saveState();
  });
  el.workflowImageModel.addEventListener("change", () => {
    state.settings.workflowImageModel = el.workflowImageModel.value || PPT_MODEL;
    syncGeminiSearchControls();
    saveState();
  });
  el.grsaiHost?.addEventListener("change", () => {
    state.settings.grsaiHost = el.grsaiHost.value || "domestic";
    if (el.quickGrsaiHost) el.quickGrsaiHost.value = state.settings.grsaiHost;
    saveState();
  });
  el.quickGrsaiHost?.addEventListener("change", () => {
    state.settings.grsaiHost = el.quickGrsaiHost.value || "domestic";
    if (el.grsaiHost) el.grsaiHost.value = state.settings.grsaiHost;
    saveState();
  });
  el.enableGeminiGoogleSearch?.addEventListener("change", () => {
    state.settings.enableGeminiGoogleSearch = Boolean(el.enableGeminiGoogleSearch.checked && usingGeminiWorkflowModel());
    saveState();
  });
  el.region.addEventListener("change", () => { state.settings.region = el.region.value; saveState(); });
  el.slideAspect.addEventListener("change", () => { state.settings.slideAspect = el.slideAspect.value; renderArtboard(); saveState(); });
  el.outputSize.addEventListener("change", () => { state.settings.outputSize = el.outputSize.value; saveState(); });
  el.seed.addEventListener("input", () => { state.settings.seed = el.seed.value.trim(); saveState(); });
  el.testApiKeyBtn.addEventListener("click", testApiKeys);
  el.quickTestApiKeyBtn?.addEventListener("click", testApiKeys);
  el.cancelTestApiKeyBtn.addEventListener("click", () => cancelAction("testApi"));

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && !el.pageImageModal?.hidden) {
      closeCurrentPageLargeImage();
    }
  });

  window.addEventListener("resize", () => {
    renderArtboard();
    renderRevise();
  });
}

function renderPageResults() {
  if (!el.pageResultStrip) return;
  const page = getSelectedPage();
  if (!page) {
    el.pageResultStrip.innerHTML = "";
    return;
  }

  const historyImages = Array.isArray(page.resultImages)
    ? Array.from(new Set(page.resultImages.filter(Boolean)))
    : [];
  if (page.baseImage && !historyImages.includes(page.baseImage)) {
    historyImages.unshift(page.baseImage);
  }
  const isGenerating = isPageGenerating(page.id);
  const statusText = isGenerating
    ? "当前请求已提交，正在生成中..."
    : page.generationStatus === "error"
      ? (page.generationError || "最近一次生成失败。")
      : historyImages.length
        ? `已保存 ${historyImages.length} 个历史版本`
        : "还没有历史生图版本。";

  el.pageResultStrip.innerHTML = `
    <div class="page-result-header">
      <strong>历史生图版本</strong>
      <span class="inline-hint">${escapeHtml(statusText)}</span>
    </div>
    ${historyImages.length ? `
      <div class="page-result-thumbs">
        ${historyImages.map((src, index) => `
          <button
            class="result-item result-thumb ${src === page.baseImage ? "is-active" : ""}"
            type="button"
            data-result-image-src="${escapeHtml(src)}"
            aria-label="${escapeHtml(`切换到第 ${index + 1} 个历史版本`)}"
            title="${escapeHtml(`第 ${index + 1} 个版本`)}"
          >
            <img src="${escapeHtml(src)}" alt="${escapeHtml(page.pageTitle || `第${page.pageNumber}页`)}" />
            <span>版本 ${index + 1}</span>
          </button>
        `).join("")}
      </div>
    ` : `<div class="inline-summary">${escapeHtml(statusText)}</div>`}
  `;

  el.pageResultStrip.querySelectorAll("[data-result-image-src]").forEach((node) => {
    node.addEventListener("click", () => {
      page.baseImage = node.dataset.resultImageSrc || page.baseImage;
      saveState();
      renderPagesWorkbench();
    });
  });
}

function formatOnscreenPreview(value) {
  const source = String(value || "")
    .replace(/\r/g, "")
    .replace(/```[\s\S]*?```/g, "")
    .trim();
  if (!source) return "";

  const lines = source
    .split(/\r?\n+/)
    .map((line) => line.trim())
    .filter(Boolean)
    .filter((line) => !/^(?:视觉元素|视觉建议|画面建议|设计说明|版式说明|构图说明|画面说明|视觉)\s*[:：]/i.test(line))
    .map((line) => line
      .replace(/^\s*(?:blocks|items|points|entries|sections)\s*[:：]\s*/i, "")
      .replace(/^\s*(?:title|subtitle|metaInfo|metainfo|abstract|summary|body|content|text|heading|detail|label|metric|value|note|type|highlight|dataPoints?)\s*[:：]\s*/i, "")
      .replace(/\s+/g, " ")
      .trim())
    .filter(Boolean);

  return lines.filter((line, index) => line !== lines[index - 1]).join("\n");
}

async function runSplit() {
  state.workflowContent = el.workflowContent.value.trim();
  state.workflowPageCount = clamp(Number(el.workflowPageCount.value || 8), 2, 120);
  state.aiProcessingMode = el.aiProcessingMode.value || "balanced";
  state.workflowEnableExpansion = Boolean(el.workflowEnableExpansion?.checked);
  state.workflowTargetChars = clamp(Number(el.workflowTargetChars?.value || 0), 0, 300);
  state.workflowMaxChars = clamp(Number(el.workflowMaxChars?.value || 200), 0, 400);
  if (state.workflowEnableExpansion && state.workflowTargetChars && state.workflowMaxChars && state.workflowTargetChars > state.workflowMaxChars) {
    state.workflowTargetChars = state.workflowMaxChars;
    if (el.workflowTargetChars) el.workflowTargetChars.value = String(state.workflowTargetChars);
  }
  state.splitTemplateText = "";
  await ensureServerConfigReady();
  if (!hasDashScopeApiKey()) {
    setStatus("请先填写 DashScope / Qwen API Key。", "error");
    switchTab("settings");
    return;
  }
  if (!state.workflowContent) {
    setStatus("请先输入主文本。", "error");
    return;
  }

  const signal = startCancelableAction("split", el.runSplitBtn, el.cancelSplitBtn, "拆分中...");
  state.workflowJob = {
    id: state.workflowJobId || "",
    status: "running",
    totalPages: state.workflowPageCount,
    preparedPages: 0,
    readyToGeneratePages: 0,
    failedPages: 0,
    pages: [],
    statusText: "正在拆分内容并准备逐页结果...",
  };
  state.selectedPageId = "";
  renderPagesWorkbench();
  setStatus("正在拆分内容并准备逐页结果...", "running");
  try {
    const data = await apiJson("/api/workflow/split", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      signal,
      body: JSON.stringify({
        apiKey: state.settings.apiKey,
        region: state.settings.region,
        content: state.workflowContent,
        pageCount: state.workflowPageCount,
        splitTemplate: "",
        aiProcessingMode: state.aiProcessingMode,
        enableExpansion: state.workflowEnableExpansion,
        targetChars: state.workflowEnableExpansion ? state.workflowTargetChars : 0,
        maxChars: state.workflowMaxChars,
        referenceFiles: state.parsedFiles,
        themeDefinition: state.themeDefinition,
        preferences: state.preferences,
        decorationLevel: state.decorationLevel,
      }),
    });
    state.workflowJobId = data.jobId;
    state.workflowJob = sanitizeRecoveredWorkflowJob(data.job);
    state.selectedPageId = data.job?.pages?.[0]?.id || "";
    ensureSelectedPage();
    startWorkflowPolling();
    renderPagesWorkbench();
    state.themeConfirmed = false;
    switchSmartStep("theme");
    setStatus("内容已拆分，下一步根据页面结构匹配风格。", "success");
    saveState();
  } catch (error) {
    if (isAbortError(error)) {
      clearWorkflowSession({ toSplit: true });
      return;
    }
    clearWorkflowSession({ toSplit: true });
    setStatus(error.message || "拆分失败。", "error");
  } finally {
    finishCancelableAction("split");
  }
}

function startWorkflowPolling() {
  stopWorkflowPolling();
  if (!state.workflowJobId) return;
  state.workflowPollTimer = setInterval(async () => {
    try {
      const data = await apiJson(`/api/workflow/jobs/${encodeURIComponent(state.workflowJobId)}`);
      state.workflowJob = sanitizeRecoveredWorkflowJob(data.job);
      ensureSelectedPage();
      renderPagesWorkbench();
      renderHistoryProjects();
      if (el.workflowRibbonMeta) {
        el.workflowRibbonMeta.textContent = `${data.job.readyToGeneratePages || 0} 页可直接生成，当前已准备 ${data.job.preparedPages || 0}/${data.job.totalPages || 0} 页。`;
      }
      if (data.job.status === "ready") {
        stopWorkflowPolling();
        setStatus(data.job.statusText || "页面准备完成。", "success");
      }
      saveState();
    } catch (error) {
      stopWorkflowPolling();
      if (isMissingWorkflowJobError(error)) {
        clearWorkflowSession({ toSplit: true });
        setStatus("之前的拆分任务已失效，请重新拆分。", "error");
        return;
      }
      setStatus(error.message || "读取任务进度失败。", "error");
    }
  }, 2200);
}

async function syncWorkflowJobOnce() {
  if (!state.workflowJobId) return;
  try {
    const data = await apiJson(`/api/workflow/jobs/${encodeURIComponent(state.workflowJobId)}`);
    state.workflowJob = sanitizeRecoveredWorkflowJob(data.job);
    ensureSelectedPage();
    renderPagesWorkbench();
    renderHistoryProjects();
    saveState();
  } catch (error) {
    if (isMissingWorkflowJobError(error)) {
      clearWorkflowSession({ toSplit: true });
      setStatus("之前的拆分任务已失效，请重新拆分。", "error");
    }
  }
}

function renderPageList() {
  const job = state.workflowJob;
  if (!job?.pages?.length) {
    el.workflowStats.textContent = "";
    el.workflowDiagnostics.textContent = "";
    el.workflowPromptTrace.textContent = "";
    el.workflowPageList.innerHTML = "";
    return;
  }
  ensureSelectedPage();
  if (el.workflowSummary) el.workflowSummary.textContent = "";
  el.workflowStats.textContent = formatJobStats(job);
  el.workflowDiagnostics.textContent = normalizeDisplayText(job.splitDiagnostics);
  el.workflowPromptTrace.textContent = stringifyTrace(job.promptTrace);
  el.workflowPageList.innerHTML = job.pages.map((page) => {
    const typeMeta = getPageTypeMeta(page.pageType);
    const pageActive = isPageGenerating(page.id);
    const serverBusy = ["preparing", "running"].includes(page.generationStatus);
    const status = page.generated
      ? "已生成"
      : pageActive || serverBusy
        ? (page.generationStatus === "preparing" ? "准备提示词" : "生成中")
        : page.prepareDone ? "可生成" : "处理中";
    const statusClass = page.generated ? "generated" : (pageActive || serverBusy || page.prepareDone) ? "ready" : "idle";
    const riskClass = page.riskLevel === "high" ? "high" : page.riskLevel === "medium" ? "medium" : "";
    return `
      <div class="page-item ${page.id === state.selectedPageId ? "is-active" : ""}" data-page-id="${page.id}">
        <div class="page-title-row">
          <strong>第${page.pageNumber} 页 · ${escapeHtml(page.pageTitle || "未命名")}</strong>
          <span class="status-pill ${statusClass}">${status}</span>
        </div>
        <div class="page-meta">
          <span class="meta-pill page-type-pill page-type-${escapeHtml(String(page.pageType || "content").toLowerCase())}">${escapeHtml(typeMeta.label)}</span>
          ${riskClass ? `<span class="risk-pill ${riskClass}">排版风险</span>` : ""}
        </div>
      </div>
    `;
  }).join("");

  el.workflowPageList.querySelectorAll("[data-page-id]").forEach((item) => {
    item.addEventListener("click", () => {
      state.selectedPageId = item.dataset.pageId;
      renderPagesWorkbench();
      saveState();
    });
  });
}

async function batchGenerateReadyPages() {
  const job = state.workflowJob;
  if (!job?.pages?.length) return;

  await ensureServerConfigReady();
  const selectedImageModel = getCurrentWorkflowImageModel();
  if (usingHostedWorkflowModel() && !hasHostedImageApiKey()) {
    setStatus("请先填写生图 API Key。", "error");
    switchTab("settings");
    return;
  }
  if (!usingHostedWorkflowModel() && !hasDashScopeApiKey()) {
    setStatus("请先填写 DashScope / Qwen API Key。", "error");
    switchTab("settings");
    return;
  }

  const candidates = job.pages.filter((page) => page.readyToGenerate && !page.generated);
  if (!candidates.length) {
    setStatus("还没有可直接批量生成的页面。", "error");
    return;
  }

  const signal = startCancelableAction("batchGenerate", el.batchGenerateReadyBtn, el.cancelBatchGenerateBtn, "批量生成中...");
  try {
    for (const page of candidates) {
      const draft = ensurePageDraft(page);
      setStatus(`正在批量生成第${page.pageNumber} 页...`, "running");
      const data = await apiJson("/api/workflow/page/generate-v2", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        signal,
        body: JSON.stringify({
          apiKey: state.settings.apiKey,
          googleApiKey: state.settings.googleApiKey,
          grsaiHost: state.settings.grsaiHost,
          region: state.settings.region,
          imageModel: selectedImageModel,
          jobId: state.workflowJobId,
          pageId: page.id,
          slideAspect: state.settings.slideAspect,
          size: getWorkflowGenerationSize(),
          seed: state.settings.seed,
          enableGeminiGoogleSearch: Boolean(state.settings.enableGeminiGoogleSearch && usingGeminiWorkflowModel()),
          extraPrompt: draft.extraPrompt || "",
          onscreenContent: formatOnscreenPreview(draft.onscreenContent || page.onscreenContentText || page.onscreenContent || ""),
          canvasImage: "",
        }),
      });
      if (!data.page?.generated) {
        throw new Error(data.page?.generationError || `第${page.pageNumber} 页没有拿到图片结果。`);
      }
      state.workflowJob.pages = state.workflowJob.pages.map((item) => item.id === data.page.id ? data.page : item);
      sanitizeRecoveredWorkflowJob(state.workflowJob);
      ensureSelectedPage();
      saveState();
    }
    setStatus("所有可生成页面已完成。", "success");
  } catch (error) {
    if (isAbortError(error)) return;
    if (isMissingWorkflowJobError(error)) {
      clearWorkflowSession({ toSplit: true });
      setStatus("之前的拆分任务已失效，请重新拆分。", "error");
      return;
    }
    setStatus(error.message || "批量生成失败。", "error");
  } finally {
    finishCancelableAction("batchGenerate");
    renderPagesWorkbench();
    syncCurrentPageGenerateUi();
  }
}

async function testApiKeys() {
  state.settings.apiKey = el.apiKey.value.trim();
  state.settings.googleApiKey = el.googleApiKey.value.trim();
  state.settings.workflowImageModel = el.workflowImageModel.value || PPT_MODEL;
  state.settings.grsaiHost = el.grsaiHost?.value || "domestic";
  state.settings.region = el.region.value;
  state.settings.slideAspect = el.slideAspect.value;
  state.settings.outputSize = el.outputSize.value;
  state.settings.seed = el.seed.value.trim();
  await ensureServerConfigReady();
  if (!hasDashScopeApiKey() && !hasHostedImageApiKey()) {
    setStatus("请先填写至少一个可用的 API Key。", "error");
    return;
  }
  const signal = startCancelableAction("testApi", el.testApiKeyBtn, el.cancelTestApiKeyBtn, "测试中...");
  setStatus("正在测试 Key...", "running");
  try {
    const tasks = [];
    const selectedWorkflowModel = getCurrentWorkflowImageModel();
    if (hasHostedImageApiKey() && usingHostedWorkflowModel()) {
      tasks.push(fetch("/api/test-image-key", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        signal,
        body: JSON.stringify({
          apiKey: state.settings.apiKey,
          googleApiKey: state.settings.googleApiKey,
          grsaiHost: state.settings.grsaiHost,
          region: state.settings.region,
          model: selectedWorkflowModel,
        }),
      }).then(async (response) => ({ ok: response.ok, data: await response.json() })));
    }
    if (hasDashScopeApiKey()) {
      tasks.push(fetch("/api/test-image-key", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        signal,
        body: JSON.stringify({
          apiKey: state.settings.apiKey,
          googleApiKey: state.settings.googleApiKey,
          region: state.settings.region,
          model: EDIT_MODEL,
        }),
      }).then(async (response) => ({ ok: response.ok, data: await response.json() })));
    }
    const results = await Promise.all(tasks);
    const failures = results.filter((item) => !item.ok || item.data?.code);
    if (failures.length) {
      throw new Error(failures.map((item) => item.data?.message || "测试失败").join("；"));
    }
    setStatus("Key 测试通过。", "success");
    saveState();
  } catch (error) {
    if (isAbortError(error)) return;
    setStatus(error.message || "测试失败。", "error");
  } finally {
    finishCancelableAction("testApi");
  }
}

function splitOnscreenContentForEditor(value, fallbackTitle = "") {
  const cleaned = formatOnscreenPreview(value);
  const normalizedFallback = normalizeDisplayText(fallbackTitle || "").split("\n")[0].trim();
  let title = normalizedFallback;
  let body = cleaned;

  if (!cleaned) {
    return { title, body: "" };
  }

  const lines = cleaned.split("\n").map((line) => line.trim()).filter(Boolean);
  if (!title && lines.length) {
    const candidate = lines[0];
    if (candidate.length <= 32 && !/[\u3002\uFF01\uFF1F\uFF1B]$/.test(candidate)) {
      title = candidate;
      lines.shift();
    }
  }

  body = lines.join("\n");
  if (title) {
    const escapedTitle = title.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    body = body
      .replace(new RegExp(`^${escapedTitle}\\s*[:\\uFF1A\\u2014\\-]\\s*`), "")
      .replace(new RegExp(`^${escapedTitle}\\s*\\n+`), "")
      .trim();
  }

  if (!body && cleaned && cleaned !== title) {
    body = cleaned;
  }

  return { title, body };
}

function composeOnscreenContentFromEditors(titleValue, bodyValue) {
  const title = normalizeDisplayText(titleValue || "").split("\n")[0].trim();
  const body = formatOnscreenPreview(bodyValue || "");
  if (title && body) {
    if (
      body === title ||
      body.startsWith(`${title}\n`) ||
      body.startsWith(`${title}\uFF1A`) ||
      body.startsWith(`${title}:`)
    ) {
      return body;
    }
    return `${title}\n${body}`.trim();
  }
  return title || body;
}

function updateCurrentPageDraftFromEditors() {
  const page = getSelectedPage();
  const draft = ensurePageDraft(page);
  if (!draft) return "";
  const title = el.pageOnscreenTitleEditor?.value.trim() || "";
  const body = el.pageOnscreenBodyEditor?.value || "";
  draft.onscreenTitle = title;
  draft.onscreenBody = body;
  draft.onscreenContent = composeOnscreenContentFromEditors(title, body);
  return draft.onscreenContent;
}

function attachEnhancedPageEditorEvents() {
  el.pageOnscreenBodyEditor = el.pageOnscreenEditor;
  if (el.pageOnscreenTitleEditor && !el.pageOnscreenTitleEditor.dataset.boundEnhanced) {
    el.pageOnscreenTitleEditor.addEventListener("input", () => {
      updateCurrentPageDraftFromEditors();
      saveState();
    });
    el.pageOnscreenTitleEditor.dataset.boundEnhanced = "true";
  }
  if (el.pageOnscreenBodyEditor && !el.pageOnscreenBodyEditor.dataset.boundEnhanced) {
    el.pageOnscreenBodyEditor.addEventListener("input", () => {
      updateCurrentPageDraftFromEditors();
      saveState();
    });
    el.pageOnscreenBodyEditor.dataset.boundEnhanced = "true";
  }
}

function syncCurrentPageGenerateUi() {
  if (!el.generateCurrentPageBtn || !el.cancelGenerateCurrentPageBtn) return;
  const page = getSelectedPage();
  const pageActive = page ? isPageGenerating(page.id) : false;
  if (pageActive) {
    el.generateCurrentPageBtn.disabled = true;
    el.generateCurrentPageBtn.textContent = "\u751f\u6210\u4e2d...";
  } else {
    el.generateCurrentPageBtn.disabled = !page;
    el.generateCurrentPageBtn.textContent = page?.generated ? "\u91cd\u65b0\u751f\u6210\u8be5\u9875" : "\u751f\u6210\u8be5\u9875";
    delete el.generateCurrentPageBtn.dataset.idleText;
  }
  el.cancelGenerateCurrentPageBtn.hidden = !pageActive;
  el.cancelGenerateCurrentPageBtn.disabled = !pageActive;
  if (el.viewCurrentPageLargeBtn) {
    el.viewCurrentPageLargeBtn.disabled = !page?.baseImage || pageActive;
  }
  if (el.copyPagePromptBtn) {
    el.copyPagePromptBtn.disabled = !page || pageActive;
  }
  if (el.exportWorkflowPptBtn) {
    const hasWorkflowPages = Array.isArray(state.workflowJob?.pages) && state.workflowJob.pages.length > 0;
    if (el.exportWorkflowPptBtn.textContent !== "导出中...") {
      el.exportWorkflowPptBtn.disabled = !hasWorkflowPages;
    }
  }
}

function upgradeSmartUiLayout() {
  const splitStage = document.querySelector('[data-step-panel="split"]');
  const splitControlGrid = splitStage?.querySelector(".split-control-grid");
  if (splitControlGrid && !document.getElementById("workflowEnableExpansion")) {
    const field = document.createElement("label");
    field.className = "field split-toggle-field";
    field.innerHTML = `
      <span>\u5185\u5bb9\u6269\u5c55</span>
      <div class="toggle-inline">
        <input id="workflowEnableExpansion" type="checkbox" />
        <span>\u9700\u8981\u6709\u4f9d\u636e\u6269\u5199</span>
      </div>
    `;
    splitControlGrid.appendChild(field);
    el.workflowEnableExpansion = field.querySelector("#workflowEnableExpansion");
  }
  if (splitControlGrid && !document.getElementById("workflowTargetChars")) {
    const field = document.createElement("label");
    field.className = "field";
    field.innerHTML = `
      <span>\u6bcf\u9875\u76ee\u6807\u5b57\u6570</span>
      <input id="workflowTargetChars" type="number" min="0" max="300" placeholder="\u52fe\u9009\u6269\u5199\u540e\u542f\u7528" />
    `;
    splitControlGrid.appendChild(field);
    el.workflowTargetChars = field.querySelector("#workflowTargetChars");
    el.workflowTargetChars.value = state.workflowTargetChars ? String(state.workflowTargetChars) : "";
  }
  if (splitControlGrid && !document.getElementById("workflowMaxChars")) {
    const field = document.createElement("label");
    field.className = "field";
    field.innerHTML = `
      <span>\u6bcf\u9875\u6700\u5927\u5b57\u6570</span>
      <input id="workflowMaxChars" type="number" min="0" max="400" placeholder="\u9ed8\u8ba4 200\uff0c0 \u8868\u793a\u4e0d\u538b\u7f29" />
    `;
    splitControlGrid.appendChild(field);
    el.workflowMaxChars = field.querySelector("#workflowMaxChars");
    el.workflowMaxChars.value = state.workflowMaxChars ? String(state.workflowMaxChars) : "";
  }
  const splitTemplateCard = splitStage?.querySelector(".split-template-card");
  const splitFooter = splitStage?.querySelector(".stage-footer");
  splitTemplateCard?.remove();
  syncSplitExpansionControls();
  if (splitFooter && !splitFooter.querySelector("#splitNextHintBtn")) {
    const placeholder = document.createElement("button");
    placeholder.type = "button";
    placeholder.className = "btn ghost";
    placeholder.id = "splitNextHintBtn";
    placeholder.disabled = true;
    placeholder.textContent = "\u4e0b\u4e00\u6b65";
    splitFooter.appendChild(placeholder);
  }

  const pagesStage = document.querySelector('[data-step-panel="pages"]');
  const onscreenCard = pagesStage?.querySelector(".onscreen-card");
  if (onscreenCard) {
    document.getElementById("pageVisualElementsBlock")?.remove();
    el.pageVisualElementsBlock = null;
    el.pageVisualElementsDisplay = null;
    const preview = onscreenCard.querySelector("#pageOnscreenPreview");
    if (preview) {
      preview.remove();
      el.pageOnscreenPreview = null;
    }
    const bodyField = onscreenCard.querySelector('label[for="pageOnscreenEditor"], label.field.field-stack.grow, label.field.field-stack.onscreen-body-field') || onscreenCard.querySelector("label.field.field-stack");
    if (bodyField && !document.getElementById("pageOnscreenTitleEditor")) {
      const titleField = document.createElement("label");
      titleField.className = "field onscreen-title-field";
      titleField.id = "pageOnscreenTitleField";
      titleField.innerHTML = `
        <span>\u6807\u9898</span>
        <textarea id="pageOnscreenTitleEditor" rows="2" placeholder="\u8fd9\u4e00\u9875\u7684\u6807\u9898"></textarea>
      `;
      bodyField.insertAdjacentElement("beforebegin", titleField);
    }
    const titleField = document.getElementById("pageOnscreenTitleEditor");
    if (titleField) {
      el.pageOnscreenTitleEditor = titleField;
    }
    if (bodyField) {
      bodyField.classList.remove("grow");
      bodyField.classList.add("onscreen-body-field");
      const label = bodyField.querySelector("span");
      if (label) label.textContent = "\u6b63\u6587";
    }
    el.pageOnscreenBodyEditor = el.pageOnscreenEditor;
    if (el.pageOnscreenBodyEditor) {
      el.pageOnscreenBodyEditor.rows = 14;
      el.pageOnscreenBodyEditor.placeholder = "\u8fd9\u4e00\u9875\u7684\u4e3b\u8981\u4e0a\u5c4f\u6587\u5b57";
    }
  }

  const artboardToolbar = pagesStage?.querySelector(".artboard-toolbar");
  if (artboardToolbar && !document.getElementById("pageDrawToolbar")) {
    const drawToolbar = document.createElement("div");
    drawToolbar.className = "page-draw-toolbar";
    drawToolbar.id = "pageDrawToolbar";
    drawToolbar.innerHTML = `
      <button class="btn ghost tool-btn" type="button" id="pageDrawPenBtn">\u753b\u7b14</button>
      <button class="btn ghost tool-btn" type="button" id="pageDrawRectBtn">\u77e9\u5f62</button>
      <label class="tool-color" for="pageDrawColorInput">
        <span>\u989c\u8272</span>
        <input id="pageDrawColorInput" type="color" value="#22d3ee" />
      </label>
      <button class="btn ghost" type="button" id="clearPageDrawingBtn">\u6e05\u7a7a\u7ed8\u5236</button>
    `;
    artboardToolbar.appendChild(drawToolbar);
    el.pageDrawToolbar = drawToolbar;
    el.pageDrawPenBtn = drawToolbar.querySelector("#pageDrawPenBtn");
    el.pageDrawRectBtn = drawToolbar.querySelector("#pageDrawRectBtn");
    el.pageDrawColorInput = drawToolbar.querySelector("#pageDrawColorInput");
    el.clearPageDrawingBtn = drawToolbar.querySelector("#clearPageDrawingBtn");
  }
  const slideStage = pagesStage?.querySelector("#slideStage");
  if (slideStage && !document.getElementById("pageDrawCanvas")) {
    const drawCanvas = document.createElement("canvas");
    drawCanvas.id = "pageDrawCanvas";
    drawCanvas.className = "page-draw-canvas";
    slideStage.insertBefore(drawCanvas, el.overlayLayer || null);
    el.pageDrawCanvas = drawCanvas;
  }
  if (pagesStage && !pagesStage.dataset.layoutUpgraded) {
    const pageFooter = Array.from(pagesStage.children).find((node) => node.classList?.contains("stage-footer"));
    if (pageFooter && !pageFooter.querySelector("#pagesNextHintBtn")) {
      const placeholder = document.createElement("button");
      placeholder.type = "button";
      placeholder.className = "btn ghost";
      placeholder.id = "pagesNextHintBtn";
      placeholder.disabled = true;
      placeholder.textContent = "\u4e0b\u4e00\u6b65";
      pageFooter.appendChild(placeholder);
    }
    pagesStage.dataset.layoutUpgraded = "true";
  }
  const pageFooter = Array.from(pagesStage?.children || []).find((node) => node.classList?.contains("stage-footer"));
  if (pageFooter) {
    pageFooter.classList.add("pages-workbench-footer");
    const exportButton = el.exportWorkflowPptBtn || document.getElementById("exportWorkflowPptBtn");
    const exportRow = exportButton?.closest?.(".export-deck-row");
    if (exportButton && !pageFooter.contains(exportButton)) {
      const nextHint = pageFooter.querySelector("#pagesNextHintBtn");
      pageFooter.insertBefore(exportButton, nextHint || null);
    }
    if (exportRow) {
      exportRow.hidden = true;
    }
  }
}

function renderPagesWorkbench() {
  renderPageList();
  const page = getSelectedPage();
  if (!page) {
    el.pageMetaHint.textContent = state.workflowJob?.statusText || "";
    if (el.pageOnscreenTitleEditor) el.pageOnscreenTitleEditor.value = "";
    if (el.pageOnscreenBodyEditor) el.pageOnscreenBodyEditor.value = "";
    if (el.pageOnscreenEditor) el.pageOnscreenEditor.value = "";
    el.pageExtraPrompt.value = "";
    el.pagePromptTrace.textContent = "";
    if (el.viewCurrentPageLargeBtn) el.viewCurrentPageLargeBtn.disabled = true;
    if (el.copyPagePromptBtn) el.copyPagePromptBtn.disabled = true;
    renderArtboard();
    renderPageResults();
    syncCurrentPageGenerateUi();
    return;
  }

  const draft = syncPageDraftFromPage(page);
  const sourceText = formatOnscreenPreview(page.onscreenContentText || page.onscreenContent || page.pageContent || "");
  draft.onscreenTitle = normalizeDisplayText(draft.onscreenTitle || page.pageTitle || "").split("\n")[0].trim();
  draft.onscreenBody = formatOnscreenPreview(draft.onscreenBody || sourceText);
  draft.onscreenContent = composeOnscreenContentFromEditors(draft.onscreenTitle, draft.onscreenBody);

  el.pageMetaHint.textContent = page.riskReason
    ? `\u7b2c${page.pageNumber}\u9875 \u00b7 ${page.pageTitle} \u00b7 \u6392\u7248\u98ce\u9669`
    : `\u7b2c${page.pageNumber}\u9875 \u00b7 ${page.pageTitle}`;
  if (el.pageOnscreenTitleEditor) el.pageOnscreenTitleEditor.value = draft.onscreenTitle;
  if (el.pageOnscreenBodyEditor) el.pageOnscreenBodyEditor.value = draft.onscreenBody;
  if (el.pageOnscreenPreview) el.pageOnscreenPreview.innerHTML = "";
  el.pageExtraPrompt.value = draft.extraPrompt || "";
  el.pagePromptTrace.textContent = stringifyTrace(page.promptTrace);
  if (el.viewCurrentPageLargeBtn) {
    el.viewCurrentPageLargeBtn.disabled = !page.baseImage || isPageGenerating(page.id);
  }
  renderArtboard();
  renderPageResults(page);
  syncCurrentPageGenerateUi();
}

async function submitCurrentPageReprepare(options = {}) {
  const page = getSelectedPage();
  if (!page) return;
  const { autoExpandToMaxChars = false } = options;
  const draft = ensurePageDraft(page);
  draft.onscreenContent = updateCurrentPageDraftFromEditors();
  if (!draft.onscreenContent) {
    setStatus("\u8bf7\u5148\u586b\u5199\u5f53\u524d\u9875\u7684\u6807\u9898\u6216\u6b63\u6587\u3002", "error");
    return;
  }
  const actionKey = autoExpandToMaxChars ? "repolish" : "reprepare";
  const actionButton = autoExpandToMaxChars ? el.aiRepolishPageBtn : el.repreparePageBtn;
  const signal = startCancelableAction(actionKey, actionButton, el.cancelRepreparePageBtn, autoExpandToMaxChars ? "重润中..." : "\u6574\u7406\u4e2d...");
  setStatus(
    autoExpandToMaxChars
      ? `\u6b63\u5728\u6309\u6700\u5927\u5b57\u6570 AI \u91cd\u6da6\u7b2c${page.pageNumber}\u9875...`
      : `\u6b63\u5728\u91cd\u65b0\u6574\u7406\u7b2c${page.pageNumber}\u9875...`,
    "running"
  );
  try {
    const data = await apiJson("/api/workflow/page/reprepare", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      signal,
      body: JSON.stringify({
        apiKey: state.settings.apiKey,
        region: state.settings.region,
        jobId: state.workflowJobId,
        pageId: page.id,
        onscreenContent: draft.onscreenContent,
        autoExpandToMaxChars,
      }),
    });
    state.workflowJob = sanitizeRecoveredWorkflowJob(data.job);
    ensureSelectedPage();
    syncPageDraftFromPage(getSelectedPage(), { force: true });
    renderPagesWorkbench();
    setStatus(
      autoExpandToMaxChars
        ? `\u7b2c${page.pageNumber}\u9875\u5df2\u6309\u6700\u5927\u5b57\u6570 AI \u91cd\u6da6\u3002`
        : `\u7b2c${page.pageNumber}\u9875\u5df2\u91cd\u65b0\u6574\u7406\u3002`,
      "success"
    );
    saveState();
  } catch (error) {
    if (isAbortError(error)) return;
    if (isMissingWorkflowJobError(error)) {
      clearWorkflowSession({ toSplit: true });
      setStatus("\u4e4b\u524d\u7684\u62c6\u5206\u4efb\u52a1\u5df2\u5931\u6548\uff0c\u8bf7\u91cd\u65b0\u62c6\u5206\u3002", "error");
      return;
    }
    setStatus(
      error.message || (autoExpandToMaxChars ? "AI 一键重润失败。" : "\u91cd\u65b0\u6574\u7406\u5931\u8d25\u3002"),
      "error"
    );
  } finally {
    finishCancelableAction(actionKey);
  }
}

async function reprepareCurrentPage() {
  return submitCurrentPageReprepare();
}

async function aiRepolishCurrentPage() {
  return submitCurrentPageReprepare({ autoExpandToMaxChars: true });
}

async function copyCurrentPagePrompt() {
  const page = getSelectedPage();
  if (!page) return;
  const draft = ensurePageDraft(page);
  draft.extraPrompt = el.pageExtraPrompt.value.trim();
  draft.onscreenContent = updateCurrentPageDraftFromEditors();
  const promptTrace = page.promptTrace?.finalImage || null;
  const pageContent = String(page.onscreenContentText || page.onscreenContent || page.pageContent || "").trim();
  const promptIsCurrent = Boolean(promptTrace?.prompt)
    && String(promptTrace.extraPrompt || "").trim() === draft.extraPrompt
    && pageContent === String(draft.onscreenContent || "").trim();
  let prompt = promptIsCurrent ? getFinalPromptFromPage(page) : "";

  if (!prompt) {
    const requestKey = `copyPrompt:${page.id}`;
    const signal = startCancelableAction(requestKey, el.copyPagePromptBtn, null, "准备中...");
    setStatus(`正在准备第${page.pageNumber}页最终提示词...`, "running");
    try {
      const canvasImage = await exportCurrentArtboard();
      const data = await apiJson("/api/workflow/page/prompt", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        signal,
        body: JSON.stringify({
          apiKey: state.settings.apiKey,
          region: state.settings.region,
          imageModel: getCurrentWorkflowImageModel(),
          jobId: state.workflowJob?.id,
          pageId: page.id,
          extraPrompt: draft.extraPrompt,
          canvasImage,
          onscreenContent: draft.onscreenContent,
          enableGeminiGoogleSearch: state.settings.enableGeminiGoogleSearch,
        }),
      });
      if (data.page && state.workflowJob?.pages) {
        state.workflowJob.pages = state.workflowJob.pages.map((item) => item.id === data.page.id ? data.page : item);
        state.workflowJob = sanitizeRecoveredWorkflowJob(state.workflowJob);
      }
      prompt = String(data.finalPrompt || data.page?.promptTrace?.finalImage?.prompt || "").trim();
      saveState();
    } catch (error) {
      if (isAbortError(error)) return;
      setStatus(error.message || "准备提示词失败。", "error");
      return;
    } finally {
      finishCancelableAction(requestKey);
      renderPagesWorkbench();
    }
  }

  if (!prompt) {
    setStatus("当前页还没有可复制的最终提示词。", "error");
    return;
  }
  try {
    const ok = await copyTextToClipboard(prompt);
    if (!ok) throw new Error("浏览器拒绝复制。");
    setStatus(`第${page.pageNumber}页完整提示词已复制。`, "success");
  } catch (error) {
    setStatus(error.message || "复制提示词失败。", "error");
  }
}

async function generateCurrentPage() {
  const page = getSelectedPage();
  if (!page) return;
  const currentPageId = page.id;
  const currentPageNumber = page.pageNumber;

  await ensureServerConfigReady();
  const selectedImageModel = getCurrentWorkflowImageModel();
  if (usingHostedWorkflowModel() && !hasHostedImageApiKey()) {
    setStatus("请先填写生图 API Key。", "error");
    switchTab("settings");
    return;
  }
  if (!usingHostedWorkflowModel() && !hasDashScopeApiKey()) {
    setStatus("\u8bf7\u5148\u586b\u5199 DashScope / Qwen API Key\u3002", "error");
    switchTab("settings");
    return;
  }

  const draft = ensurePageDraft(page);
  draft.extraPrompt = el.pageExtraPrompt.value.trim();
  draft.onscreenContent = updateCurrentPageDraftFromEditors();
  const requestKey = getPageGenerateRequestKey(page.id);
  const signal = startCancelableAction(requestKey, null, null);
  page.generationStatus = "preparing";
  page.generationError = "";
  renderPagesWorkbench();
  setStatus(`\u7b2c${page.pageNumber}\u9875\u5df2\u8fdb\u5165\u751f\u6210\u961f\u5217\uff0c\u6b63\u5728\u51c6\u5907\u63d0\u793a\u8bcd...`, "running");
  try {
    const canvasImage = await exportCurrentArtboard();
    const generatePromise = apiJson("/api/workflow/page/generate-v2", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      signal,
      body: JSON.stringify({
        apiKey: state.settings.apiKey,
        googleApiKey: state.settings.googleApiKey,
        grsaiHost: state.settings.grsaiHost,
        region: state.settings.region,
        imageModel: selectedImageModel,
        jobId: state.workflowJobId,
        pageId: page.id,
        slideAspect: state.settings.slideAspect,
        size: getWorkflowGenerationSize(),
        seed: state.settings.seed,
        extraPrompt: draft.extraPrompt,
        onscreenContent: draft.onscreenContent,
        canvasImage,
      }),
    });
    const nextPage = getNextWorkflowPage(currentPageId);
    if (nextPage && state.selectedPageId === currentPageId) {
      state.selectedPageId = nextPage.id;
      renderPagesWorkbench();
      saveState();
    }
    const data = await generatePromise;
    if (!data.page?.generated) {
      throw new Error(data.page?.generationError || "\u8fd9\u4e00\u9875\u6ca1\u6709\u62ff\u5230\u56fe\u7247\u7ed3\u679c\u3002");
    }
    state.workflowJob.pages = state.workflowJob.pages.map((item) => item.id === data.page.id ? data.page : item);
    state.workflowJob = sanitizeRecoveredWorkflowJob(state.workflowJob);
    setStatus(`\u7b2c${currentPageNumber}\u9875\u5df2\u751f\u6210\u3002`, "success");
    saveState();
  } catch (error) {
    if (isAbortError(error)) return;
    if (isMissingWorkflowJobError(error)) {
      clearWorkflowSession({ toSplit: true });
      setStatus("\u4e4b\u524d\u7684\u62c6\u5206\u4efb\u52a1\u5df2\u5931\u6548\uff0c\u8bf7\u91cd\u65b0\u62c6\u5206\u3002", "error");
      return;
    }
    const current = state.workflowJob?.pages?.find((item) => item.id === currentPageId);
    if (current) {
      current.generationStatus = "error";
      current.generationError = error.message || "\u751f\u6210\u5931\u8d25\u3002";
    }
    setStatus(error.message || "\u751f\u6210\u5931\u8d25\u3002", "error");
  } finally {
    finishCancelableAction(requestKey);
    renderPagesWorkbench();
    syncCurrentPageGenerateUi();
  }
}

function initialize() {
  cacheElements();
  loadState();
  state.workflowJob = sanitizeRecoveredWorkflowJob(state.workflowJob);
  refreshServerConfig();
  applyStateToUi();
  renderPreferenceSummary();
  renderSplitPresets();
  renderReferenceFiles();
  renderHistoryProjects();
  updateThemeView();
  ensureSelectedPage();
  renderPagesWorkbench();
  renderRevise();
  switchTab(state.activeTab);
  switchSmartStep(state.smartStep);
  upgradeSmartUiLayout();
  renderPagesWorkbench();
  bindEvents();
  attachEnhancedPageEditorEvents();
  el.workflowTargetChars?.addEventListener("change", () => {
    state.workflowTargetChars = clamp(Number(el.workflowTargetChars.value || 0), 0, 300);
    if (state.workflowTargetChars && state.workflowMaxChars && state.workflowTargetChars > state.workflowMaxChars) {
      state.workflowTargetChars = state.workflowMaxChars;
      el.workflowTargetChars.value = String(state.workflowTargetChars);
    }
    saveState();
  });
  el.workflowMaxChars?.addEventListener("change", () => {
    state.workflowMaxChars = clamp(Number(el.workflowMaxChars.value || 200), 0, 400);
    if (state.workflowTargetChars && state.workflowMaxChars && state.workflowTargetChars > state.workflowMaxChars) {
      state.workflowTargetChars = state.workflowMaxChars;
      if (el.workflowTargetChars) el.workflowTargetChars.value = String(state.workflowTargetChars);
    }
    saveState();
  });
  setupPageDrawingInteractions();
  renderArtboard();
  updatePageDrawToolbar();
  setupReviseCanvasInteractions();
  syncWorkflowJobOnce();
  if (state.workflowJobId && state.workflowJob?.status !== "ready") {
    startWorkflowPolling();
  }
}

document.addEventListener("DOMContentLoaded", initialize);
