// 基于 v2 的 app.js 核心逻辑，针对 v3 进行存储隔离和视觉兼容调整
const STORAGE_KEY = "ppt-studio-v3-glass"; // 独立存储键名
const DEFAULT_REGION = "beijing";
const PPT_MODEL = "gemini-3.1-flash-image-preview";

// 保持原有的状态结构
const state = {
  activeTab: "smart",
  smartStep: "theme",
  settings: {
    apiKey: "",
    googleApiKey: "",
    workflowImageModel: PPT_MODEL,
    region: DEFAULT_REGION,
    slideAspect: "16:9",
    outputSize: "2K",
    seed: "",
  },
  workspaceZoom: 100,
  themeName: "",
  decorationLevel: "medium",
  preferences: {
    styleMode: "business",
    layoutVariety: "balanced",
    detailLevel: "polished",
    visualDensity: "balanced",
    compositionFocus: "balanced",
    dataNarrative: "balanced",
    pageMood: "modern",
  },
  themeDefinition: null,
  themePromptTrace: null,
  themeConfirmed: false,
  workflowContent: "",
  workflowPageCount: 8,
  aiProcessingMode: "balanced",
  workflowJobId: "",
  workflowJob: null,
  selectedPageId: "",
  pageDrafts: {},
};

// 工具函数：获取元素
const el = {};
function cacheElements() {
  const ids = [
    "workspaceZoomRange", "workspaceZoomValue", "themeName", "themeDecorationLevel",
    "prefStyleMode", "prefLayoutVariety", "prefDetailLevel", "prefVisualDensity",
    "prefCompositionFocus", "prefPageMood", "preferenceSummary", "generateThemeBtn",
    "confirmThemeBtn", "goSplitBtn", "themeSummaryPreview", "themeModelPrompt",
    "workflowPageCount", "workflowContent", "splitTemplateInput", "aiProcessingMode",
    "runSplitBtn", "workflowStats", "workflowPageList", "pageMetaHint", "pageOnscreenEditor",
    "repreparePageBtn", "batchGenerateReadyBtn", "slideBaseImage", "slideEmptyState",
    "generateCurrentPageBtn", "pageExtraPrompt", "apiKey", "googleApiKey",
    "workflowImageModel", "region", "slideAspect", "seed", "testApiKeyBtn", "statusToast"
  ];
  ids.forEach(id => {
    el[id] = document.getElementById(id);
  });
}

// 状态持久化
function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function loadState() {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (saved) {
    Object.assign(state, JSON.parse(saved));
  }
}

// UI 更新逻辑 (简化示例，实际应包含 v2 的完整业务逻辑)
function switchTab(tab) {
  state.activeTab = tab;
  document.querySelectorAll(".sidebar-tab").forEach(btn => {
    btn.classList.toggle("is-active", btn.dataset.tab === tab);
  });
  document.querySelectorAll(".view").forEach(view => {
    view.classList.toggle("is-active", view.dataset.view === tab);
  });
  saveState();
}

function switchSmartStep(step) {
  state.smartStep = step;
  document.querySelectorAll(".ribbon-step").forEach(btn => {
    btn.classList.toggle("is-active", btn.dataset.step === step);
  });
  document.querySelectorAll(".smart-stage").forEach(panel => {
    panel.classList.toggle("is-active", panel.dataset.stepPanel === step);
  });
  saveState();
}

// 初始化
function init() {
  cacheElements();
  loadState();
  
  // 绑定基础事件
  document.querySelectorAll(".sidebar-tab").forEach(btn => {
    btn.addEventListener("click", () => switchTab(btn.dataset.tab));
  });
  
  document.querySelectorAll(".ribbon-step").forEach(btn => {
    btn.addEventListener("click", () => switchSmartStep(btn.dataset.step));
  });

  // 更多事件绑定... (实际应复制 v2 的业务方法)
  
  console.log("PPT Studio v3 Initialized with Glassmorphism 2.0");
}

window.addEventListener("DOMContentLoaded", init);
