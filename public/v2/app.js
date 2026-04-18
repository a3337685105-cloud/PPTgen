const STORAGE_KEY = "ppt-studio-v2-mainline";
const DEFAULT_REGION = "beijing";
const PPT_MODEL = "gemini-3-pro-image-preview";
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
  styleMode: { business: "商务", academic: "学术", creative: "创意" },
  layoutVariety: { uniform: "统一稳定", balanced: "平衡变化", diverse: "尽量多样" },
  detailLevel: { minimal: "偏简约", polished: "精致平衡", rich: "偏精细" },
  visualDensity: { airy: "留白更多", balanced: "均衡信息量", dense: "信息更满" },
  compositionFocus: { imageLead: "视觉主导", balanced: "图文平衡", textLead: "内容主导" },
  dataNarrative: { clean: "清晰克制", balanced: "适度信息图", expressive: "更强视觉化" },
  pageMood: { steady: "稳重统一", modern: "现代清爽", dramatic: "更有冲击" },
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
  "16:9": { width: 1600, height: 900 },
  "4:3": { width: 1400, height: 1050 },
  "1:1": { width: 1200, height: 1200 },
};

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
  preferences: { ...DEFAULT_PREFERENCES },
  themeDefinition: null,
  themePromptTrace: null,
  themeConfirmed: false,
  workflowContent: "",
  workflowPageCount: 8,
  splitPresetId: "balanced",
  splitTemplateText: SPLIT_PRESETS[0].text,
  parsedFiles: [],
  workflowJobId: "",
  workflowJob: null,
  workflowPollTimer: null,
  selectedPageId: "",
  pageDrafts: {},
  revise: {
    images: [],
    selectedImageId: "",
    prompt: "",
    results: [],
    drawing: null,
  },
};

const el = {};

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
    "confirmThemeBtn",
    "goSplitBtn",
    "themeStatus",
    "themeSummaryPreview",
    "themeModelPrompt",
    "workflowPageCount",
    "workflowContent",
    "splitTemplateInput",
    "splitPresetToolbar",
    "pickReferenceFilesBtn",
    "referenceFilesInput",
    "referenceFilesList",
    "runSplitBtn",
    "backToThemeBtn",
    "backToSplitBtn",
    "workflowSummary",
    "workflowStats",
    "workflowDiagnostics",
    "workflowPromptTrace",
    "workflowRibbonMeta",
    "workflowPageList",
    "pageMetaHint",
    "pageOnscreenEditor",
    "repreparePageBtn",
    "batchGenerateReadyBtn",
    "uploadOverlayBtn",
    "overlayFileInput",
    "clearOverlayBtn",
    "slideStage",
    "slideFrame",
    "slideBaseImage",
    "slideEmptyState",
    "overlayLayer",
    "generateCurrentPageBtn",
    "pageExtraPrompt",
    "pagePromptTrace",
    "pageResultStrip",
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
    "reviseResultStrip",
    "apiKey",
    "googleApiKey",
    "workflowImageModel",
    "region",
    "slideAspect",
    "outputSize",
    "seed",
    "testApiKeyBtn",
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

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function saveState() {
  const draftToStore = Object.fromEntries(Object.entries(state.pageDrafts).map(([pageId, draft]) => [
    pageId,
    {
      onscreenContent: draft.onscreenContent || "",
      extraPrompt: draft.extraPrompt || "",
      overlays: (draft.overlays || []).filter((item) => /^https?:\/\//i.test(item.src) || item.src.startsWith("/generated-images/")),
    },
  ]));
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
    themeConfirmed: state.themeConfirmed,
    workflowContent: state.workflowContent,
    workflowPageCount: state.workflowPageCount,
    splitPresetId: state.splitPresetId,
    splitTemplateText: state.splitTemplateText,
    parsedFiles: state.parsedFiles,
    workflowJobId: state.workflowJobId,
    workflowJob: state.workflowJob,
    selectedPageId: state.selectedPageId,
    pageDrafts: draftToStore,
  };
  localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
}

function loadState() {
  const parsed = safeJsonParse(localStorage.getItem(STORAGE_KEY) || "");
  if (!parsed || typeof parsed !== "object") return;
  state.activeTab = ["smart", "revise", "settings"].includes(parsed.activeTab) ? parsed.activeTab : "smart";
  state.smartStep = ["theme", "split", "pages"].includes(parsed.smartStep) ? parsed.smartStep : "theme";
  state.settings = { ...state.settings, ...(parsed.settings || {}) };
  state.workspaceZoom = clamp(Number(parsed.workspaceZoom || 100), 50, 140);
  state.themeName = String(parsed.themeName || "");
  state.decorationLevel = String(parsed.decorationLevel || "medium");
  state.preferences = { ...DEFAULT_PREFERENCES, ...(parsed.preferences || {}) };
  state.themeDefinition = parsed.themeDefinition || null;
  state.themePromptTrace = parsed.themePromptTrace || null;
  state.themeConfirmed = Boolean(parsed.themeConfirmed);
  state.workflowContent = String(parsed.workflowContent || "");
  state.workflowPageCount = clamp(Number(parsed.workflowPageCount || 8), 2, 20);
  state.splitPresetId = String(parsed.splitPresetId || "balanced");
  state.splitTemplateText = String(parsed.splitTemplateText || SPLIT_PRESETS[0].text);
  state.parsedFiles = Array.isArray(parsed.parsedFiles) ? parsed.parsedFiles : [];
  state.workflowJobId = String(parsed.workflowJobId || "");
  state.workflowJob = parsed.workflowJob || null;
  state.selectedPageId = String(parsed.selectedPageId || "");
  state.pageDrafts = parsed.pageDrafts && typeof parsed.pageDrafts === "object" ? parsed.pageDrafts : {};
}

function applyStateToUi() {
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
  el.workflowContent.value = state.workflowContent;
  el.splitTemplateInput.value = state.splitTemplateText;
  el.apiKey.value = state.settings.apiKey || "";
  el.googleApiKey.value = state.settings.googleApiKey || "";
  el.workflowImageModel.value = state.settings.workflowImageModel || PPT_MODEL;
  el.region.value = state.settings.region || DEFAULT_REGION;
  el.slideAspect.value = state.settings.slideAspect || "16:9";
  el.outputSize.value = state.settings.outputSize || "2K";
  el.seed.value = state.settings.seed || "";
  el.revisePrompt.value = state.revise.prompt || "";
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

function switchTab(tab) {
  state.activeTab = tab;
  document.querySelectorAll(".sidebar-tab").forEach((button) => {
    const active = button.dataset.tab === tab;
    button.classList.toggle("is-active", active);
  });
  document.querySelectorAll(".view").forEach((view) => {
    view.classList.toggle("is-active", view.dataset.view === tab);
  });
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
    theme: "先把全局风格确认下来，再进入拆分。",
    split: "主文本是主输入，参考文件只是补充材料。",
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

function ensurePageDraft(page) {
  if (!page) return null;
  if (!state.pageDrafts[page.id]) {
    state.pageDrafts[page.id] = {
      onscreenContent: page.onscreenContent || "",
      extraPrompt: page.extraPrompt || "",
      overlays: [],
    };
  }
  if (!state.pageDrafts[page.id].onscreenContent && page.onscreenContent) {
    state.pageDrafts[page.id].onscreenContent = page.onscreenContent;
  }
  if (!state.pageDrafts[page.id].extraPrompt && page.extraPrompt) {
    state.pageDrafts[page.id].extraPrompt = page.extraPrompt;
  }
  if (!Array.isArray(state.pageDrafts[page.id].overlays)) {
    state.pageDrafts[page.id].overlays = [];
  }
  return state.pageDrafts[page.id];
}

function updateThemeView() {
  el.confirmThemeBtn.disabled = !state.themeDefinition;
  el.goSplitBtn.disabled = !state.themeConfirmed;
  el.themeSummaryPreview.textContent = state.themeDefinition?.displaySummaryZh || "风格摘要会显示在这里。";
  el.themeModelPrompt.textContent = state.themePromptTrace
    ? stringifyTrace(state.themePromptTrace)
    : (state.themeDefinition?.modelPrompt || "还没有生成模型总纲。");
}

function renderSplitPresets() {
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

function renderReferenceFiles() {
  if (!state.parsedFiles.length) {
    el.referenceFilesList.innerHTML = `<div class="inline-hint">暂无参考文件</div>`;
    return;
  }
  if (!state.parsedFiles.length) {
    el.referenceFilesList.innerHTML = `<div class="inline-hint">还没有上传参考文件。若主文本已经足够，可以直接拆分。</div>`;
    return;
  }
  el.referenceFilesList.innerHTML = state.parsedFiles.map((file) => `
    <div class="file-item">
      <strong>${escapeHtml(file.name)}</strong>
      <div class="file-meta">
        <span class="meta-pill">${escapeHtml(file.category || "unknown")}</span>
        <span class="meta-pill">${escapeHtml(file.parseStatus || "unknown")}</span>
      </div>
      <p class="card-copy">${escapeHtml(file.parseNote || "已加入参考材料。")}</p>
      ${file.previewText ? `<details class="trace-details"><summary>预览文本</summary><pre>${escapeHtml(file.previewText)}</pre></details>` : ""}
    </div>
  `).join("");
}

function formatJobStats(job) {
  if (!job) return "还没有开始拆分。";
  return [
    `总页数 ${job.totalPages || 0}`,
    `已准备 ${job.preparedPages || 0}`,
    `可生成 ${job.readyToGeneratePages || 0}`,
    `失败 ${job.failedPages || 0}`,
  ].join(" · ");
}

function renderPageList() {
  const job = state.workflowJob;
  if (!job?.pages?.length) {
    el.workflowSummary.textContent = "";
    el.workflowStats.textContent = "";
    el.workflowDiagnostics.textContent = "";
    el.workflowPromptTrace.textContent = "";
    el.workflowPageList.innerHTML = `<div class="inline-hint">还没有页面</div>`;
    return;
  }
  if (!job?.pages?.length) {
    el.workflowSummary.textContent = "先完成拆分，这里会逐页显示页面状态。";
    el.workflowStats.textContent = "暂无任务";
    el.workflowDiagnostics.textContent = "暂无拆分诊断。";
    el.workflowPromptTrace.textContent = "还没有工作流流档。";
    el.workflowPageList.innerHTML = `<div class="inline-hint">还没有页面结果。</div>`;
    return;
  }

  ensureSelectedPage();
  el.workflowSummary.textContent = job.documentSummary || "拆分页已经生成。";
  el.workflowStats.textContent = formatJobStats(job);
  el.workflowDiagnostics.textContent = job.splitDiagnostics || "本次拆分没有返回额外诊断。";
  el.workflowPromptTrace.textContent = stringifyTrace(job.promptTrace);
  el.workflowPageList.innerHTML = job.pages.map((page) => {
    const status = page.generated ? "已生成" : page.readyToGenerate ? "已排版" : page.prepareDone ? "待确认" : "已拆分";
    const statusClass = page.generated ? "generated" : page.readyToGenerate ? "ready" : "idle";
    const riskClass = page.riskLevel === "high" ? "high" : page.riskLevel === "medium" ? "medium" : "";
    return `
      <div class="page-item ${page.id === state.selectedPageId ? "is-active" : ""}" data-page-id="${page.id}">
        <div class="page-title-row">
          <strong>第 ${page.pageNumber} 页 · ${escapeHtml(page.pageTitle || "未命名")}</strong>
          <span class="status-pill ${statusClass}">${status}</span>
        </div>
        <div class="page-meta">
          <span class="meta-pill">${escapeHtml(page.pageType)}</span>
          ${riskClass ? `<span class="risk-pill ${riskClass}">${page.riskLevel === "high" ? "风险" : "注意"}</span>` : ""}
        </div>
        ${page.riskReason ? `<p class="card-copy">${escapeHtml(page.riskReason)}</p>` : ""}
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

function stringifyTrace(trace) {
  if (!trace) return "还没有 prompt trace。";
  return JSON.stringify(trace, null, 2);
}

function renderPagesWorkbench() {
  renderPageList();
  const page = getSelectedPage();
  if (!page) {
    el.pageMetaHint.textContent = "";
    el.pagePromptTrace.textContent = "";
    el.pageMetaHint.textContent = "还没有页面可编辑。";
    el.pageOnscreenEditor.value = "";
    el.pageExtraPrompt.value = "";
    el.pagePromptTrace.textContent = "还没有 prompt trace。";
    el.pageResultStrip.innerHTML = "";
    renderArtboard();
    return;
  }
  const draft = ensurePageDraft(page);
  el.pageMetaHint.textContent = page.riskReason
    ? `第 ${page.pageNumber} 页 · ${page.pageTitle} · 需确认`
    : `第 ${page.pageNumber} 页 · ${page.pageTitle}`;
  el.pageMetaHint.textContent = page.riskReason
    ? `当前页：第 ${page.pageNumber} 页 · ${page.pageTitle}。左侧已标出风险。`
    : `当前页：第 ${page.pageNumber} 页 · ${page.pageTitle}。`;
  el.pageMetaHint.textContent = page.riskReason
    ? `第 ${page.pageNumber} 页 · ${page.pageTitle} · 需确认`
    : `第 ${page.pageNumber} 页 · ${page.pageTitle}`;
  el.pageOnscreenEditor.value = draft.onscreenContent || page.onscreenContent || page.pageContent || "";
  el.pageExtraPrompt.value = draft.extraPrompt || "";
  el.pagePromptTrace.textContent = stringifyTrace(page.promptTrace);
  renderArtboard();
  renderPageResults(page);
}

function renderPageResults(page) {
  if (!page?.resultImages?.length) {
    el.pageResultStrip.innerHTML = "";
    return;
  }
  el.pageResultStrip.innerHTML = page.resultImages.map((src) => `
    <div class="result-item">
      <img src="${escapeHtml(src)}" alt="页面结果图" />
    </div>
  `).join("");
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
    return;
  }
  if (baseImage) {
    el.slideBaseImage.src = baseImage;
    el.slideBaseImage.hidden = false;
    el.slideEmptyState.hidden = true;
  } else {
    el.slideBaseImage.hidden = true;
    el.slideBaseImage.src = "";
    el.slideEmptyState.hidden = false;
  }

  const overlays = draft?.overlays || [];
  el.overlayLayer.innerHTML = overlays.map((item) => `
    <div class="overlay-item" data-overlay-id="${item.id}" style="left:${item.x}%;top:${item.y}%;width:${item.w}%;height:${item.h}%;">
      <img src="${escapeHtml(item.src)}" alt="overlay" />
    </div>
  `).join("");

  el.overlayLayer.querySelectorAll(".overlay-item").forEach((node) => {
    const overlayId = node.dataset.overlayId;
    node.addEventListener("pointerdown", (event) => beginOverlayDrag(event, overlayId));
  });
}

function beginOverlayDrag(event, overlayId) {
  const page = getSelectedPage();
  const draft = ensurePageDraft(page);
  const overlay = draft.overlays.find((item) => item.id === overlayId);
  if (!overlay) return;
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
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("图片加载失败。"));
    image.src = src;
  });
}

async function exportCurrentArtboard() {
  const page = getSelectedPage();
  if (!page) return "";
  const draft = ensurePageDraft(page);
  if (!page.baseImage && !draft.overlays.length) return "";
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
}

async function apiJson(url, options = {}) {
  const response = await fetch(url, options);
  const data = await response.json();
  if (!response.ok || data.code) {
    throw new Error(data.message || "请求失败。");
  }
  return data;
}

async function generateTheme() {
  state.themeName = el.themeName.value.trim();
  state.decorationLevel = el.themeDecorationLevel.value;
  state.preferences = getCurrentPreferences();
  if (!state.settings.apiKey) {
    setStatus("请先填写 DashScope / Qwen API Key。", "error");
    switchTab("settings");
    return;
  }
  if (!state.themeName) {
    setStatus("请先输入风格主题。", "error");
    return;
  }
  setButtonLoading(el.generateThemeBtn, true, "生成中...");
  setStatus("正在生成风格主题...", "running");
  try {
    const data = await apiJson("/api/workflow/theme", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        apiKey: state.settings.apiKey,
        region: state.settings.region,
        themeName: state.themeName,
        decorationLevel: state.decorationLevel,
        preferences: state.preferences,
      }),
    });
    state.themeDefinition = data.themeDefinition;
    state.themePromptTrace = data.promptTrace || null;
    state.themeConfirmed = false;
    updateThemeView();
    setStatus("风格模板已生成，确认后即可使用。", "success");
    saveState();
  } catch (error) {
    setStatus(error.message || "生成风格失败。", "error");
  } finally {
    setButtonLoading(el.generateThemeBtn, false);
  }
}

function confirmTheme() {
  if (!state.themeDefinition) return;
  state.themeConfirmed = true;
  updateThemeView();
  setStatus("风格已确认，可以进入文本拆分。", "success");
  saveState();
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

async function runSplit() {
  state.workflowContent = el.workflowContent.value.trim();
  state.workflowPageCount = clamp(Number(el.workflowPageCount.value || 8), 2, 20);
  state.splitTemplateText = el.splitTemplateInput.value.trim();
  if (!state.themeConfirmed) {
    setStatus("请先确认风格。", "error");
    switchSmartStep("theme");
    return;
  }
  if (!state.settings.apiKey) {
    setStatus("请先填写 DashScope / Qwen API Key。", "error");
    switchTab("settings");
    return;
  }
  if (!state.workflowContent) {
    setStatus("请先输入主文本。", "error");
    return;
  }

  setButtonLoading(el.runSplitBtn, true, "拆分中...");
  setStatus("正在拆分内容并准备逐页结果...", "running");
  try {
    const data = await apiJson("/api/workflow/split", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        apiKey: state.settings.apiKey,
        region: state.settings.region,
        content: state.workflowContent,
        pageCount: state.workflowPageCount,
        splitTemplate: state.splitTemplateText,
        referenceFiles: state.parsedFiles,
        themeDefinition: state.themeDefinition,
        preferences: state.preferences,
        decorationLevel: state.decorationLevel,
      }),
    });
    state.workflowJobId = data.jobId;
    state.workflowJob = data.job;
    state.selectedPageId = data.job?.pages?.[0]?.id || "";
    ensureSelectedPage();
    switchSmartStep("pages");
    startWorkflowPolling();
    renderPagesWorkbench();
    setStatus("拆分已完成，正在逐页准备上屏内容与排版。", "success");
    saveState();
  } catch (error) {
    setStatus(error.message || "拆分失败。", "error");
  } finally {
    setButtonLoading(el.runSplitBtn, false);
  }
}

function stopWorkflowPolling() {
  if (state.workflowPollTimer) {
    clearInterval(state.workflowPollTimer);
    state.workflowPollTimer = null;
  }
}

function startWorkflowPolling() {
  stopWorkflowPolling();
  if (!state.workflowJobId) return;
  state.workflowPollTimer = setInterval(async () => {
    try {
      const data = await apiJson(`/api/workflow/jobs/${encodeURIComponent(state.workflowJobId)}`);
      state.workflowJob = data.job;
      ensureSelectedPage();
      renderPagesWorkbench();
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
      setStatus(error.message || "读取任务进度失败。", "error");
    }
  }, 2200);
}

async function reprepareCurrentPage() {
  const page = getSelectedPage();
  if (!page) return;
  const draft = ensurePageDraft(page);
  draft.onscreenContent = el.pageOnscreenEditor.value.trim();
  if (!draft.onscreenContent) {
    setStatus("请先填写当前页的上屏内容。", "error");
    return;
  }
  setButtonLoading(el.repreparePageBtn, true, "整理中...");
  setStatus(`正在重新整理第 ${page.pageNumber} 页...`, "running");
  try {
    const data = await apiJson("/api/workflow/page/reprepare", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        apiKey: state.settings.apiKey,
        region: state.settings.region,
        jobId: state.workflowJobId,
        pageId: page.id,
        onscreenContent: draft.onscreenContent,
      }),
    });
    state.workflowJob = data.job;
    ensureSelectedPage();
    renderPagesWorkbench();
    setStatus(`第 ${page.pageNumber} 页已重新整理。`, "success");
    saveState();
  } catch (error) {
    setStatus(error.message || "重新整理失败。", "error");
  } finally {
    setButtonLoading(el.repreparePageBtn, false);
  }
}

async function generateCurrentPage() {
  const page = getSelectedPage();
  if (!page) return;
  if (!state.settings.googleApiKey) {
    setStatus("请先填写 Google API Key。", "error");
    switchTab("settings");
    return;
  }
  const draft = ensurePageDraft(page);
  draft.extraPrompt = el.pageExtraPrompt.value.trim();
  setButtonLoading(el.generateCurrentPageBtn, true, "生成中...");
  setStatus(`正在生成第 ${page.pageNumber} 页...`, "running");
  try {
    const canvasImage = await exportCurrentArtboard();
    const data = await apiJson("/api/workflow/page/generate-v2", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        googleApiKey: state.settings.googleApiKey,
        jobId: state.workflowJobId,
        pageId: page.id,
        slideAspect: state.settings.slideAspect,
        size: state.settings.outputSize,
        seed: state.settings.seed,
        extraPrompt: draft.extraPrompt,
        canvasImage,
      }),
    });
    state.workflowJob.pages = state.workflowJob.pages.map((item) => item.id === data.page.id ? data.page : item);
    renderPagesWorkbench();
    setStatus(`第 ${page.pageNumber} 页已生成。`, "success");
    saveState();
  } catch (error) {
    setStatus(error.message || "生成失败。", "error");
  } finally {
    setButtonLoading(el.generateCurrentPageBtn, false);
  }
}

async function batchGenerateReadyPages() {
  const job = state.workflowJob;
  if (!job?.pages?.length) return;
  if (!state.settings.googleApiKey) {
    setStatus("请先填写 Google API Key。", "error");
    switchTab("settings");
    return;
  }
  const candidates = job.pages.filter((page) => page.readyToGenerate && !page.generated);
  if (!candidates.length) {
    setStatus("还没有可直接批量生成的页面。", "error");
    return;
  }
  setButtonLoading(el.batchGenerateReadyBtn, true, "批量生成中...");
  try {
    for (const page of candidates) {
      const draft = ensurePageDraft(page);
      setStatus(`正在批量生成第 ${page.pageNumber} 页...`, "running");
      const data = await apiJson("/api/workflow/page/generate-v2", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          googleApiKey: state.settings.googleApiKey,
          jobId: state.workflowJobId,
          pageId: page.id,
          slideAspect: state.settings.slideAspect,
          size: state.settings.outputSize,
          seed: state.settings.seed,
          extraPrompt: draft.extraPrompt || "",
          canvasImage: "",
        }),
      });
      state.workflowJob.pages = state.workflowJob.pages.map((item) => item.id === data.page.id ? data.page : item);
      ensureSelectedPage();
      renderPagesWorkbench();
      saveState();
    }
    setStatus("所有已排版页已完成生成。", "success");
  } catch (error) {
    setStatus(error.message || "批量生成失败。", "error");
  } finally {
    setButtonLoading(el.batchGenerateReadyBtn, false);
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

async function testApiKeys() {
  state.settings.apiKey = el.apiKey.value.trim();
  state.settings.googleApiKey = el.googleApiKey.value.trim();
  state.settings.workflowImageModel = el.workflowImageModel.value || PPT_MODEL;
  state.settings.region = el.region.value;
  state.settings.slideAspect = el.slideAspect.value;
  state.settings.outputSize = el.outputSize.value;
  state.settings.seed = el.seed.value.trim();
  if (!state.settings.apiKey && !state.settings.googleApiKey) {
    setStatus("请先填写至少一个可用的 API Key。", "error");
    return;
  }
  setButtonLoading(el.testApiKeyBtn, true, "测试中...");
  setStatus("正在测试 Key...", "running");
  try {
    const tasks = [];
    if (state.settings.googleApiKey) {
      tasks.push(fetch("/api/test-image-key", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          apiKey: state.settings.apiKey,
          googleApiKey: state.settings.googleApiKey,
          region: state.settings.region,
          model: PPT_MODEL,
        }),
      }).then(async (response) => ({ ok: response.ok, data: await response.json() })));
    }
    if (state.settings.apiKey) {
      tasks.push(fetch("/api/test-image-key", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
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
    setStatus(error.message || "测试失败。", "error");
  } finally {
    setButtonLoading(el.testApiKeyBtn, false);
  }
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

async function sendRevise() {
  const image = getCurrentReviseImage();
  state.revise.prompt = el.revisePrompt.value.trim();
  if (!state.settings.apiKey) {
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
      size: state.settings.outputSize,
      n: 1,
    },
  };
  if (state.settings.seed) payload.parameters.seed = Number(state.settings.seed);
  if (Array.isArray(image.boxes) && image.boxes.length) {
    payload.parameters.bbox_list = [image.boxes];
  }

  setButtonLoading(el.sendReviseBtn, true, "改图中...");
  setStatus("正在调用 Wan 改图...", "running");
  try {
    const response = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
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
    image.src = results[0];
    image.boxes = [];
    const loaded = await loadImage(results[0]);
    image.naturalWidth = loaded.naturalWidth;
    image.naturalHeight = loaded.naturalHeight;
    renderRevise();
    setStatus("改图已完成。", "success");
  } catch (error) {
    setStatus(error.message || "改图失败。", "error");
  } finally {
    setButtonLoading(el.sendReviseBtn, false);
  }
}

function normalizeDisplayText(text) {
  return String(text || "").trim();
}

function formatOnscreenPreview(value) {
  const source = String(value || "").trim();
  if (!source) return "";
  const fieldPrefixPattern = /^\s*(title|subtitle|metainfo|metaInfo|abstract|summary|body|content|keypoints?|bullets?|visualelements?|datapoints?|metric|value|highlight|note|label|type)\s*[:：]\s*/i;
  return source
    .split(/\r?\n+/)
    .map((line) => line.replace(fieldPrefixPattern, "").trim())
    .filter((line) => !/^highlight\s*[:：]\s*(true|false)$/i.test(line))
    .filter(Boolean)
    .join("\n");
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
      ${file.previewText ? `<details class="trace-details"><summary>预览文本</summary><pre>${escapeHtml(file.previewText)}</pre></details>` : ""}
    </div>
  `).join("");
}

function confirmTheme() {
  if (!state.themeDefinition) return;
  state.themeConfirmed = true;
  updateThemeView();
  switchSmartStep("split");
  setStatus("风格已确认，继续输入文本并拆分。", "success");
  saveState();
}

function renderPageList() {
  const job = state.workflowJob;
  if (!job?.pages?.length) {
    if (el.workflowSummary) el.workflowSummary.textContent = "";
    el.workflowStats.textContent = "";
    el.workflowDiagnostics.textContent = "";
    el.workflowPromptTrace.textContent = "";
    el.workflowPageList.innerHTML = `<div class="inline-hint">还没有页面</div>`;
    return;
  }

  ensureSelectedPage();
  if (el.workflowSummary) el.workflowSummary.textContent = "";
  el.workflowStats.textContent = formatJobStats(job);
  el.workflowDiagnostics.textContent = normalizeDisplayText(job.splitDiagnostics);
  el.workflowPromptTrace.textContent = stringifyTrace(job.promptTrace);
  el.workflowPageList.innerHTML = job.pages.map((page) => {
    const status = page.generated ? "已生成" : page.readyToGenerate ? "可生成" : page.prepareDone ? "待确认" : "处理中";
    const statusClass = page.generated ? "generated" : page.readyToGenerate ? "ready" : "idle";
    const riskClass = page.riskLevel === "high" ? "high" : page.riskLevel === "medium" ? "medium" : "";
    return `
      <div class="page-item ${page.id === state.selectedPageId ? "is-active" : ""}" data-page-id="${page.id}">
        <div class="page-title-row">
          <strong>第 ${page.pageNumber} 页 · ${escapeHtml(page.pageTitle || "未命名")}</strong>
          <span class="status-pill ${statusClass}">${status}</span>
        </div>
        <div class="page-meta">
          <span class="meta-pill">${escapeHtml(page.pageType || "content")}</span>
          ${riskClass ? `<span class="risk-pill ${riskClass}">${page.riskLevel === "high" ? "风险" : "注意"}</span>` : ""}
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

function renderPagesWorkbench() {
  renderPageList();
  const page = getSelectedPage();
  if (!page) {
    el.pageMetaHint.textContent = "";
    el.pageOnscreenEditor.value = "";
    el.pageExtraPrompt.value = "";
    el.pagePromptTrace.textContent = "";
    el.pageResultStrip.innerHTML = "";
    renderArtboard();
    return;
  }
  const draft = ensurePageDraft(page);
  el.pageMetaHint.textContent = page.riskReason
    ? `第 ${page.pageNumber} 页 · ${page.pageTitle} · 需确认`
    : `第 ${page.pageNumber} 页 · ${page.pageTitle}`;
  const onscreenText = formatOnscreenPreview(draft.onscreenContent || page.onscreenContent || page.pageContent);
  draft.onscreenContent = onscreenText;
  el.pageOnscreenEditor.value = onscreenText;
  el.pageExtraPrompt.value = draft.extraPrompt || "";
  el.pagePromptTrace.textContent = stringifyTrace(page.promptTrace);
  renderArtboard();
  renderPageResults(page);
}

async function generateCurrentPage() {
  const page = getSelectedPage();
  if (!page) return;
  if (!state.settings.googleApiKey) {
    setStatus("请先填写 Google API Key。", "error");
    switchTab("settings");
    return;
  }
  const draft = ensurePageDraft(page);
  draft.extraPrompt = el.pageExtraPrompt.value.trim();
  draft.onscreenContent = el.pageOnscreenEditor.value.trim();
  if (draft.onscreenContent && formatOnscreenPreview(draft.onscreenContent) !== formatOnscreenPreview(page.onscreenContent)) {
    setStatus("这一页上屏内容已改动，请先点“确认修改并重新整理”。", "error");
    return;
  }
  setButtonLoading(el.generateCurrentPageBtn, true, "生成中...");
  setStatus(`正在生成第 ${page.pageNumber} 页...`, "running");
  try {
    const canvasImage = await exportCurrentArtboard();
    const data = await apiJson("/api/workflow/page/generate-v2", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        googleApiKey: state.settings.googleApiKey,
        jobId: state.workflowJobId,
        pageId: page.id,
        slideAspect: state.settings.slideAspect,
        size: state.settings.outputSize,
        seed: state.settings.seed,
        extraPrompt: draft.extraPrompt,
        canvasImage,
      }),
    });
    if (!data.page?.generated) {
      throw new Error(data.page?.generationError || "这页没有拿到图片结果。");
    }
    state.workflowJob.pages = state.workflowJob.pages.map((item) => item.id === data.page.id ? data.page : item);
    renderPagesWorkbench();
    setStatus(`第 ${page.pageNumber} 页已生成。`, "success");
    saveState();
  } catch (error) {
    setStatus(error.message || "生成失败。", "error");
  } finally {
    setButtonLoading(el.generateCurrentPageBtn, false);
  }
}

async function batchGenerateReadyPages() {
  const job = state.workflowJob;
  if (!job?.pages?.length) return;
  if (!state.settings.googleApiKey) {
    setStatus("请先填写 Google API Key。", "error");
    switchTab("settings");
    return;
  }
  const candidates = job.pages.filter((page) => page.readyToGenerate && !page.generated);
  if (!candidates.length) {
    setStatus("还没有可直接批量生成的页面。", "error");
    return;
  }
  setButtonLoading(el.batchGenerateReadyBtn, true, "批量生成中...");
  try {
    for (const page of candidates) {
      const draft = ensurePageDraft(page);
      setStatus(`正在批量生成第 ${page.pageNumber} 页...`, "running");
      const data = await apiJson("/api/workflow/page/generate-v2", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          googleApiKey: state.settings.googleApiKey,
          jobId: state.workflowJobId,
          pageId: page.id,
          slideAspect: state.settings.slideAspect,
          size: state.settings.outputSize,
          seed: state.settings.seed,
          extraPrompt: draft.extraPrompt || "",
          canvasImage: "",
        }),
      });
      if (!data.page?.generated) {
        throw new Error(data.page?.generationError || `第 ${page.pageNumber} 页没有拿到图片结果。`);
      }
      state.workflowJob.pages = state.workflowJob.pages.map((item) => item.id === data.page.id ? data.page : item);
      ensureSelectedPage();
      renderPagesWorkbench();
      saveState();
    }
    setStatus("所有可生成页面已完成。", "success");
  } catch (error) {
    setStatus(error.message || "批量生成失败。", "error");
  } finally {
    setButtonLoading(el.batchGenerateReadyBtn, false);
  }
}

async function syncWorkflowJobOnce() {
  if (!state.workflowJobId) return;
  try {
    const data = await apiJson(`/api/workflow/jobs/${encodeURIComponent(state.workflowJobId)}`);
    state.workflowJob = data.job;
    ensureSelectedPage();
    renderPagesWorkbench();
    saveState();
  } catch {
    // Ignore one-off refresh failures; the user can still rerun split.
  }
}

function formatOnscreenPreview(value) {
  const source = String(value || "").trim();
  if (!source) return "";

  const normalizeKey = (input) => String(input || "").replace(/[\s_-]+/g, "").toLowerCase();
  const wrapperKeys = new Set(["blocks", "items", "points", "entries", "sections", "visualelements", "datapoints"]);
  const hiddenKeys = new Set(["type", "highlight", "index", "order", "sort", "priority"]);
  const lines = [];
  let pendingHeading = "";
  let pendingMetric = "";
  let pendingValue = "";
  let pendingNote = "";

  const pushLine = (text) => {
    const clean = String(text || "").replace(/\s+/g, " ").trim();
    if (!clean) return;
    if (lines[lines.length - 1] === clean) return;
    lines.push(clean);
  };

  const flushHeading = () => {
    if (!pendingHeading) return;
    pushLine(pendingHeading);
    pendingHeading = "";
  };

  const flushMetric = () => {
    if (!pendingMetric) return;
    let line = pendingMetric;
    if (pendingValue) line = `${line}：${pendingValue}`;
    if (pendingNote) line = `${line}（${pendingNote}）`;
    pushLine(line);
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
      pushLine(line);
      return;
    }

    const key = normalizeKey(match[1]);
    const content = String(match[2] || "").trim();

    if (wrapperKeys.has(key) || hiddenKeys.has(key)) return;

    if (key === "title") {
      flushPending();
      pushLine(content);
      return;
    }

    if (key === "subtitle") {
      flushPending();
      pushLine(content ? `副标题：${content}` : "");
      return;
    }

    if (key === "heading" || key === "label") {
      flushPending();
      pendingHeading = content;
      return;
    }

    if (key === "detail") {
      if (pendingHeading) {
        pushLine(content ? `${pendingHeading}：${content}` : pendingHeading);
        pendingHeading = "";
      } else {
        pushLine(content);
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
        pushLine(content);
      }
      return;
    }

    if (["metainfo", "summary", "abstract", "body", "content", "text"].includes(key)) {
      flushPending();
      pushLine(content);
      return;
    }

    flushPending();
    pushLine(content || line);
  });

  flushPending();
  return lines.join("\n");
}

function renderPageList() {
  const job = state.workflowJob;
  if (!job?.pages?.length) {
    if (el.workflowSummary) el.workflowSummary.textContent = "";
    el.workflowStats.textContent = "";
    el.workflowDiagnostics.textContent = "";
    el.workflowPromptTrace.textContent = "";
    el.workflowPageList.innerHTML = `<div class="inline-hint">还没有页面</div>`;
    return;
  }

  ensureSelectedPage();
  if (el.workflowSummary) el.workflowSummary.textContent = "";
  el.workflowStats.textContent = formatJobStats(job);
  el.workflowDiagnostics.textContent = normalizeDisplayText(job.splitDiagnostics);
  el.workflowPromptTrace.textContent = stringifyTrace(job.promptTrace);
  el.workflowPageList.innerHTML = job.pages.map((page) => {
    const status = page.generated ? "已生成" : page.prepareDone ? "可生成" : "处理中";
    const statusClass = page.generated ? "generated" : page.prepareDone ? "ready" : "idle";
    const riskClass = page.riskLevel === "high" ? "high" : page.riskLevel === "medium" ? "medium" : "";
    return `
      <div class="page-item ${page.id === state.selectedPageId ? "is-active" : ""}" data-page-id="${page.id}">
        <div class="page-title-row">
          <strong>第 ${page.pageNumber} 页 · ${escapeHtml(page.pageTitle || "未命名")}</strong>
          <span class="status-pill ${statusClass}">${status}</span>
        </div>
        <div class="page-meta">
          <span class="meta-pill">${escapeHtml(page.pageType || "content")}</span>
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

function renderPagesWorkbench() {
  renderPageList();
  const page = getSelectedPage();
  if (!page) {
    el.pageMetaHint.textContent = "";
    el.pageOnscreenEditor.value = "";
    el.pageExtraPrompt.value = "";
    el.pagePromptTrace.textContent = "";
    el.pageResultStrip.innerHTML = "";
    renderArtboard();
    return;
  }

  const draft = ensurePageDraft(page);
  el.pageMetaHint.textContent = page.riskReason
    ? `第 ${page.pageNumber} 页 · ${page.pageTitle} · 排版风险`
    : `第 ${page.pageNumber} 页 · ${page.pageTitle}`;
  const onscreenText = formatOnscreenPreview(draft.onscreenContent || page.onscreenContent || page.pageContent);
  draft.onscreenContent = onscreenText;
  el.pageOnscreenEditor.value = onscreenText;
  el.pageExtraPrompt.value = draft.extraPrompt || "";
  el.pagePromptTrace.textContent = stringifyTrace(page.promptTrace);
  renderArtboard();
  renderPageResults(page);
}

async function generateCurrentPage() {
  const page = getSelectedPage();
  if (!page) return;
  const imageModel = state.settings.workflowImageModel || PPT_MODEL;
  const requiresGoogle = imageModel === PPT_MODEL;

  if (requiresGoogle && !state.settings.googleApiKey) {
    setStatus("请先填写 Google API Key。", "error");
    switchTab("settings");
    return;
  }
  if (!requiresGoogle && !state.settings.apiKey) {
    setStatus("请先填写 DashScope / Qwen API Key。", "error");
    switchTab("settings");
    return;
  }

  const draft = ensurePageDraft(page);
  draft.extraPrompt = el.pageExtraPrompt.value.trim();
  draft.onscreenContent = el.pageOnscreenEditor.value.trim();
  if (draft.onscreenContent && formatOnscreenPreview(draft.onscreenContent) !== formatOnscreenPreview(page.onscreenContent)) {
    setStatus("这一页上屏内容已改动，请先点“确认修改并重新整理”。", "error");
    return;
  }

  setButtonLoading(el.generateCurrentPageBtn, true, "生成中...");
  setStatus(`正在生成第 ${page.pageNumber} 页...`, "running");
  try {
    const canvasImage = await exportCurrentArtboard();
    const data = await apiJson("/api/workflow/page/generate-v2", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        apiKey: state.settings.apiKey,
        googleApiKey: state.settings.googleApiKey,
        region: state.settings.region,
        imageModel,
        jobId: state.workflowJobId,
        pageId: page.id,
        slideAspect: state.settings.slideAspect,
        size: state.settings.outputSize,
        seed: state.settings.seed,
        extraPrompt: draft.extraPrompt,
        canvasImage,
      }),
    });
    if (!data.page?.generated) {
      throw new Error(data.page?.generationError || "这一页没有拿到图片结果。");
    }
    state.workflowJob.pages = state.workflowJob.pages.map((item) => item.id === data.page.id ? data.page : item);
    renderPagesWorkbench();
    setStatus(`第 ${page.pageNumber} 页已生成。`, "success");
    saveState();
  } catch (error) {
    setStatus(error.message || "生成失败。", "error");
  } finally {
    setButtonLoading(el.generateCurrentPageBtn, false);
  }
}

async function batchGenerateReadyPages() {
  const job = state.workflowJob;
  if (!job?.pages?.length) return;

  const imageModel = state.settings.workflowImageModel || PPT_MODEL;
  const requiresGoogle = imageModel === PPT_MODEL;
  if (requiresGoogle && !state.settings.googleApiKey) {
    setStatus("请先填写 Google API Key。", "error");
    switchTab("settings");
    return;
  }
  if (!requiresGoogle && !state.settings.apiKey) {
    setStatus("请先填写 DashScope / Qwen API Key。", "error");
    switchTab("settings");
    return;
  }

  const candidates = job.pages.filter((page) => page.prepareDone && !page.generated);
  if (!candidates.length) {
    setStatus("还没有可直接批量生成的页面。", "error");
    return;
  }

  setButtonLoading(el.batchGenerateReadyBtn, true, "批量生成中...");
  try {
    for (const page of candidates) {
      const draft = ensurePageDraft(page);
      setStatus(`正在批量生成第 ${page.pageNumber} 页...`, "running");
      const data = await apiJson("/api/workflow/page/generate-v2", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          workflowRouteVersion: 2,
          apiKey: state.settings.apiKey,
          googleApiKey: state.settings.googleApiKey,
          region: state.settings.region,
          imageModel,
          jobId: state.workflowJobId,
          pageId: page.id,
          slideAspect: state.settings.slideAspect,
          size: state.settings.outputSize,
          seed: state.settings.seed,
          extraPrompt: draft.extraPrompt || "",
          canvasImage: "",
        }),
      });
      if (!data.page?.generated) {
        throw new Error(data.page?.generationError || `第 ${page.pageNumber} 页没有拿到图片结果。`);
      }
      state.workflowJob.pages = state.workflowJob.pages.map((item) => item.id === data.page.id ? data.page : item);
      ensureSelectedPage();
      renderPagesWorkbench();
      saveState();
    }
    setStatus("所有可生成页面已完成。", "success");
  } catch (error) {
    setStatus(error.message || "批量生成失败。", "error");
  } finally {
    setButtonLoading(el.batchGenerateReadyBtn, false);
  }
}

function bindEvents() {
  document.querySelectorAll(".sidebar-tab").forEach((button) => {
    button.addEventListener("click", () => switchTab(button.dataset.tab));
  });
  document.querySelectorAll(".ribbon-step").forEach((button) => {
    button.addEventListener("click", () => {
      const next = button.dataset.step;
      if (next === "split" && !state.themeConfirmed) return;
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
  el.confirmThemeBtn.addEventListener("click", confirmTheme);
  el.goSplitBtn.addEventListener("click", () => switchSmartStep("split"));
  el.backToThemeBtn.addEventListener("click", () => switchSmartStep("theme"));
  el.backToSplitBtn.addEventListener("click", () => switchSmartStep("split"));
  el.pickReferenceFilesBtn.addEventListener("click", () => el.referenceFilesInput.click());
  el.referenceFilesInput.addEventListener("change", handleReferenceFiles);
  el.runSplitBtn.addEventListener("click", runSplit);
  el.workflowContent.addEventListener("input", () => {
    state.workflowContent = el.workflowContent.value;
    saveState();
  });
  el.workflowPageCount.addEventListener("change", () => {
    state.workflowPageCount = clamp(Number(el.workflowPageCount.value || 8), 2, 20);
    saveState();
  });
  el.splitTemplateInput.addEventListener("input", () => {
    state.splitTemplateText = el.splitTemplateInput.value;
    saveState();
  });
  el.pageOnscreenEditor.addEventListener("input", () => {
    const page = getSelectedPage();
    const draft = ensurePageDraft(page);
    if (!draft) return;
    draft.onscreenContent = el.pageOnscreenEditor.value;
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
  el.batchGenerateReadyBtn.addEventListener("click", batchGenerateReadyPages);
  el.uploadOverlayBtn.addEventListener("click", () => el.overlayFileInput.click());
  el.overlayFileInput.addEventListener("change", handleOverlayFiles);
  el.clearOverlayBtn.addEventListener("click", clearCurrentOverlays);
  el.generateCurrentPageBtn.addEventListener("click", generateCurrentPage);

  el.reviseImportBtn.addEventListener("click", () => el.reviseFileInput.click());
  el.reviseFileInput.addEventListener("change", handleReviseFiles);
  el.revisePrevBtn.addEventListener("click", () => stepReviseImage(-1));
  el.reviseNextBtn.addEventListener("click", () => stepReviseImage(1));
  el.reviseDeleteBtn.addEventListener("click", deleteCurrentReviseImage);
  el.revisePrompt.addEventListener("input", () => {
    state.revise.prompt = el.revisePrompt.value;
  });
  el.sendReviseBtn.addEventListener("click", sendRevise);

  el.apiKey.addEventListener("input", () => { state.settings.apiKey = el.apiKey.value.trim(); saveState(); });
  el.googleApiKey.addEventListener("input", () => { state.settings.googleApiKey = el.googleApiKey.value.trim(); saveState(); });
  el.workflowImageModel.addEventListener("change", () => { state.settings.workflowImageModel = el.workflowImageModel.value || PPT_MODEL; saveState(); });
  el.region.addEventListener("change", () => { state.settings.region = el.region.value; saveState(); });
  el.slideAspect.addEventListener("change", () => { state.settings.slideAspect = el.slideAspect.value; renderArtboard(); saveState(); });
  el.outputSize.addEventListener("change", () => { state.settings.outputSize = el.outputSize.value; saveState(); });
  el.seed.addEventListener("input", () => { state.settings.seed = el.seed.value.trim(); saveState(); });
  el.testApiKeyBtn.addEventListener("click", testApiKeys);

  window.addEventListener("resize", () => {
    renderArtboard();
    renderRevise();
  });
}

function initialize() {
  cacheElements();
  loadState();
  applyStateToUi();
  renderPreferenceSummary();
  renderSplitPresets();
  renderReferenceFiles();
  updateThemeView();
  ensureSelectedPage();
  renderPagesWorkbench();
  renderRevise();
  switchTab(state.activeTab);
  switchSmartStep(state.smartStep);
  bindEvents();
  setupReviseCanvasInteractions();
  syncWorkflowJobOnce();
  if (state.workflowJobId && state.workflowJob?.status !== "ready") {
    startWorkflowPolling();
  }
}

document.addEventListener("DOMContentLoaded", initialize);
