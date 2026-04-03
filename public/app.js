const STORAGE_KEY = "ppt-image-studio-settings";
const MAX_BOXES_PER_IMAGE = 2;

const SLIDE_ASPECTS = {
  "16:9": { label: "16:9", canvasWidth: 800, canvasHeight: 450, outputWidth: 1600, outputHeight: 900 },
  "4:3": { label: "4:3", canvasWidth: 800, canvasHeight: 600, outputWidth: 1600, outputHeight: 1200 },
  "1:1": { label: "1:1", canvasWidth: 700, canvasHeight: 700, outputWidth: 1400, outputHeight: 1400 },
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

const TAB_DEFAULTS = {
  main: "settings",
  planning: "layout",
  generation: "input",
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
  workflowRunning: false,
  activeTabs: { ...TAB_DEFAULTS },
};

const el = {
  apiKey: document.querySelector("#apiKey"),
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
  workflowPlanBtn: document.querySelector("#workflowPlanBtn"),
  workflowBatchBtn: document.querySelector("#workflowBatchBtn"),
  workflowCopyBtn: document.querySelector("#workflowCopyBtn"),
  workflowClearBtn: document.querySelector("#workflowClearBtn"),
  workflowThemePreview: document.querySelector("#workflowThemePreview"),
  workflowPlanSummary: document.querySelector("#workflowPlanSummary"),
  workflowPlanCards: document.querySelector("#workflowPlanCards"),
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

function getCurrentAspectMeta() {
  return SLIDE_ASPECTS[el.slideAspect.value] || SLIDE_ASPECTS["16:9"];
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

function loadSettings() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}");
    el.apiKey.value = saved.apiKey || "";
    el.requestMode.value = saved.requestMode || "sync";
    el.region.value = saved.region || "singapore";
    el.model.value = saved.model || "wan2.7-image-pro";
    el.sizeMode.value = saved.sizeMode || "preset";
    el.presetSize.value = saved.presetSize || "2K";
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
    el.workflowPageCount.value = saved.workflowPageCount || "4";
    el.workflowTheme.value = saved.workflowTheme || "papercraft";
    el.workflowContent.value = saved.workflowContent || "";
    state.activeTabs = {
      main: saved.activeTabs?.main || TAB_DEFAULTS.main,
      planning: saved.activeTabs?.planning || TAB_DEFAULTS.planning,
      generation: saved.activeTabs?.generation || TAB_DEFAULTS.generation,
    };
    state.palette = Array.isArray(saved.palette) ? saved.palette : [];
    state.slideRegions = Array.isArray(saved.slideRegions) ? saved.slideRegions : [];
    state.editTargetImageId = saved.editTargetImageId || null;
  } catch {
    state.palette = [];
    state.slideRegions = [];
    state.editTargetImageId = null;
    state.activeTabs = { ...TAB_DEFAULTS };
  }
}

function saveSettings() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify({
    apiKey: el.apiKey.value.trim(),
    requestMode: el.requestMode.value,
    region: el.region.value,
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
    workflowPageCount: el.workflowPageCount.value,
    workflowTheme: el.workflowTheme.value,
    workflowContent: el.workflowContent.value,
    activeTabs: state.activeTabs,
  }));
}

function applyTabState() {
  const main = state.activeTabs.main || TAB_DEFAULTS.main;

  document.querySelectorAll("[data-main-tab]").forEach((button) => {
    const active = button.dataset.mainTab === main;
    button.classList.toggle("active", active);
    button.setAttribute("aria-selected", active ? "true" : "false");
  });

  document.querySelectorAll("[data-subtab-button]").forEach((button) => {
    const group = button.dataset.subtabButton;
    const active = state.activeTabs[group] === button.dataset.tabTarget;
    button.classList.toggle("active", active);
    button.setAttribute("aria-selected", active ? "true" : "false");
  });

  document.querySelectorAll("[data-main-panel]").forEach((panel) => {
    const panelMain = panel.dataset.mainPanel;
    let visible = panelMain === main;

    if (visible && panel.dataset.subPanel) {
      const [group, subId] = panel.dataset.subPanel.split(":");
      visible = state.activeTabs[group] === subId;
    }

    panel.hidden = !visible;
  });

  document.querySelectorAll("[data-subtab-nav]").forEach((panel) => {
    panel.hidden = panel.dataset.subtabNav !== main;
  });

  if (main === "planning" && state.activeTabs.planning === "layout") {
    renderSlideCanvas();
  }
}

function bindTabNavigation() {
  document.querySelectorAll("[data-main-tab]").forEach((button) => {
    button.addEventListener("click", () => {
      state.activeTabs.main = button.dataset.mainTab || TAB_DEFAULTS.main;
      if (!state.activeTabs[state.activeTabs.main] && TAB_DEFAULTS[state.activeTabs.main]) {
        state.activeTabs[state.activeTabs.main] = TAB_DEFAULTS[state.activeTabs.main];
      }
      saveSettings();
      applyTabState();
    });
  });

  document.querySelectorAll("[data-subtab-button]").forEach((button) => {
    button.addEventListener("click", () => {
      const group = button.dataset.subtabButton;
      if (!group) return;
      state.activeTabs[group] = button.dataset.tabTarget || TAB_DEFAULTS[group];
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
    ctx.fillStyle = "#132238";
    ctx.fillRect(x, Math.max(0, y - 24), Math.min(width, 140), 22);
    ctx.fillStyle = "#f7f3eb";
    ctx.font = '12px "Segoe UI", sans-serif';
    ctx.fillText(region.label || `区域 ${index + 1}`, x + 8, Math.max(14, y - 8));
  });
  if (state.slideDraft) {
    ctx.strokeStyle = "#32c8ff";
    ctx.setLineDash([6, 4]);
    ctx.strokeRect(state.slideDraft.x, state.slideDraft.y, state.slideDraft.width, state.slideDraft.height);
    ctx.setLineDash([]);
  }
}

function renderSlidePlanner() {
  renderSlideCanvas();
  renderSlideRegionList();
  renderSlideRegionPreview();
  saveSettings();
}

function setupSlideCanvas() {
  const canvas = el.slideCanvas;
  const getCoords = (event) => {
    const rect = canvas.getBoundingClientRect();
    return {
      x: clamp(event.clientX - rect.left, 0, canvas.width),
      y: clamp(event.clientY - rect.top, 0, canvas.height),
    };
  };
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
    state.slideRegions.push({ id: uid(), label: `区域 ${state.slideRegions.length + 1}`, x, y, w, h });
    renderSlidePlanner();
    refreshWorkflowPagePrompts();
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
    const scale = Math.min(1, 320 / preview.naturalWidth);
    canvas.width = Math.round(preview.naturalWidth * scale);
    canvas.height = Math.round(preview.naturalHeight * scale);
    const draw = (draft) => {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(preview, 0, 0, canvas.width, canvas.height);
      ctx.lineWidth = 2;
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
        ctx.fillRect(x1, Math.max(0, y1 - 22), 56, 20);
        ctx.fillStyle = "#f7f3eb";
        ctx.font = '12px "Segoe UI", sans-serif';
        ctx.fillText(`框 ${index + 1}`, x1 + 7, Math.max(13, y1 - 8));
        ctx.fillStyle = "rgba(255, 107, 87, 0.18)";
      });
      if (draft) {
        ctx.strokeStyle = "#32c8ff";
        ctx.setLineDash([6, 4]);
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
      const rect = canvas.getBoundingClientRect();
      const canvasX = clamp(event.clientX - rect.left, 0, canvas.width);
      const canvasY = clamp(event.clientY - rect.top, 0, canvas.height);
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
          <button type="button" class="ghost-btn move-up">上移</button>
          <button type="button" class="ghost-btn move-down">下移</button>
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
    card.querySelector(".move-up").addEventListener("click", () => moveImage(image.id, -1));
    card.querySelector(".move-down").addEventListener("click", () => moveImage(image.id, 1));
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
          <a href="${imageUrl}" target="_blank" rel="noreferrer" class="ghost-btn">打开原图</a>
          <button type="button" class="ghost-btn use-target-btn">作为改图底图</button>
          <button type="button" class="ghost-btn save-local-btn">下载到本地</button>
          ${saved ? `<a href="${saved.localUrl}" target="_blank" rel="noreferrer" class="ghost-btn">打开本地文件</a>` : ""}
        </div>
      </div>
    `;
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
  state.workflowPages.forEach((page) => {
    const promptChanged = page.pagePrompt && page.pagePrompt !== buildWorkflowPagePrompt(page);
    page.pagePrompt = buildWorkflowPagePrompt(page);
    if (promptChanged) {
      page.status = "idle";
      page.error = "提示词已更新，请重新生成本页。";
      page.resultImages = [];
      page.savedResults = {};
      page.requestId = "";
      page.taskId = "";
    }
  });
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
        pagePrompt: "",
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
  if (!response.ok) throw new Error(data.message || "查询任务失败。");
  return data;
}

async function requestGeneration(payload) {
  const response = await fetch("/api/generate", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      apiKey: el.apiKey.value.trim(),
      region: el.region.value,
      asyncMode: el.requestMode.value === "async",
      payload,
    }),
  });
  const data = await response.json();
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
  state.workflowRunning = false;
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
    page.status = "success";
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

init();
