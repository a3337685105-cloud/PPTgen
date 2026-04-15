function normalizeDensityBandValue(value) {
  const normalized = String(value || "").trim().toLowerCase();
  return ["lite", "concise", "standard", "dense"].includes(normalized) ? normalized : "";
}

function normalizeLayoutRiskValue(value) {
  const normalized = String(value || "").trim().toLowerCase();
  return ["low", "medium", "high"].includes(normalized) ? normalized : "";
}

function getPptTextLength(value) {
  return String(value || "").replace(/\s+/g, "").length;
}

function inferWorkflowDensityBand(pageOrText) {
  const source = typeof pageOrText === "string"
    ? pageOrText
    : [pageOrText?.pageTitle, pageOrText?.pageContent].filter(Boolean).join(" ");
  const length = getPptTextLength(source);
  if (length < 50) return "lite";
  if (length <= 150) return "concise";
  if (length <= 250) return "standard";
  return "dense";
}

function inferWorkflowLayoutRisk(page) {
  const densityBand = normalizeDensityBandValue(page?.densityBand) || inferWorkflowDensityBand(page);
  if (densityBand === "dense") return "high";
  if (densityBand === "standard") {
    return page?.pageType === "data" || getPptTextLength(page?.pageContent) > 220 ? "medium" : "low";
  }
  if (densityBand === "concise") {
    return page?.pageType === "data" ? "medium" : "low";
  }
  return "low";
}

function applyHarnessMetaToPage(page) {
  if (!page || typeof page !== "object") return page;
  page.densityBand = normalizeDensityBandValue(page.densityBand) || inferWorkflowDensityBand(page);
  page.layoutRisk = normalizeLayoutRiskValue(page.layoutRisk) || inferWorkflowLayoutRisk(page);
  return page;
}

function sanitizeWorkflowPageFields(page) {
  if (!page || typeof page !== "object") return page;
  page.decorationLevel = getPageDecorationLevel(page);
  page.layoutSummary = stringifyStructuredField(page.layoutSummary);
  page.textHierarchy = stringifyStructuredField(page.textHierarchy);
  page.visualFocus = stringifyStructuredField(page.visualFocus);
  page.readabilityNotes = stringifyStructuredField(page.readabilityNotes);
  if (!Object.prototype.hasOwnProperty.call(page, "displayTitleOverride")) page.displayTitleOverride = null;
  if (!Object.prototype.hasOwnProperty.call(page, "displayBodyOverride")) page.displayBodyOverride = null;
  if (!Object.prototype.hasOwnProperty.call(page, "researchStatus")) page.researchStatus = "idle";
  if (!Object.prototype.hasOwnProperty.call(page, "researchSummary")) page.researchSummary = "";
  if (!Object.prototype.hasOwnProperty.call(page, "researchCandidates")) page.researchCandidates = [];
  if (!Object.prototype.hasOwnProperty.call(page, "researchRaw")) page.researchRaw = "";
  if (!Object.prototype.hasOwnProperty.call(page, "researchQueryOverride")) page.researchQueryOverride = "";
  if (!Object.prototype.hasOwnProperty.call(page, "researchLastQuery")) page.researchLastQuery = "";
  if (promptLooksTextless(page.pagePrompt, page) && hasConfirmedThemeDefinition()) {
    page.pagePrompt = buildWorkflowPagePrompt(page);
  }
  return page;
}

function getDensityBandGuide(band) {
  switch (normalizeDensityBandValue(band)) {
    case "lite":
      return "极少文字，优先极简海报式构图，只保留一个强标题或一个核心结论。";
    case "concise":
      return "中低密度信息，优先大标题加 2 到 3 个清晰信息块。";
    case "standard":
      return "标准信息量，优先稳定的双栏、分屏或模块化图文布局。";
    case "dense":
      return "高密度信息，必须优先分栏、多块、图表化或延续建议，禁止压成整页小字。";
    default:
      return "";
  }
}

function getLayoutRiskGuide(risk) {
  switch (normalizeLayoutRiskValue(risk)) {
    case "high":
      return "这页很容易做成小字墙，必须优先牺牲装饰性来换可读性。";
    case "medium":
      return "这页需要明显层级和节奏，防止重点被正文或图表淹没。";
    case "low":
      return "这页可以保留更多视觉张力，但仍要遵守标题与正文层级。";
    default:
      return "";
  }
}

function isPptHarnessEnabled() {
  return !el.pptHarnessEnabled || el.pptHarnessEnabled.checked !== false;
}

function getPptHarnessPack() {
  return state.pptHarnessReady && state.pptHarnessPack && Array.isArray(state.pptHarnessPack.chunks)
    ? state.pptHarnessPack
    : null;
}

async function loadPptHarness() {
  try {
    const response = await fetch("/ppt-harness.json");
    if (!response.ok) throw new Error("PPT Harness 资源加载失败。");
    const pack = await response.json();
    if (!pack || !Array.isArray(pack.chunks)) throw new Error("PPT Harness 资源格式无效。");
    state.pptHarnessPack = pack;
    state.pptHarnessReady = true;
    state.pptHarnessLoadError = "";
    return pack;
  } catch (error) {
    state.pptHarnessPack = null;
    state.pptHarnessReady = false;
    state.pptHarnessLoadError = error.message || "PPT Harness 已回退到旧规则。";
    return null;
  }
}

function getPptHarnessChunkById(id) {
  const pack = getPptHarnessPack();
  if (!pack) return null;
  return pack.chunks.find((chunk) => chunk.id === id) || null;
}

function getPptHarnessChunks(appliesTo, chunkIds = []) {
  const pack = getPptHarnessPack();
  if (!pack) return [];
  if (Array.isArray(chunkIds) && chunkIds.length) {
    return chunkIds.map((id) => getPptHarnessChunkById(id)).filter(Boolean);
  }
  return pack.chunks.filter((chunk) => Array.isArray(chunk.appliesTo) && chunk.appliesTo.includes(appliesTo));
}

function inferHarnessStyleProfile(seed) {
  const pack = getPptHarnessPack();
  if (!pack?.styleProfiles) return null;
  const text = typeof seed === "string"
    ? seed
    : [seed?.label, seed?.basic, seed?.cover, seed?.content, seed?.data].filter(Boolean).join(" ");
  const normalized = text.toLowerCase();

  if (/学术|研究|论文|答辩|综述|教育|课程|知识库|报告/.test(normalized)) return pack.styleProfiles.academic || null;
  if (/创意|艺术|品牌|活动|海报|国风|纸艺|插画|动漫|潮流|时尚/.test(normalized)) return pack.styleProfiles.creative || null;
  if (/科技|商业|商务|企业|汇报|路演|战略|方案|金融|产品|行业|平台|医疗|制造/.test(normalized)) return pack.styleProfiles.business || null;
  return pack.styleProfiles.general || null;
}

function buildHarnessPromptBlock(appliesTo, options = {}) {
  if (!isPptHarnessEnabled()) return "";
  const chunks = getPptHarnessChunks(appliesTo, options.chunkIds);
  if (!chunks.length) return "";

  const lines = ["【PPT Harness】以下规则来自内置 PPT 排版指南，优先级高于纯风格化表达："];
  chunks.forEach((chunk) => {
    if (chunk?.promptSnippet) lines.push(`${chunk.title}：${chunk.promptSnippet}`);
    if (options.includeHardRules) {
      (chunk.hardRules || []).slice(0, options.maxHardRules || 1).forEach((rule) => lines.push(`- ${rule}`));
    }
  });

  const profile = inferHarnessStyleProfile(options.styleSeed || options.theme || getThemeName());
  if (profile?.promptSnippet) {
    lines.push(`风格适配：${String(profile.promptSnippet).trim()}`);
    if (options.includeHardRules) {
      (profile.hardRules || []).slice(0, 1).forEach((rule) => lines.push(`- ${rule}`));
    }
  }

  if (options.densityBand) lines.push(`页面密度档位：${options.densityBand}。${getDensityBandGuide(options.densityBand)}`);
  if (options.layoutRisk) lines.push(`版式风险：${options.layoutRisk}。${getLayoutRiskGuide(options.layoutRisk)}`);
  return lines.join("\n");
}

function getHarnessStatusLabel() {
  if (!isPptHarnessEnabled()) return "Harness 关";
  return state.pptHarnessReady ? "Harness 开" : "Harness 回退";
}

function stringifyStructuredField(value, depth = 0) {
  if (value === undefined || value === null) return "";
  if (typeof value === "string") {
    const normalized = value.trim();
    return normalized === "[object Object]" ? "" : normalized;
  }
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  if (Array.isArray(value)) {
    return value
      .map((item) => stringifyStructuredField(item, depth + 1))
      .filter(Boolean)
      .map((item) => (depth === 0 ? `- ${item}` : item))
      .join("\n");
  }
  if (typeof value === "object") {
    return Object.entries(value)
      .map(([key, entry]) => {
        const next = stringifyStructuredField(entry, depth + 1);
        return next ? `${key}：${next.replace(/\n+/g, "；")}` : "";
      })
      .filter(Boolean)
      .join("\n");
  }
  return String(value).trim();
}

function getWorkflowRequiredTextLines(page) {
  const source = String(page?.pageContent || "").replace(/\r/g, "").trim();
  if (!source) return [];
  const lines = [];
  source
    .split(/\n+/)
    .map((line) => line.trim())
    .filter(Boolean)
    .forEach((line) => {
      if (!lines.includes(line)) lines.push(line);
    });
  return lines;
}

function hasWorkflowTitleOverride(page) {
  return page && Object.prototype.hasOwnProperty.call(page, "displayTitleOverride") && page.displayTitleOverride !== null && page.displayTitleOverride !== undefined;
}

function hasWorkflowBodyOverride(page) {
  return page && Object.prototype.hasOwnProperty.call(page, "displayBodyOverride") && page.displayBodyOverride !== null && page.displayBodyOverride !== undefined;
}

function getWorkflowDisplayTitle(page) {
  return hasWorkflowTitleOverride(page) ? String(page.displayTitleOverride || "").trim() : String(page?.pageTitle || "").trim();
}

function parseWorkflowVisibleBodyText(value) {
  return String(value || "")
    .replace(/\r/g, "")
    .split(/\n+/)
    .map((line) => line.trim())
    .filter(Boolean);
}

function isWorkflowPriorityLine(line, page) {
  const text = String(line || "").trim();
  if (!text) return false;
  if (page?.pageType === "cover") return true;
  if (/^【.+】$/.test(text)) return true;
  if (/^[0-9０-９]+[\.、．]/.test(text)) return true;
  if (/^[\-•●▪◦→]/.test(text)) return true;
  if (/：|:/.test(text)) return true;
  if (/[0-9０-９%％$¥￥℃°VvWwkK亿万千百篇次年]/.test(text)) return true;
  if (/成员|作者|署名|机构|来源|参考|文献|学校|大学|研究所|实验室|团队|分工|案例|场景|趋势|挑战|结论|摘要|定义|功能|参数|指标/.test(text)) return true;
  if (/^[A-Z][A-Z0-9\s\-]{1,30}$/.test(text)) return true;
  if (text.length <= 18) return true;
  return false;
}

function splitWorkflowTextBlocks(page) {
  const source = String(page?.pageContent || "").replace(/\r/g, "").trim();
  if (!source) return [];
  const normalized = source
    .replace(/([。！？；])/g, "$1\n")
    .replace(/([：:])\s*/g, "$1 ")
    .split(/\n+/)
    .map((line) => line.trim())
    .filter(Boolean);
  const blocks = [];
  normalized.forEach((line) => {
    const segments = line
      .split(/(?<=。|；|！|？)\s+/)
      .map((item) => item.trim())
      .filter(Boolean);
    segments.forEach((segment) => {
      if (!blocks.includes(segment)) blocks.push(segment);
    });
  });
  return blocks;
}

function getWorkflowVisibleLineLimit(page) {
  if (page?.pageType === "cover") return 12;
  switch (normalizeDensityBandValue(page?.densityBand)) {
    case "lite":
      return 6;
    case "concise":
      return 7;
    case "standard":
      return 8;
    case "dense":
      return 10;
    default:
      return 8;
  }
}

function buildWorkflowVisibleTextPlan(page) {
  const title = getWorkflowDisplayTitle(page);
  const rawLines = getWorkflowRequiredTextLines(page);
  if (hasWorkflowBodyOverride(page)) {
    const visibleLines = parseWorkflowVisibleBodyText(page.displayBodyOverride);
    return {
      title,
      rawLines,
      priorityLines: [...visibleLines],
      secondaryLines: [],
      visibleLines,
      omittedLines: rawLines.filter((line) => !visibleLines.includes(line)),
      limit: visibleLines.length || getWorkflowVisibleLineLimit(page),
      isManualOverride: true,
    };
  }
  const limit = getWorkflowVisibleLineLimit(page);
  const priorityLines = [];
  const secondaryLines = [];

  rawLines.forEach((line) => {
    if (isWorkflowPriorityLine(line, page)) {
      priorityLines.push(line);
    } else {
      secondaryLines.push(line);
    }
  });

  const visibleLines = [];
  priorityLines.forEach((line) => {
    if (!visibleLines.includes(line)) visibleLines.push(line);
  });
  secondaryLines.forEach((line) => {
    if (visibleLines.length < limit && !visibleLines.includes(line)) visibleLines.push(line);
  });

  if (!visibleLines.length && rawLines.length) {
    rawLines.slice(0, limit).forEach((line) => visibleLines.push(line));
  }

  const omittedLines = rawLines.filter((line) => !visibleLines.includes(line));
  return {
    title,
    rawLines,
    priorityLines,
    secondaryLines,
    visibleLines,
    omittedLines,
    limit,
    isManualOverride: false,
  };
}

function inferWorkflowSurfaceMode(page, plan = null) {
  const nextPlan = plan || buildWorkflowVisibleTextPlan(page);
  const title = String(nextPlan.title || page?.pageTitle || "").trim();
  const lines = nextPlan.visibleLines;
  const numberedCount = lines.filter((line) => /^(?:第?\s*[0-9０-９]{1,2}|[0-9０-９]{1,2})[\s、．.]?/.test(line)).length;
  const colonCount = lines.filter((line) => /[:：]/.test(line)).length;

  if (/^(目录|contents?)$/i.test(title) || numberedCount >= 2) return "toc";
  if (/成员|分工|团队|小组|作者|研究人员/.test(title) || colonCount >= 2) return "members";
  if (numberedCount + colonCount >= 2) return "structured";
  return "default";
}

function buildWorkflowSurfaceStrategy(page, plan = null) {
  const nextPlan = plan || buildWorkflowVisibleTextPlan(page);
  const mode = inferWorkflowSurfaceMode(page, nextPlan);
  if (mode === "toc") {
    return "本页是目录/章节索引页。编号开头的行视为目录项标题，紧随其后的非编号行视为该目录项说明；优先做成 2 到 4 个目录卡片或模块，不要把整页做成长段正文。";
  }
  if (mode === "members") {
    return "本页是成员/分工页。每一条含冒号的文本都应作为独立成员卡主文案逐字落图，不得根据姓名或角色自行补写职责介绍、扩展说明或总结句。";
  }
  if (mode === "structured") {
    return "本页属于结构化信息页。请把每一行视作独立信息块优先排成卡片、列表或分栏，不要把它们改写成长段说明。";
  }
  return "";
}

function inferWorkflowHierarchyLevel(line, page, context = {}) {
  const text = String(line || "").trim();
  if (!text) return 0;
  const mode = context.mode || inferWorkflowSurfaceMode(page, context.plan || null);

  if (mode === "members" && /[:：]/.test(text)) return 2;
  if (/^(?:【[^】]+】|CHAPTER\b|PART\b|SECTION\b|(?:第?\s*[一二三四五六七八九十0-9]+\s*[章节部分页])|(?:0?\d{1,2}|[A-Z])[、.．)\s]|[一二三四五六七八九十]+[、.．])/.test(text)) {
    return 2;
  }
  if (/^[\-•▪●▶✓⚠🔹👤🏢🌍🏥🏭📚💼🎨]/.test(text)) return context.hasPrimary ? 3 : 2;
  if (/[:：]/.test(text) && text.length <= 32) return mode === "members" ? 2 : (context.hasPrimary ? 3 : 2);
  if (text.length <= 18 && !/[。；]/.test(text)) return context.hasPrimary ? 3 : 2;
  if (/[。；]/.test(text) || text.length > 32) return context.hasSecondary ? 4 : (context.hasPrimary ? 3 : 2);
  return context.hasPrimary ? 3 : 2;
}

function buildWorkflowHierarchyOutline(page, plan = null) {
  const nextPlan = plan || buildWorkflowVisibleTextPlan(page);
  const mode = inferWorkflowSurfaceMode(page, nextPlan);
  const title = String(nextPlan.title || page?.pageTitle || "").trim();
  const lines = [];
  if (title) lines.push(`L1 页标题：${title}`);

  let hasPrimary = false;
  let hasSecondary = false;
  nextPlan.visibleLines.forEach((line) => {
    let level = inferWorkflowHierarchyLevel(line, page, { mode, plan: nextPlan, hasPrimary, hasSecondary });
    if (!level) return;
    if (level > 4) level = 4;
    if (level === 2) {
      hasPrimary = true;
      hasSecondary = false;
    } else if (level === 3) {
      hasSecondary = true;
    } else if (level === 4 && !hasSecondary) {
      level = 3;
      hasSecondary = true;
    }
    const indent = "  ".repeat(Math.max(level - 2, 0));
    lines.push(`${indent}L${level} ${String(line).trim()}`);
  });

  if (!lines.length && title) lines.push(`L1 页标题：${title}`);
  return lines.join("\n");
}

function buildWorkflowVisibleTextPreview(page) {
  const plan = buildWorkflowVisibleTextPlan(page);
  const lines = [];
  if (plan.title) lines.push(`标题：${plan.title}`);
  if (plan.visibleLines.length) {
    lines.push("正文：");
    plan.visibleLines.forEach((line) => lines.push(line));
  } else if (plan.title) {
    lines.push("正文：本页仅保留标题。");
  } else {
    lines.push("正文：暂无可上屏文本。");
  }
  return lines.join("\n");
}

function normalizeWorkflowResearchCandidates(candidates) {
  if (!Array.isArray(candidates)) return [];
  return candidates
    .map((candidate) => ({
      text: String(candidate?.text || "").trim(),
      why: String(candidate?.why || candidate?.reason || "").trim(),
      sources: Array.isArray(candidate?.sources)
        ? candidate.sources
          .map((source) => ({
            title: String(source?.title || "").trim(),
            url: String(source?.url || "").trim(),
          }))
          .filter((source) => source.title && source.url)
        : [],
      selected: Boolean(candidate?.selected),
    }))
    .filter((candidate) => candidate.text && candidate.sources.length);
}

function stringifyResearchRaw(raw) {
  if (typeof raw === "string") return raw.trim();
  if (!raw || typeof raw !== "object") return "";
  try {
    return JSON.stringify(raw, null, 2);
  } catch {
    return String(raw);
  }
}

function buildWorkflowTextPolicySummary(page) {
  const plan = buildWorkflowVisibleTextPlan(page);
  if (plan.isManualOverride) {
    return `本页使用手工确认的最终上屏内容，共 ${plan.visibleLines.length} 行。`;
  }
  const kept = plan.visibleLines.length;
  const omitted = plan.omittedLines.length;
  if (page?.pageType === "cover") {
    return omitted ? `封面页已保留标题和 ${kept} 行正文，压缩 ${omitted} 行非核心文本。` : `封面页已保留标题和 ${kept} 行正文。`;
  }
  return omitted ? `本页将保留 ${kept} 行核心文本，压缩 ${omitted} 行次级解释。` : `本页将保留全部 ${kept} 行文本。`;
}

function buildWorkflowTextPlacementBlock(page) {
  const plan = buildWorkflowVisibleTextPlan(page);
  const title = plan.title || String(page?.pageTitle || `第 ${page?.pageNumber || ""} 页`).trim();
  const isCover = page?.pageType === "cover";
  const surfaceStrategy = buildWorkflowSurfaceStrategy(page, plan);
  const hierarchyOutline = buildWorkflowHierarchyOutline(page, plan);
  const lines = [
    "文字落图硬约束（优先级高于前面的风格和版式描述）：",
    "1. 这不是纯背景图，必须把下面的中文标题和正文真实排进画面里。",
    "2. 严禁输出“无文字”“只保留主视觉”“后续再加字”“纯背景图”“预留文字区域”等做法。",
    "3. 标题必须作为主标题直接出现在画面中，正文必须以清晰的大字号中文排版呈现。",
    "4. 若正文较多，优先拆成 2 到 4 个文本块、卡片或分栏；不要删掉关键信息，不要缩成密集小字墙。",
    "5. pageContent 中出现的成员名单、署名、机构、年份、参考信息、单位和关键数字都属于必保留信息，不能遗漏。",
    "6. 下方白名单中的每一行都必须逐字使用；禁止同义改写、扩写、释义、补全、润色，禁止额外生成“本页文字：”“说明：”“小结：”这类白名单外标签。",
    "7. 只能使用下面白名单里的文本，禁止新增任何用户未提供的机构名、作者名、年份、英文副标题、页脚说明、来源、数字或占位词。",
    "8. 允许增加非文字的视觉补充，例如图标、结构示意、时间线、工艺剖面、场景缩略图、箭头关系、容器模块、背景几何和材质细节，但这些补充不得包含新的可读文字、标签或数字。",
    buildNoExtraTextConstraintBlock({ requireWhitelist: true }),
    `主标题：${title}`,
  ];
  if (surfaceStrategy) lines.push(`结构化排版策略：${surfaceStrategy}`);
  if (hierarchyOutline) {
    lines.push("内容层级提纲：");
    lines.push(hierarchyOutline);
  }
  if (isCover) {
    lines.push("封面页特殊要求：如果封面包含副标题、研究成员、机构或补充说明，它们都必须清晰可辨地出现在画面中，不能降成装饰性纹理、模糊小字或被省略。");
  }
  if (plan.isManualOverride) {
    lines.push("当前页已由用户手工确认最终上屏内容，下面白名单就是唯一允许出现的文字；白名单以外的原文全部不要显示。");
  }
  if (plan.visibleLines.length) {
    lines.push("允许上屏文本白名单：");
    plan.visibleLines.forEach((block) => lines.push(`- 「${block}」`));
  } else if (page?.pageContent) {
    lines.push(`正文文本块：\n- 「${String(page.pageContent).trim()}」`);
  }
  if (!plan.isManualOverride && plan.omittedLines.length) {
    lines.push("可压缩但不可改写的原文：");
    plan.omittedLines.forEach((line, index) => lines.push(`${index + 1}. ${line}`));
  }
  lines.push(buildWorkflowTextPolicySummary(page));
  lines.push("排版提醒：标题显著放大，正文保证远距离可读，关键数字和结论优先放大。你可以用更丰富的无字视觉元素把内容做得更像信息图，但如果白名单里没有某类文字信息，就不要擅自补出那类信息。");
  return lines.join("\n");
}

function promptLooksTextless(prompt, page) {
  const normalized = String(prompt || "").trim();
  if (!normalized) return true;
  if (/\b无文字\b|不要文字|仅保留主视觉|只保留主视觉|纯背景图|背景图|后续添加|后续再加|预留.*文字|仅保留视觉元素/i.test(normalized)) {
    return true;
  }
  const title = String(page?.pageTitle || "").trim();
  if (title && !normalized.includes(title)) return true;
  const plan = buildWorkflowVisibleTextPlan(page);
  const mustIncludeLines = page?.pageType === "cover"
    ? plan.visibleLines
    : plan.visibleLines.slice(0, Math.min(plan.visibleLines.length, 6));
  if (mustIncludeLines.length && !mustIncludeLines.every((line) => normalized.includes(line.slice(0, Math.min(line.length, 10))))) {
    return true;
  }
  if (!plan.rawLines.some((line) => /研究所|大学|实验室|学院|公司|集团|202[0-9]|20[0-9]{2}/.test(line))
    && /研究所|大学|实验室|学院|公司|集团|202[0-9]|20[0-9]{2}/.test(normalized)) {
    return true;
  }
  return false;
}

function buildEffectiveWorkflowPagePrompt(page) {
  const rawPrompt = String(page?.pagePrompt || "").trim();
  const basePrompt = promptLooksTextless(rawPrompt, page)
    ? (hasConfirmedThemeDefinition() ? buildWorkflowPagePrompt(page) : rawPrompt)
    : rawPrompt;
  return [basePrompt, buildWorkflowTextPlacementBlock(page)].filter(Boolean).join("\n\n");
}

function buildSettingsSnapshot() {
  return {
    themeWorkflowVersion: THEME_WORKFLOW_VERSION,
    apiKey: el.apiKey.value.trim(),
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
    pptHarnessEnabled: isPptHarnessEnabled(),
    workflowPageCount: el.workflowPageCount.value,
    workflowTheme: el.workflowTheme.value || el.themeName.value.trim(),
    workflowContent: el.workflowContent.value,
    workflowPlanSummary: state.workflowPlanSummary,
    workflowPages: state.workflowPages.map((page) => serializeWorkflowPage(page)),
    workflowPlanLibrary: state.workflowPlanLibrary.map((entry) => ({ ...entry, pages: entry.pages.map((page) => serializeWorkflowPage(page)) })),
    workflowPlanLibraryActiveId: state.workflowPlanLibraryActiveId,
    activeTabs: state.activeTabs,
  };
}

function buildLibraryDoc() {
  const settings = buildSettingsSnapshot();
  return {
    version: LIBRARY_DOC_VERSION,
    savedAt: new Date().toISOString(),
    settings: {
      apiKey: settings.apiKey,
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
      pptHarnessEnabled: settings.pptHarnessEnabled,
      workflowPlanLibraryActiveId: settings.workflowPlanLibraryActiveId,
    },
    themes: state.themeLibrary.map((entry) => ({ ...entry, definition: entry.definition })),
    workflowPlans: state.workflowPlanLibrary.map((entry) => ({ ...entry, pages: entry.pages.map((page) => serializeWorkflowPage(page)) })),
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
  assignChecked(el.pptHarnessEnabled, settings.pptHarnessEnabled);

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
  if (typeof settings.themeDefinitionRaw === "string") state.themeDefinitionRaw = settings.themeDefinitionRaw;
  if (typeof settings.themeDefinitionSource === "string") state.themeDefinitionSource = settings.themeDefinitionSource;
  if (settings.themeConfirmed !== undefined) state.themeConfirmed = Boolean(settings.themeConfirmed);
  if (typeof settings.themeConfirmedSource === "string") state.themeConfirmedSource = settings.themeConfirmedSource;
  if (typeof settings.workflowPlanLibraryActiveId === "string") state.workflowPlanLibraryActiveId = settings.workflowPlanLibraryActiveId;
}

function serializeWorkflowPage(page) {
  const next = applyHarnessMetaToPage({ ...page });
  return {
    pageNumber: next.pageNumber,
    pageType: next.pageType,
    pageTitle: next.pageTitle,
    pageContent: next.pageContent,
    densityBand: next.densityBand,
    layoutRisk: next.layoutRisk,
    layoutSummary: stringifyStructuredField(next.layoutSummary),
    textHierarchy: stringifyStructuredField(next.textHierarchy),
    visualFocus: stringifyStructuredField(next.visualFocus),
    readabilityNotes: stringifyStructuredField(next.readabilityNotes),
    pagePrompt: next.pagePrompt || "",
    displayTitleOverride: next.displayTitleOverride,
    displayBodyOverride: next.displayBodyOverride,
    researchStatus: next.researchStatus || "idle",
    researchSummary: next.researchSummary || "",
    researchCandidates: normalizeWorkflowResearchCandidates(next.researchCandidates),
    researchRaw: next.researchRaw || "",
    researchQueryOverride: next.researchQueryOverride || "",
    researchLastQuery: next.researchLastQuery || "",
    resultImages: Array.isArray(next.resultImages) ? next.resultImages.filter(Boolean) : [],
    savedResults: next.savedResults && typeof next.savedResults === "object" ? next.savedResults : {},
    requestId: next.requestId || "",
    taskId: next.taskId || "",
    detailBackdropUrl: next.detailBackdropUrl || "",
  };
}

function hydrateWorkflowPages(savedPages) {
  if (!Array.isArray(savedPages)) return [];
  return savedPages
    .map((page, index) => applyHarnessMetaToPage({
      id: uid(),
      pageNumber: Number(page.pageNumber ?? page.page_number ?? index + 1) || index + 1,
      pageType: normalizeWorkflowPageType(page.pageType ?? page.page_type, index),
      pageTitle: String(page.pageTitle ?? page.page_title ?? `第 ${index + 1} 页`).trim(),
      pageContent: String(page.pageContent ?? page.page_content ?? page.pageTitle ?? "").trim(),
      densityBand: page.densityBand ?? page.density_band,
      layoutRisk: page.layoutRisk ?? page.layout_risk,
      layoutSummary: stringifyStructuredField(page.layoutSummary || page.layout_summary),
      textHierarchy: stringifyStructuredField(page.textHierarchy || page.text_hierarchy),
      visualFocus: stringifyStructuredField(page.visualFocus || page.visual_focus),
      readabilityNotes: stringifyStructuredField(page.readabilityNotes || page.readability_notes),
      pagePrompt: String(page.pagePrompt || page.page_prompt || "").trim(),
      displayTitleOverride: Object.prototype.hasOwnProperty.call(page, "displayTitleOverride") ? page.displayTitleOverride : (Object.prototype.hasOwnProperty.call(page, "display_title_override") ? page.display_title_override : null),
      displayBodyOverride: Object.prototype.hasOwnProperty.call(page, "displayBodyOverride") ? page.displayBodyOverride : (Object.prototype.hasOwnProperty.call(page, "display_body_override") ? page.display_body_override : null),
      researchStatus: String(page.researchStatus || page.research_status || "idle").trim() || "idle",
      researchSummary: String(page.researchSummary || page.research_summary || "").trim(),
      researchCandidates: normalizeWorkflowResearchCandidates(page.researchCandidates || page.research_candidates),
      researchRaw: String(page.researchRaw || page.research_raw || "").trim(),
      researchQueryOverride: String(page.researchQueryOverride || page.research_query_override || "").trim(),
      researchLastQuery: String(page.researchLastQuery || page.research_last_query || "").trim(),
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
    }))
    .filter((page) => page.pageTitle || page.pageContent)
    .sort((a, b) => a.pageNumber - b.pageNumber);
}

function normalizeWorkflowPlan(parsed) {
  const rawPages = Array.isArray(parsed?.pagePlan) ? parsed.pagePlan : Array.isArray(parsed?.pages) ? parsed.pages : Array.isArray(parsed) ? parsed : [];
  const pages = rawPages
    .map((page, index) => applyHarnessMetaToPage({
      id: uid(),
      pageNumber: Number(page.pageNumber ?? page.page_number ?? index + 1) || index + 1,
      pageType: normalizeWorkflowPageType(page.pageType ?? page.page_type, index),
      pageTitle: String(page.pageTitle ?? page.page_title ?? (index === 0 ? page.pageContent ?? page.page_content ?? "" : `第 ${index + 1} 页`)).trim() || `第 ${index + 1} 页`,
      pageContent: String(page.pageContent ?? page.page_content ?? page.pageTitle ?? page.page_title ?? "").trim(),
      densityBand: page.densityBand ?? page.density_band,
      layoutRisk: page.layoutRisk ?? page.layout_risk,
      layoutSummary: stringifyStructuredField(page.layoutSummary || page.layout_summary),
      textHierarchy: stringifyStructuredField(page.textHierarchy || page.text_hierarchy),
      visualFocus: stringifyStructuredField(page.visualFocus || page.visual_focus),
      readabilityNotes: stringifyStructuredField(page.readabilityNotes || page.readability_notes),
      pagePrompt: "",
      displayTitleOverride: Object.prototype.hasOwnProperty.call(page, "displayTitleOverride") ? page.displayTitleOverride : (Object.prototype.hasOwnProperty.call(page, "display_title_override") ? page.display_title_override : null),
      displayBodyOverride: Object.prototype.hasOwnProperty.call(page, "displayBodyOverride") ? page.displayBodyOverride : (Object.prototype.hasOwnProperty.call(page, "display_body_override") ? page.display_body_override : null),
      researchStatus: String(page.researchStatus || page.research_status || "idle").trim() || "idle",
      researchSummary: String(page.researchSummary || page.research_summary || "").trim(),
      researchCandidates: normalizeWorkflowResearchCandidates(page.researchCandidates || page.research_candidates),
      researchRaw: String(page.researchRaw || page.research_raw || "").trim(),
      researchQueryOverride: String(page.researchQueryOverride || page.research_query_override || "").trim(),
      researchLastQuery: String(page.researchLastQuery || page.research_last_query || "").trim(),
      layoutStatus: "idle",
      layoutError: "",
      layoutPromise: null,
      status: "idle",
      error: "",
      resultImages: [],
      savedResults: {},
      requestId: "",
      taskId: "",
    }))
    .filter((page) => page.pageTitle || page.pageContent)
    .sort((a, b) => a.pageNumber - b.pageNumber)
    .map((page, index) => applyHarnessMetaToPage({ ...page, pageNumber: index + 1, pageType: index === 0 ? "cover" : page.pageType }));

  pages.forEach((page) => { page.pagePrompt = buildWorkflowPagePrompt(page); });
  return {
    summary: String(parsed?.summary || parsed?.planSummary || "").trim(),
    pages,
  };
}

function getThemeAgentSystemPrompt(themeName = getThemeName()) {
  const harnessBlock = buildHarnessPromptBlock("theme", {
    chunkIds: ["core_principles", "readability_rules", "typography_system", "layout_system", "quality_checks"],
    includeHardRules: true,
    maxHardRules: 1,
    styleSeed: themeName,
  });
  return [
    "# Role",
    "你是一位世界顶级的视觉艺术总监、Prompt Engineer 与 UI/UX 专家，擅长把抽象风格主题扩展成稳定、可读、可投影的 PPT 视觉语言系统。",
    "",
    "# Task",
    "请根据用户提供的目标风格主题，输出一份高质量的演示文稿视觉风格定义 JSON。",
    "",
    "# Constraints",
    "1. 必须把简单风格词扩展成完整视觉系统：风格融合、光影材质、配色方案、网格容器、视觉锚点、渲染质量都要明确。",
    "2. basic 中必须融合 3 种相关高级风格，并明确底色、4 色点缀色谱、材质、光照、网格系统、3D 主视觉和渲染审美。",
    "3. cover 必须强制采用极简海报模式，屏蔽全局 UI 网格；背景保留大面积留白，只保留一个核心 3D 主视觉物体。",
    "4. content 必须强调基于网格系统的信息布局、容器材质、边缘、阴影和留白。",
    "5. data 必须把图表做物体化转译，不允许平面图表直出。",
    "6. basic、cover、content、data 必须全部使用简体中文输出。",
    harnessBlock,
    "",
    "# Output",
    "只输出 JSON object，不要解释，不要 markdown，不要代码块。",
    "{\"basic\":\"...\",\"cover\":\"...\",\"content\":\"...\",\"data\":\"...\"}",
  ].filter(Boolean).join("\n");
}

function buildThemeDefinitionPayload(themeName) {
  return {
    model: "qwen3.6-plus",
    input: {
      messages: [
        { role: "system", content: getThemeAgentSystemPrompt(themeName) },
        { role: "user", content: [{ text: `现在，请生成“${themeName}”的提示词模板，并输出对应 JSON，不要包含任何解释和说明性内容。` }] },
      ],
    },
    parameters: { result_format: "message", response_format: { type: "json_object" }, enable_thinking: el.assistantThinking.checked },
  };
}

function buildWorkflowPagePrompt(page) {
  const theme = requireConfirmedTheme();
  const aspect = getCurrentAspectMeta();
  const next = applyHarnessMetaToPage({ ...page });
  const displayTitle = getWorkflowDisplayTitle(next);
  const visiblePlan = buildWorkflowVisibleTextPlan(next);
  const harnessBlock = buildHarnessPromptBlock("pageDesign", {
    chunkIds: ["readability_rules", "simplicity_rules", "layout_system", "page_density_strategies", "content_element_rules", "quality_checks"],
    includeHardRules: true,
    maxHardRules: 1,
    styleSeed: theme,
    densityBand: next.densityBand,
    layoutRisk: next.layoutRisk,
  });
  const fallbackReadability = "文字必须明显偏大，标题优先使用大号中文标题，正文宁可减少碎句也不要做成密集小字。";
  return [
    theme.basic,
    theme[next.pageType] || theme.content,
    buildDecorationPromptBlock(getPageDecorationLevel(next)),
    buildNoExtraTextConstraintBlock({ scope: "decoration" }),
    `输出要求：${aspect.label} 画幅，目标尺寸 ${aspect.outputWidth}x${aspect.outputHeight}。这是约 2K 的 PPT 画面，必须优先保证页面清晰度、标题醒目度和中文正文可读性。`,
    harnessBlock,
    "除非特殊说明，否则画面中的文字请使用与原文一致的语言。",
    visiblePlan.isManualOverride
      ? "用户已手工确认最终上屏内容，允许出现的文字仅限当前白名单，不得补出任何额外信息。"
      : "当前页面内容中的每一行都视为必要文本，成员、署名、机构、年份、单位、参考信息和关键数字不得遗漏。",
    `当前页面类型：${PAGE_TYPE_LABELS[next.pageType] || next.pageType}`,
    displayTitle ? `当前页面标题：${displayTitle}` : "",
    next.visualFocus ? `视觉焦点：${next.visualFocus}` : "",
    next.layoutSummary ? `版式安排：${next.layoutSummary}` : "",
    next.textHierarchy ? `文字层级：${next.textHierarchy}` : "",
    next.readabilityNotes ? `可读性要求：${next.readabilityNotes}` : fallbackReadability,
    "视觉增强许可：允许加入与主题强相关的无字信息图元素，例如时间线、技术结构示意、功能图标、场景小图、流程箭头、对比模块、卡片容器和背景材质层次，只要这些新增元素不引入新的可读文字或数字。",
    `当前页面内容：\n${next.pageContent || next.pageTitle || ""}`,
  ].filter(Boolean).join("\n\n");
}

function buildManualPromptDraft() {
  const theme = requireConfirmedTheme();
  const pageGoal = el.manualPageGoal.value.trim();
  const regionText = buildSlideRegionText();
  const pageType = inferManualPageType();
  const harnessBlock = buildHarnessPromptBlock("manual", {
    chunkIds: ["readability_rules", "layout_system", "page_density_strategies", "quality_checks"],
    includeHardRules: true,
    maxHardRules: 1,
    styleSeed: theme,
    densityBand: inferWorkflowDensityBand([pageGoal, regionText].filter(Boolean).join(" ")),
  });
  return [
    theme.basic,
    theme[pageType] || theme.content,
    buildDecorationPromptBlock(getGlobalDecorationLevel()),
    `输出要求：${getPptOutputDescription()}，请严格按 PPT 画幅构图。`,
    pageGoal ? `当前页面目标：${pageGoal}` : "",
    regionText,
    harnessBlock || "请把版式秩序、文本留白和视觉层级一起考虑进去，优先保证标题和正文可读性。",
  ].filter(Boolean).join("\n\n");
}

function buildPromptText() {
  const prompt = el.prompt.value.trim();
  if (!prompt) throw new Error("请先填写提示词。");
  const parts = [];
  if (hasConfirmedThemeDefinition() && state.activeTabs.main === "manual") {
    const theme = getCurrentThemePreset();
    const pageType = inferManualPageType();
    const themePrefix = [theme.basic, theme[pageType] || theme.content];
    if (!themePrefix.some((entry) => entry && prompt.includes(entry.slice(0, 20)))) parts.push(...themePrefix.filter(Boolean));
    const harnessBlock = buildHarnessPromptBlock("manual", {
      chunkIds: ["readability_rules", "layout_system", "page_density_strategies", "quality_checks"],
      includeHardRules: true,
      maxHardRules: 1,
      styleSeed: theme,
      densityBand: inferWorkflowDensityBand([el.manualPageGoal.value, buildSlideRegionText()].join(" ")),
    });
    if (harnessBlock && !prompt.includes("【PPT Harness】")) parts.push(harnessBlock);
    if (!prompt.includes("装饰强度")) parts.push(buildDecorationPromptBlock(getGlobalDecorationLevel()));
  }
  parts.push(prompt);
  if (el.includeRegionsInPrompt.checked && state.slideRegions.length) parts.push(buildSlideRegionText());
  return parts.filter(Boolean).join("\n\n");
}

function buildWorkflowAssistantPayload() {
  const content = el.workflowContent.value.trim();
  const pageCount = Number(el.workflowPageCount.value || 0);
  const theme = requireConfirmedTheme();
  const harnessBlock = buildHarnessPromptBlock("planner", {
    chunkIds: ["core_principles", "simplicity_rules", "layout_system", "page_density_strategies", "content_element_rules", "quality_checks"],
    includeHardRules: true,
    maxHardRules: 1,
    styleSeed: theme,
  });
  if (!content) throw new Error("请先粘贴需要拆分的长文本或讲稿。");
  if (!Number.isInteger(pageCount) || pageCount < 2 || pageCount > 20) throw new Error("目标页数请输入 2 到 20 之间的整数。");
  return {
    model: "qwen3.6-plus",
    input: {
      messages: [
        { role: "system", content: ["你是一名专业的 PPT 内容策划师和视觉脚本编辑。", "你需要把用户原文拆成指定页数的 PPT 页面规划，并严格返回 JSON。", "除了封面标题外，其余页的 pageContent 尽量直接使用用户原文，不要改写、总结或漏掉关键信息。", harnessBlock].filter(Boolean).join("\n") },
        { role: "user", content: [{ text: [`当前视觉主题：${theme.label}`, `当前主题风格定义：\n${formatJSON({ basic: theme.basic, cover: theme.cover, content: theme.content, data: theme.data })}`, `输出尺寸固定为：${getPptOutputDescription()}`, `请把下面内容规划成恰好 ${pageCount} 页 PPT。`, "硬性规则：", `1. 总页数必须严格等于 ${pageCount}。`, "2. 第 1 页必须是 cover。", "3. cover 页只放标题或主题概括，允许从原文提炼一个简洁标题。", "4. 第 2 页及之后的 pageContent 尽量直接使用用户原文，不要总结、不要洗稿、不要重写。", "5. pageType 只能是 cover、content、data；包含明确数字、占比、同比、环比、KPI 或统计结果的页优先使用 data。", "6. 拆分优先按章节、主题边界、自然段和逻辑单元进行，不要在句子中间硬切。", "7. 输出格式必须是 JSON object，字段包含 summary 和 pagePlan。", "8. pagePlan 中每一页都必须包含 pageNumber、pageType、pageTitle、pageContent。", "9. 可以额外输出 densityBand 和 layoutRisk；如果未输出，系统会自动计算。", "10. 当某页内容明显偏密时，不要私自减少页数或删除内容，只能在 summary 或页面元数据中标记阅读风险。", "11. 不要输出图片提示词，后续系统会继续为每一页单独做版式整理和出图 prompt。", "输出示例：", "{\"summary\":\"...\",\"pagePlan\":[{\"pageNumber\":1,\"pageType\":\"cover\",\"pageTitle\":\"...\",\"pageContent\":\"...\",\"densityBand\":\"lite\",\"layoutRisk\":\"low\"},{\"pageNumber\":2,\"pageType\":\"content\",\"pageTitle\":\"...\",\"pageContent\":\"...\",\"densityBand\":\"standard\",\"layoutRisk\":\"medium\"}]}", "用户原文：", content].join("\n\n") }] },
      ],
    },
    parameters: { result_format: "message", response_format: { type: "json_object" }, enable_thinking: el.assistantThinking.checked },
  };
}

function buildWorkflowPageDesignPayload(page) {
  const theme = requireConfirmedTheme();
  const aspect = getCurrentAspectMeta();
  const next = applyHarnessMetaToPage({ ...page });
  const displayTitle = getWorkflowDisplayTitle(next);
  const textPlacementBlock = buildWorkflowTextPlacementBlock(next);
  const hierarchyOutline = buildWorkflowHierarchyOutline(next);
  const hierarchyRules = [
    "内容分层规则：先把当前页内容整理成最多 4 层，再决定版式。",
    "第 1 层是页标题。",
    "第 2 层是一级标题或一级信息组。",
    "第 3 层是二级标题，或一级标题的附属说明、数据、注释。",
    "第 4 层是二级标题的附属内容，只保留必要补充。",
    "结构关系只允许总分结构或并列结构，不要把不同层级混成大段文本。",
    "textHierarchy 必须按这 4 层以内的结构写清楚，明确哪些是页标题、一级、二级和附属内容。",
  ].join("\n");
  const harnessBlock = buildHarnessPromptBlock("pageDesign", {
    chunkIds: ["core_principles", "readability_rules", "simplicity_rules", "layout_system", "page_density_strategies", "content_element_rules", "quality_checks"],
    includeHardRules: true,
    maxHardRules: 1,
    styleSeed: theme,
    densityBand: next.densityBand,
    layoutRisk: next.layoutRisk,
  });
  return {
    model: "qwen3.6-plus",
    input: {
      messages: [
        { role: "system", content: ["你是一位顶级 PPT 页面导演、信息设计师和提示词工程师。", "你的任务不是重写原文，而是基于当前页面的内容、页面类型和全局风格模板，为这一页生成更适合出图的排版策略与最终页面提示词。", "重点处理四件事：信息主次、视觉重心、留白与版式、字体大小和可读性。", "必须优先让标题、关键数字、核心结论和正文在 PPT 场景中清楚可读，避免满屏密集小字。", "可读文字必须严格受控，但非文字的视觉补充可以更大胆、更像信息图。", harnessBlock, "请只输出 JSON object，不要输出 markdown。"].filter(Boolean).join("\n") },
        { role: "user", content: [{ text: [`全局风格主题：${theme.label}`, `全局风格模板：\n${formatJSON({ basic: theme.basic, cover: theme.cover, content: theme.content, data: theme.data })}`, `当前页面类型：${PAGE_TYPE_LABELS[next.pageType] || next.pageType}`, `当前页面标题：${displayTitle || `第 ${next.pageNumber} 页`}`, `当前页面原文：\n${next.pageContent || next.pageTitle || ""}`, `输出尺寸固定为 ${aspect.label}，${aspect.outputWidth}x${aspect.outputHeight}，目标是清晰的 2K PPT 画面。`, `当前页面密度档位：${next.densityBand}`, `当前页面版式风险：${next.layoutRisk}`, textPlacementBlock, "请根据这一页内容单独思考版式，不要把所有页都套成同一种布局。", "请特别注意：", "1. 标题区必须明显更大，优先保证远距离可读。", "2. 正文宁可拆成 2 到 4 个大内容块，也不要做成密集小字墙。", "3. 数据页要放大关键数字或核心指标，不要把数据埋在角落里。", "4. dense 页面要优先使用分栏、多块、图表化和延续建议，必要时降低装饰性。", "5. pagePrompt 必须是一整段可以直接发给生图模型的中文提示词，且这段提示词里必须包含当前页面标题与正文文本块，禁止写成只出背景或后续再加字。", "6. 如果当前页包含作者、成员、机构、来源、时间、页脚说明或副标题，这些都属于必要信息，不得省略，不得降成模糊装饰。", "7. 允许你补充更丰富的无字视觉表达，例如技术原理示意、时间线、场景小图、数据图形、图标组、结构剖面或卡片模块，只要不新增未经提供的可读文字或事实。", "8. layoutSummary、textHierarchy、visualFocus、readabilityNotes 必须输出为简洁中文字符串，不要输出对象或数组。", "返回 JSON，字段必须包含：layoutSummary、textHierarchy、visualFocus、readabilityNotes、pagePrompt、densityBand、layoutRisk。"].join("\n\n") }] },
      ],
    },
    parameters: { result_format: "message", response_format: { type: "json_object" }, enable_thinking: el.assistantThinking.checked },
  };
}

function applyWorkflowPageDesign(page, parsed) {
  page.layoutSummary = stringifyStructuredField(parsed?.layoutSummary || parsed?.layout_summary);
  page.textHierarchy = stringifyStructuredField(parsed?.textHierarchy || parsed?.text_hierarchy);
  page.visualFocus = stringifyStructuredField(parsed?.visualFocus || parsed?.visual_focus);
  page.readabilityNotes = stringifyStructuredField(parsed?.readabilityNotes || parsed?.readability_notes);
  page.densityBand = normalizeDensityBandValue(parsed?.densityBand || parsed?.density_band || page.densityBand) || inferWorkflowDensityBand(page);
  page.layoutRisk = normalizeLayoutRiskValue(parsed?.layoutRisk || parsed?.layout_risk || page.layoutRisk) || inferWorkflowLayoutRisk(page);
  page.pagePrompt = String(parsed?.pagePrompt || parsed?.page_prompt || "").trim() || buildWorkflowPagePrompt(page);
  page.layoutStatus = "ready";
  page.layoutError = "";
}

function buildWorkflowPageLayoutText(page) {
  const next = applyHarnessMetaToPage({ ...page });
  if (next.layoutStatus === "running") return "正在根据这一页内容、页面类型和共享风格模板整理个性化版式...";
  return [
    `页面类型：${PAGE_TYPE_LABELS[next.pageType] || next.pageType}`,
    `输出尺寸：${getPptOutputDescription()}`,
    `页面密度：${next.densityBand}`,
    `版式风险：${next.layoutRisk}`,
    next.layoutSummary || "这一页还没有生成单独的版式安排。拆分页完成后，系统会自动补做这一页的版式整理。",
    next.textHierarchy ? `文字层级：${next.textHierarchy}` : "",
    next.visualFocus ? `视觉焦点：${next.visualFocus}` : "",
    next.readabilityNotes ? `可读性要求：${next.readabilityNotes}` : "默认要求：标题更大、正文不要密集小字、关键结论和数字优先放大。",
  ].filter(Boolean).join("\n\n");
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
          content: [{ text: buildEffectiveWorkflowPagePrompt(page) }],
        },
      ],
    },
    parameters,
  };
}

async function requestWorkflowResearchSupplements(page) {
  if (!page) return;
  if (!el.apiKey.value.trim()) {
    setStatus("请先填写 API Key。", "error");
    return;
  }

  page.researchStatus = "running";
  page.researchSummary = "";
  page.researchRaw = "";
  renderWorkflowDetail();
  saveSettings();

  try {
    const response = await fetch("/api/research-supplements", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        apiKey: el.apiKey.value.trim(),
        region: el.region.value,
        themeLabel: getCurrentThemePreset().label,
        searchQuery: page.researchQueryOverride || "",
        page: {
          pageNumber: page.pageNumber,
          pageType: page.pageType,
          pageTitle: getWorkflowDisplayTitle(page) || page.pageTitle,
          pageContent: page.pageContent,
        },
        visibleText: buildWorkflowVisibleTextPreview(page),
      }),
    });

    const data = await response.json();
    if (!response.ok || data.code) {
      page.researchStatus = "error";
      page.researchSummary = String(data?.message || "").trim();
      page.researchRaw = stringifyResearchRaw(data?.raw || data);
      renderWorkflowDetail();
      saveSettings();
      setStatus(data.message || "联网补充失败。", "error");
      return;
    }

    page.researchStatus = "ready";
    page.researchSummary = String(data.summary || "").trim();
    page.researchCandidates = normalizeWorkflowResearchCandidates(data.candidates);
    page.researchRaw = String(data.raw || "").trim();
    page.researchLastQuery = String(data.searchQuery || page.researchQueryOverride || "").trim();
    renderWorkflowDetail();
    saveSettings();
    setStatus(`第 ${page.pageNumber} 页已生成 ${page.researchCandidates.length} 条联网补充候选。`, "success");
  } catch (error) {
    page.researchStatus = "error";
    page.researchSummary = String(error?.message || "").trim();
    page.researchRaw = stringifyResearchRaw(error?.stack || error?.message || error);
    renderWorkflowDetail();
    saveSettings();
    setStatus(error.message || "联网补充请求失败。", "error");
  }
}

function applySelectedWorkflowResearchCandidates(page) {
  if (!page) return;
  const selected = normalizeWorkflowResearchCandidates(page.researchCandidates).filter((item) => item.selected);
  if (!selected.length) {
    setStatus("请先勾选要并入的联网补充候选。", "error");
    return;
  }

  const plan = buildWorkflowVisibleTextPlan(page);
  const nextLines = [...plan.visibleLines];
  selected.forEach((item) => {
    if (!nextLines.includes(item.text)) nextLines.push(item.text);
  });
  page.displayBodyOverride = nextLines.join("\n");
  saveSettings();
  renderWorkflowPlan();
  renderWorkflowDetail();
  updatePayloadPreview();
  setStatus(`已将 ${selected.length} 条联网补充并入第 ${page.pageNumber} 页的最终上屏内容。`, "success");
}

function renderWorkflowDetail() {
  const page = state.workflowPages.find((item) => item.id === state.workflowDetailPageId);
  if (!page) {
    el.workflowDetailModal?.classList.add("hidden");
    el.workflowDetailModal?.setAttribute("aria-hidden", "true");
    return;
  }

  const theme = getCurrentThemePreset();
  const updateDetailDerivedContent = () => {
    const effectivePrompt = buildEffectiveWorkflowPagePrompt(page);
    const visiblePreview = buildWorkflowVisibleTextPreview(page);
    const policySummary = buildWorkflowTextPolicySummary(page);
    el.workflowDetailMeta.textContent = `${PAGE_TYPE_LABELS[page.pageType] || page.pageType} · ${getWorkflowDisplayTitle(page) || page.pageTitle || `第 ${page.pageNumber} 页`} · 输出 ${getPptOutputDescription()}`;
    el.workflowDetailContent.textContent = `最终上屏内容预览：\n${visiblePreview}\n\n保留策略：\n${policySummary}\n\n原文内容：\n${page.pageContent || ""}`;
    el.workflowDetailPrompt.textContent = effectivePrompt || "// 这一页的最终出图提示词还没准备好。打开详情后会先整理本页版式，再生成最终提示词。";
    el.workflowDetailUseBtn.disabled = page.status === "running" || !effectivePrompt;
    el.workflowDetailCopyBtn.disabled = page.status === "running" || !effectivePrompt;
    if (el.workflowDetailResearchStatus) {
      el.workflowDetailResearchStatus.textContent = page.researchStatus === "running"
        ? "正在联网搜索并生成补充候选..."
        : page.researchStatus === "error"
          ? (page.researchSummary || "联网补充失败，请稍后重试。")
          : page.researchStatus === "ready"
            ? (page.researchSummary || `已生成 ${normalizeWorkflowResearchCandidates(page.researchCandidates).length} 条可核验补充候选。`)
            : "还没有联网补充候选。";
    }
    if (el.workflowDetailResearchApplyBtn) {
      el.workflowDetailResearchApplyBtn.disabled = page.researchStatus === "running" || !normalizeWorkflowResearchCandidates(page.researchCandidates).some((item) => item.selected);
      el.workflowDetailResearchApplyBtn.textContent = "加入待确认内容";
    }
    if (el.workflowDetailResearchQueryInput) {
      el.workflowDetailResearchQueryInput.placeholder = page.researchLastQuery || "例如：smart window electrochromic thermochromic PDLC history";
    }
    if (el.workflowDetailResearchBtn) {
      el.workflowDetailResearchBtn.disabled = page.researchStatus === "running";
      el.workflowDetailResearchBtn.textContent = page.researchStatus === "running" ? "联网补充中..." : "联网补充候选";
    }
    if (el.workflowDetailResearchList) {
      const candidates = normalizeWorkflowResearchCandidates(page.researchCandidates);
      if (!candidates.length) {
        el.workflowDetailResearchList.innerHTML = "";
      } else {
        el.workflowDetailResearchList.innerHTML = candidates.map((candidate, index) => `
          <article class="workflow-research-item">
            <label>
              <input type="checkbox" class="workflow-research-checkbox" data-index="${index}" ${candidate.selected ? "checked" : ""} />
              <div>
                <strong>${escapeHtml(candidate.text)}</strong>
                <p>${escapeHtml(candidate.why || "这条内容可作为有依据的补充信息。")}</p>
                <div class="workflow-research-sources">
                  ${candidate.sources.map((source) => `<a href="${escapeHtml(source.url)}" target="_blank" rel="noreferrer">${escapeHtml(source.title)}</a>`).join("")}
                </div>
              </div>
            </label>
          </article>
        `).join("");
        el.workflowDetailResearchList.querySelectorAll(".workflow-research-checkbox").forEach((checkbox) => {
          checkbox.addEventListener("change", (event) => {
            const index = Number(event.currentTarget.dataset.index);
            if (!Number.isInteger(index) || !page.researchCandidates[index]) return;
            page.researchCandidates[index].selected = Boolean(event.currentTarget.checked);
            saveSettings();
            updateDetailDerivedContent();
          });
        });
      }
    }
  };
  const plan = buildWorkflowVisibleTextPlan(page);
  el.workflowDetailTitle.textContent = `第 ${page.pageNumber} 页详情`;
  updateDetailDerivedContent();
  if (el.workflowDetailDisplayTitleInput) {
    el.workflowDetailDisplayTitleInput.value = getWorkflowDisplayTitle(page);
    el.workflowDetailDisplayTitleInput.oninput = () => {
      page.displayTitleOverride = el.workflowDetailDisplayTitleInput.value;
      saveSettings();
      renderWorkflowPlan();
      updateDetailDerivedContent();
      updatePayloadPreview();
    };
  }
  if (el.workflowDetailVisibleTextEditor) {
    el.workflowDetailVisibleTextEditor.value = plan.visibleLines.join("\n");
    el.workflowDetailVisibleTextEditor.oninput = () => {
      page.displayBodyOverride = el.workflowDetailVisibleTextEditor.value;
      saveSettings();
      renderWorkflowPlan();
      updateDetailDerivedContent();
      updatePayloadPreview();
    };
  }
  if (el.workflowDetailResearchQueryInput) {
    el.workflowDetailResearchQueryInput.value = page.researchQueryOverride || "";
    el.workflowDetailResearchQueryInput.oninput = () => {
      page.researchQueryOverride = el.workflowDetailResearchQueryInput.value.trim();
      saveSettings();
      updateDetailDerivedContent();
    };
  }
  if (el.workflowDetailVisibleResetBtn) {
    el.workflowDetailVisibleResetBtn.onclick = () => {
      page.displayTitleOverride = null;
      page.displayBodyOverride = null;
      saveSettings();
      renderWorkflowPlan();
      if (el.workflowDetailDisplayTitleInput) el.workflowDetailDisplayTitleInput.value = getWorkflowDisplayTitle(page);
      if (el.workflowDetailVisibleTextEditor) el.workflowDetailVisibleTextEditor.value = buildWorkflowVisibleTextPlan(page).visibleLines.join("\n");
      updateDetailDerivedContent();
      updatePayloadPreview();
      setStatus(`第 ${page.pageNumber} 页已恢复系统建议上屏内容。`, "success");
    };
  }
  if (el.workflowDetailResearchBtn) {
    el.workflowDetailResearchBtn.onclick = () => requestWorkflowResearchSupplements(page);
  }
  if (el.workflowDetailResearchApplyBtn) {
    el.workflowDetailResearchApplyBtn.onclick = () => applySelectedWorkflowResearchCandidates(page);
  }
  el.workflowDetailLayout.textContent = buildWorkflowPageLayoutText(page);
  el.workflowDetailTheme.textContent = [
    `Basic：${theme.basic}`,
    "",
    `本页分类风格：${theme[page.pageType] || theme.content}`,
  ].join("\n");
  el.workflowDetailRunBtn.disabled = page.status === "running" || state.workflowRunning || page.layoutStatus === "running";
  el.workflowDetailRunBtn.textContent = page.status === "running" ? "生成中..." : "生成本页";

  el.workflowDetailUseBtn.onclick = () => {
    const effectivePrompt = buildEffectiveWorkflowPagePrompt(page);
    el.prompt.value = effectivePrompt || "";
    saveSettings();
    updatePayloadPreview();
    setStatus(`已将第 ${page.pageNumber} 页提示词写入主提示词。`, "success");
  };
  el.workflowDetailCopyBtn.onclick = async () => {
    try {
      const effectivePrompt = buildEffectiveWorkflowPagePrompt(page);
      await navigator.clipboard.writeText(effectivePrompt || "");
      setStatus(`第 ${page.pageNumber} 页提示词已复制。`, "success");
    } catch (error) {
      setStatus(error.message || "复制提示词失败。", "error");
    }
  };
  el.workflowDetailRunBtn.onclick = () => generateWorkflowPage(page.id);
  el.workflowDetailCloseBtn.onclick = closeWorkflowDetail;
}

function renderThemeStatus() {
  const themeName = getThemeName();
  const aspect = getCurrentAspectMeta();
  const outputLabel = `${aspect.label} · ${aspect.outputWidth}×${aspect.outputHeight}`;
  const harnessLabel = getHarnessStatusLabel();
  if (!themeName) {
    el.themeStatus.textContent = `未生成 · ${harnessLabel}`;
  } else if (hasConfirmedThemeDefinition(themeName)) {
    el.themeStatus.textContent = `已确认 · ${themeName} · ${outputLabel} · ${harnessLabel}`;
  } else if (hasCurrentThemeDefinition(themeName)) {
    el.themeStatus.textContent = `待确认 · ${themeName} · ${harnessLabel}`;
  } else if (state.themeDefinitionSource && state.themeDefinitionSource !== themeName) {
    el.themeStatus.textContent = `需重生成 · ${themeName} · ${harnessLabel}`;
  } else {
    el.themeStatus.textContent = `未生成 · ${themeName} · ${harnessLabel}`;
  }
  if (el.themeDefinitionPreview) {
    el.themeDefinitionPreview.textContent = hasCurrentThemeDefinition(themeName) ? (state.themeDefinitionRaw || formatJSON(state.themeDefinition)) : "";
  }
  if (el.workflowTheme) el.workflowTheme.value = themeName;
  renderThemeLibrary();
  renderThemeReviewPanel();
}

function bindPptHarnessEvents() {
  if (!el.pptHarnessEnabled || el.pptHarnessEnabled.dataset.bound === "true") return;
  el.pptHarnessEnabled.dataset.bound = "true";
  el.pptHarnessEnabled.addEventListener("change", () => {
    if (state.workflowPages.length) {
      resetAllWorkflowPageDesigns(isPptHarnessEnabled() ? "PPT Harness 已开启，请重新整理逐页版式。" : "PPT Harness 已关闭，请重新整理逐页版式。");
      state.workflowPages.forEach((page) => {
        applyHarnessMetaToPage(page);
        if (hasConfirmedThemeDefinition()) page.pagePrompt = buildWorkflowPagePrompt(page);
      });
    }
    renderThemeStatus();
    renderWorkflowPlan();
    renderWorkflowDetail();
    renderManualLayoutPreview();
    saveSettings();
    updatePayloadPreview();
  });
}

function buildWorkflowSuggestedTextPlan(page) {
  return buildWorkflowVisibleTextPlan({
    ...page,
    displayTitleOverride: null,
    displayBodyOverride: null,
  });
}

function primeWorkflowDetailDraft(page, options = {}) {
  if (!page) return null;
  const plan = options.useSuggested ? buildWorkflowSuggestedTextPlan(page) : buildWorkflowVisibleTextPlan(page);
  state.workflowDetailDraft = {
    pageId: page.id,
    title: options.useSuggested ? String(page.pageTitle || "").trim() : getWorkflowDisplayTitle(page),
    body: plan.visibleLines.join("\n"),
    dirty: false,
  };
  return state.workflowDetailDraft;
}

function getWorkflowDetailDraft(page) {
  if (!page) return { pageId: "", title: "", body: "", dirty: false };
  if (!state.workflowDetailDraft || state.workflowDetailDraft.pageId !== page.id) {
    return primeWorkflowDetailDraft(page);
  }
  return state.workflowDetailDraft;
}

function updateWorkflowDetailDraft(page, updates = {}) {
  const draft = getWorkflowDetailDraft(page);
  Object.assign(draft, updates);
  const confirmedTitle = getWorkflowDisplayTitle(page).trim();
  const confirmedBody = buildWorkflowVisibleTextPlan(page).visibleLines.join("\n").trim();
  draft.dirty = draft.title.trim() !== confirmedTitle || draft.body.trim() !== confirmedBody;
  return draft;
}

function hasWorkflowConfirmedContent(page) {
  return hasWorkflowTitleOverride(page) || hasWorkflowBodyOverride(page);
}

function getWorkflowDetailConfirmStatusText(page) {
  const draft = getWorkflowDetailDraft(page);
  if (draft.dirty) return "有未确认修改，先点“确认内容”后才会生效。";
  if (hasWorkflowConfirmedContent(page)) return "当前使用你确认后的上屏内容。";
  return "当前使用系统建议。";
}

function sanitizeWorkflowVisualPromptTemplate(prompt, page) {
  let text = String(prompt || "").replace(/\r/g, "").trim();
  if (!text) return "";

  const plan = buildWorkflowVisibleTextPlan(page);
  const sourceSnippets = [
    String(page?.pageTitle || "").trim(),
    String(page?.pageContent || "").trim(),
    ...plan.rawLines,
    ...plan.visibleLines,
    ...plan.omittedLines,
  ]
    .filter(Boolean)
    .sort((a, b) => b.length - a.length);

  sourceSnippets.forEach((snippet) => {
    if (snippet.length >= 2) {
      text = text.split(snippet).join("");
    }
  });

  text = text
    .replace(/当前页面(?:标题|内容|原文)[：:].*/g, "")
    .replace(/文字落图硬约束[\s\S]*$/g, "")
    .replace(/允许上屏文本白名单[\s\S]*$/g, "")
    .replace(/正文文本块[\s\S]*$/g, "")
    .replace(/保留策略[\s\S]*$/g, "")
    .replace(/最终上屏内容预览[\s\S]*$/g, "")
    .replace(/原文内容[\s\S]*$/g, "")
    .replace(/^\s*(标题|正文|系统建议上屏|已确认上屏预览|结构化排版策略|页面标题|页面内容)[：:].*$/gmu, "")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  return text;
}

function promptLooksTextless(prompt, page) {
  const normalized = String(prompt || "").trim();
  if (!normalized) return true;
  if (/\b无文字\b|不要文字|仅保留主视觉|只保留主视觉|纯背景图|背景图后续添加|后续再加|预留.*文字|仅保留视觉元素/i.test(normalized)) {
    return true;
  }
  if (!String(page?.pageContent || "").match(/研究所|大学|实验室|学院|公司|集团|202[0-9]|20[0-9]{2}/)
    && /研究所|大学|实验室|学院|公司|集团|202[0-9]|20[0-9]{2}/.test(normalized)) {
    return true;
  }
  return false;
}

function buildWorkflowPagePrompt(page) {
  const theme = requireConfirmedTheme();
  const aspect = getCurrentAspectMeta();
  const next = applyHarnessMetaToPage({ ...page });
  const visiblePlan = buildWorkflowVisibleTextPlan(next);
  const hierarchyOutline = buildWorkflowHierarchyOutline(next, visiblePlan);
  const harnessBlock = buildHarnessPromptBlock("pageDesign", {
    chunkIds: ["readability_rules", "simplicity_rules", "layout_system", "page_density_strategies", "content_element_rules", "quality_checks"],
    includeHardRules: true,
    maxHardRules: 1,
    styleSeed: theme,
    densityBand: next.densityBand,
    layoutRisk: next.layoutRisk,
  });
  const fallbackReadability = "文字必须明显偏大，标题优先使用大号中文标题，正文宁可拆成 2 到 4 个大内容块，也不要变成密集小字。";
  const surfaceStrategy = buildWorkflowSurfaceStrategy(next, visiblePlan);
  return [
    theme.basic,
    theme[next.pageType] || theme.content,
    buildDecorationPromptBlock(getPageDecorationLevel(next)),
    buildNoExtraTextConstraintBlock({ scope: "decoration" }),
    `输出要求：${aspect.label} 画幅，目标尺寸 ${aspect.outputWidth}x${aspect.outputHeight}，用于清晰的 4K PPT 页面。`,
    harnessBlock,
    `当前页面类型：${PAGE_TYPE_LABELS[next.pageType] || next.pageType}`,
    getWorkflowDisplayTitle(next) ? `当前页面标题：${getWorkflowDisplayTitle(next)}` : "",
    next.layoutSummary ? `版式骨架：${next.layoutSummary}` : "",
    next.textHierarchy ? `文字层级策略：${next.textHierarchy}` : "",
    hierarchyOutline ? `内容层级提纲：\n${hierarchyOutline}` : "",
    next.visualFocus ? `视觉焦点：${next.visualFocus}` : "",
    next.readabilityNotes ? `可读性要求：${next.readabilityNotes}` : fallbackReadability,
    surfaceStrategy ? `结构策略：${surfaceStrategy}` : "",
    "这段提示词只负责定义画面风格、构图、卡片容器、材质、无字信息图和视觉强化，不负责重复完整正文，也不要额外发明任何上屏文字。",
    "允许加入与主题强相关的无字视觉增强，例如时间线、结构示意、图标组、场景小图、工艺剖面、卡片模块、背景几何和材质层次。",
  ].filter(Boolean).join("\n\n");
}

function getWorkflowVisualPromptTemplate(page) {
  const stored = sanitizeWorkflowVisualPromptTemplate(page?.visualPromptTemplate || page?.pagePrompt || "", page);
  if (stored && !promptLooksTextless(stored, page) && stored.length >= 40) return stored;
  return hasConfirmedThemeDefinition() ? buildWorkflowPagePrompt(page) : stored;
}

function buildWorkflowTextPlacementBlock(page) {
  const plan = buildWorkflowVisibleTextPlan(page);
  const title = plan.title || String(page?.pageTitle || `第 ${page?.pageNumber || ""} 页`).trim();
  const isCover = page?.pageType === "cover";
  const surfaceStrategy = buildWorkflowSurfaceStrategy(page, plan);
  const lines = [
    "文字落图硬约束（优先级高于前面的风格和版式描述）：",
    "1. 这不是纯背景图，必须把下面的中文标题和正文真实排进画面里。",
    "2. 严禁输出“无文字”“只保留主视觉”“后续再加字”“纯背景图”“预留文字区域”等做法。",
    "3. 标题必须作为主标题直接出现在画面中，正文必须以清晰的大字号中文排版呈现。",
    "4. 若正文较多，优先拆成 2 到 4 个文本块、卡片或分栏；不要删掉关键信息，不要缩成密集小字墙。",
    "5. pageContent 中出现的成员名单、署名、机构、年份、参考信息、单位和关键数字都属于必保留信息，不能遗漏。",
    "6. 下方白名单中的每一行都必须逐字使用；禁止同义改写、扩写、释义、补全、润色，禁止额外生成“本页文字：”“说明：”“小结：”这类白名单外标签。",
    "7. 只能使用下面白名单里的文本，禁止新增任何用户未提供的机构名、作者名、年份、英文副标题、页脚说明、来源、数字或占位词。",
    "8. 允许增加非文字的视觉补充，例如图标、结构示意、时间线、工艺剖面、场景缩略图、箭头关系、容器模块、背景几何和材质细节，但这些补充不得包含新的可读文字、标签或数字。",
    buildNoExtraTextConstraintBlock({ requireWhitelist: true }),
    `主标题：${title}`,
  ];
  if (surfaceStrategy) lines.push(`结构化排版策略：${surfaceStrategy}`);
  if (isCover) {
    lines.push("封面页特殊要求：如果封面包含副标题、研究成员、机构或补充说明，它们都必须清晰可辨地出现在画面中，不能降成装饰性纹理、模糊小字或被省略。");
  }
  if (plan.isManualOverride) {
    lines.push("当前页已由用户手工确认最终上屏内容，下面白名单就是唯一允许出现的文字；白名单以外的原文全部不要显示。");
  }
  if (plan.visibleLines.length) {
    lines.push("允许上屏文本白名单：");
    plan.visibleLines.forEach((block) => lines.push(`- 「${block}」`));
  } else {
    lines.push("正文白名单为空：除标题外不要出现任何正文文字。");
  }
  if (!plan.isManualOverride && plan.omittedLines.length) {
    lines.push("可压缩但不可改写的原文：");
    plan.omittedLines.forEach((line, index) => lines.push(`${index + 1}. ${line}`));
  }
  lines.push(buildWorkflowTextPolicySummary(page));
  lines.push("排版提醒：标题显著放大，正文保证远距离可读，关键数字和结论优先放大。你可以用更丰富的无字视觉元素把内容做得更像信息图，但如果白名单里没有某类文字信息，就不要擅自补出那类信息。");
  return lines.join("\n");
}

function buildEffectiveWorkflowPagePrompt(page) {
  return [
    getWorkflowVisualPromptTemplate(page),
    buildDecorationPromptBlock(getPageDecorationLevel(page), { override: true }),
    buildWorkflowTextPlacementBlock(page),
  ].filter(Boolean).join("\n\n");
}

function sanitizeWorkflowPageFields(page) {
  if (!page || typeof page !== "object") return page;
  page.layoutSummary = stringifyStructuredField(page.layoutSummary);
  page.textHierarchy = stringifyStructuredField(page.textHierarchy);
  page.visualFocus = stringifyStructuredField(page.visualFocus);
  page.readabilityNotes = stringifyStructuredField(page.readabilityNotes);
  const hasStructuredLayout = [
    page.layoutSummary,
    page.textHierarchy,
    page.visualFocus,
    page.readabilityNotes,
  ].some((value) => String(value || "").trim());
  page.visualPromptTemplate = sanitizeWorkflowVisualPromptTemplate(page.visualPromptTemplate || page.pagePrompt || "", page);
  page.pagePrompt = page.visualPromptTemplate || "";
  if (!Object.prototype.hasOwnProperty.call(page, "displayTitleOverride")) page.displayTitleOverride = null;
  if (!Object.prototype.hasOwnProperty.call(page, "displayBodyOverride")) page.displayBodyOverride = null;
  if (!Object.prototype.hasOwnProperty.call(page, "researchStatus")) page.researchStatus = "idle";
  if (!Object.prototype.hasOwnProperty.call(page, "researchSummary")) page.researchSummary = "";
  if (!Object.prototype.hasOwnProperty.call(page, "researchCandidates")) page.researchCandidates = [];
  if (!Object.prototype.hasOwnProperty.call(page, "researchRaw")) page.researchRaw = "";
  if (!Object.prototype.hasOwnProperty.call(page, "researchQueryOverride")) page.researchQueryOverride = "";
  if (!Object.prototype.hasOwnProperty.call(page, "researchLastQuery")) page.researchLastQuery = "";
  if (!Object.prototype.hasOwnProperty.call(page, "resultImages")) page.resultImages = [];
  if (!Object.prototype.hasOwnProperty.call(page, "savedResults")) page.savedResults = {};
  if (!Object.prototype.hasOwnProperty.call(page, "status")) page.status = "idle";
  if (!Object.prototype.hasOwnProperty.call(page, "error")) page.error = "";
  if (!Object.prototype.hasOwnProperty.call(page, "layoutStatus")) {
    page.layoutStatus = hasStructuredLayout ? "ready" : "idle";
  } else if (String(page.layoutStatus || "").trim().toLowerCase() === "ready" && !hasStructuredLayout) {
    page.layoutStatus = "idle";
  }
  if (!Object.prototype.hasOwnProperty.call(page, "layoutError")) page.layoutError = "";
  if (!page.pagePrompt && hasConfirmedThemeDefinition()) {
    page.visualPromptTemplate = buildWorkflowPagePrompt(page);
    page.pagePrompt = page.visualPromptTemplate;
  }
  return page;
}

function serializeWorkflowPage(page) {
  const next = sanitizeWorkflowPageFields(applyHarnessMetaToPage({ ...page }));
  return {
    pageNumber: next.pageNumber,
    pageType: next.pageType,
    pageTitle: next.pageTitle,
    pageContent: next.pageContent,
    decorationLevel: next.decorationLevel,
    densityBand: next.densityBand,
    layoutRisk: next.layoutRisk,
    layoutSummary: stringifyStructuredField(next.layoutSummary),
    textHierarchy: stringifyStructuredField(next.textHierarchy),
    visualFocus: stringifyStructuredField(next.visualFocus),
    readabilityNotes: stringifyStructuredField(next.readabilityNotes),
    pagePrompt: next.pagePrompt || "",
    visualPromptTemplate: next.visualPromptTemplate || "",
    displayTitleOverride: next.displayTitleOverride,
    displayBodyOverride: next.displayBodyOverride,
    researchStatus: next.researchStatus || "idle",
    researchSummary: next.researchSummary || "",
    researchCandidates: normalizeWorkflowResearchCandidates(next.researchCandidates),
    researchRaw: next.researchRaw || "",
    researchQueryOverride: next.researchQueryOverride || "",
    researchLastQuery: next.researchLastQuery || "",
    layoutStatus: next.layoutStatus || "idle",
    layoutError: next.layoutError || "",
    status: next.status || "idle",
    error: next.error || "",
    resultImages: Array.isArray(next.resultImages) ? next.resultImages.filter(Boolean) : [],
    savedResults: next.savedResults && typeof next.savedResults === "object" ? next.savedResults : {},
    requestId: next.requestId || "",
    taskId: next.taskId || "",
    detailBackdropUrl: next.detailBackdropUrl || "",
  };
}

function hydrateWorkflowPages(savedPages) {
  if (!Array.isArray(savedPages)) return [];
  return savedPages
    .map((page, index) => {
      const visualPromptTemplate = sanitizeWorkflowVisualPromptTemplate(
        page.visualPromptTemplate || page.visual_prompt_template || page.pagePrompt || page.page_prompt || "",
        page,
      );
      const resultImages = Array.isArray(page.resultImages || page.result_images)
        ? (page.resultImages || page.result_images).filter(Boolean)
        : [];
      const status = String(page.status || "").trim().toLowerCase();
      const layoutStatus = String(page.layoutStatus || page.layout_status || "").trim().toLowerCase();
      return sanitizeWorkflowPageFields(applyHarnessMetaToPage({
        id: uid(),
        pageNumber: Number(page.pageNumber ?? page.page_number ?? index + 1) || index + 1,
        pageType: normalizeWorkflowPageType(page.pageType ?? page.page_type, index),
        pageTitle: String(page.pageTitle ?? page.page_title ?? `第 ${index + 1} 页`).trim(),
        pageContent: String(page.pageContent ?? page.page_content ?? page.pageTitle ?? "").trim(),
        decorationLevel: page.decorationLevel ?? page.decoration_level,
        densityBand: page.densityBand ?? page.density_band,
        layoutRisk: page.layoutRisk ?? page.layout_risk,
        layoutSummary: stringifyStructuredField(page.layoutSummary || page.layout_summary),
        textHierarchy: stringifyStructuredField(page.textHierarchy || page.text_hierarchy),
        visualFocus: stringifyStructuredField(page.visualFocus || page.visual_focus),
        readabilityNotes: stringifyStructuredField(page.readabilityNotes || page.readability_notes),
        pagePrompt: visualPromptTemplate,
        visualPromptTemplate,
        displayTitleOverride: Object.prototype.hasOwnProperty.call(page, "displayTitleOverride") ? page.displayTitleOverride : (Object.prototype.hasOwnProperty.call(page, "display_title_override") ? page.display_title_override : null),
        displayBodyOverride: Object.prototype.hasOwnProperty.call(page, "displayBodyOverride") ? page.displayBodyOverride : (Object.prototype.hasOwnProperty.call(page, "display_body_override") ? page.display_body_override : null),
        researchStatus: String(page.researchStatus || page.research_status || "idle").trim() || "idle",
        researchSummary: String(page.researchSummary || page.research_summary || "").trim(),
        researchCandidates: normalizeWorkflowResearchCandidates(page.researchCandidates || page.research_candidates),
        researchRaw: String(page.researchRaw || page.research_raw || "").trim(),
        researchQueryOverride: String(page.researchQueryOverride || page.research_query_override || "").trim(),
        researchLastQuery: String(page.researchLastQuery || page.research_last_query || "").trim(),
        layoutStatus: layoutStatus === "running" ? (visualPromptTemplate ? "ready" : "idle") : (layoutStatus || (visualPromptTemplate ? "ready" : "idle")),
        layoutError: String(page.layoutError || page.layout_error || "").trim(),
        layoutPromise: null,
        status: status === "running" ? (resultImages.length ? "success" : "idle") : (status || (resultImages.length ? "success" : "idle")),
        error: String(page.error || "").trim(),
        resultImages,
        savedResults: page.savedResults && typeof page.savedResults === "object" ? page.savedResults : {},
        requestId: String(page.requestId || page.request_id || "").trim(),
        taskId: String(page.taskId || page.task_id || "").trim(),
        detailBackdropUrl: String(page.detailBackdropUrl || page.detail_backdrop_url || "").trim(),
      }));
    })
    .filter((page) => page.pageTitle || page.pageContent)
    .sort((a, b) => a.pageNumber - b.pageNumber);
}

function normalizeWorkflowPlan(parsed) {
  const rawPages = Array.isArray(parsed?.pagePlan) ? parsed.pagePlan : Array.isArray(parsed?.pages) ? parsed.pages : Array.isArray(parsed) ? parsed : [];
  const pages = rawPages
    .map((page, index) => sanitizeWorkflowPageFields(applyHarnessMetaToPage({
      id: uid(),
      pageNumber: Number(page.pageNumber ?? page.page_number ?? index + 1) || index + 1,
      pageType: normalizeWorkflowPageType(page.pageType ?? page.page_type, index),
      pageTitle: String(page.pageTitle ?? page.page_title ?? (index === 0 ? page.pageContent ?? page.page_content ?? "" : `第 ${index + 1} 页`)).trim() || `第 ${index + 1} 页`,
      pageContent: String(page.pageContent ?? page.page_content ?? page.pageTitle ?? page.page_title ?? "").trim(),
      decorationLevel: page.decorationLevel ?? page.decoration_level ?? getGlobalDecorationLevel(),
      densityBand: page.densityBand ?? page.density_band,
      layoutRisk: page.layoutRisk ?? page.layout_risk,
      layoutSummary: stringifyStructuredField(page.layoutSummary || page.layout_summary),
      textHierarchy: stringifyStructuredField(page.textHierarchy || page.text_hierarchy),
      visualFocus: stringifyStructuredField(page.visualFocus || page.visual_focus),
      readabilityNotes: stringifyStructuredField(page.readabilityNotes || page.readability_notes),
      pagePrompt: "",
      visualPromptTemplate: "",
      displayTitleOverride: Object.prototype.hasOwnProperty.call(page, "displayTitleOverride") ? page.displayTitleOverride : (Object.prototype.hasOwnProperty.call(page, "display_title_override") ? page.display_title_override : null),
      displayBodyOverride: Object.prototype.hasOwnProperty.call(page, "displayBodyOverride") ? page.displayBodyOverride : (Object.prototype.hasOwnProperty.call(page, "display_body_override") ? page.display_body_override : null),
      researchStatus: String(page.researchStatus || page.research_status || "idle").trim() || "idle",
      researchSummary: String(page.researchSummary || page.research_summary || "").trim(),
      researchCandidates: normalizeWorkflowResearchCandidates(page.researchCandidates || page.research_candidates),
      researchRaw: String(page.researchRaw || page.research_raw || "").trim(),
      researchQueryOverride: String(page.researchQueryOverride || page.research_query_override || "").trim(),
      researchLastQuery: String(page.researchLastQuery || page.research_last_query || "").trim(),
      layoutStatus: "idle",
      layoutError: "",
      layoutPromise: null,
      status: "idle",
      error: "",
      resultImages: [],
      savedResults: {},
      requestId: "",
      taskId: "",
    })))
    .filter((page) => page.pageTitle || page.pageContent)
    .sort((a, b) => a.pageNumber - b.pageNumber)
    .map((page, index) => sanitizeWorkflowPageFields(applyHarnessMetaToPage({
      ...page,
      pageNumber: index + 1,
      pageType: index === 0 ? "cover" : page.pageType,
      visualPromptTemplate: buildWorkflowPagePrompt(page),
      pagePrompt: buildWorkflowPagePrompt(page),
    })));

  return {
    summary: String(parsed?.summary || parsed?.planSummary || "").trim(),
    pages,
  };
}

function buildThemeDefinitionPayload(themeName, options = {}) {
  const referenceImages = Array.isArray(options.referenceImages) ? options.referenceImages.filter((item) => item?.source) : [];
  const hasReferences = referenceImages.length > 0;
  return {
    model: "qwen3.6-plus",
    input: {
      messages: [
        {
          role: "system",
          content: [
            getThemeAgentSystemPrompt(themeName),
            hasReferences ? "如果用户附带 PPT 截图，请先抽取截图中的配色、容器、留白、网格、字体气质、图表表现和材质语言，再转译成可复用的风格模板。只学习风格，不要照搬截图里的具体文案、机构名、年份和数字。" : "",
          ].filter(Boolean).join("\n\n"),
        },
        {
          role: "user",
          content: [
            ...referenceImages.slice(0, 6).map((item) => ({ image: item.source })),
            {
              text: hasReferences
                ? [
                  themeName ? `目标风格名称：${themeName}` : "请根据截图自动概括一个风格标签。",
                  "请先学习这些 PPT 截图的共同视觉语言：配色、留白、标题强度、正文密度、卡片容器、图表气质、图片材质和整体构图。",
                  "只模仿风格，不要复制截图里的具体文字内容、事实信息和页面结构编号。",
                  "最后输出一份可复用的 basic / cover / content / data 风格 JSON，不要包含任何解释。",
                ].join("\n\n")
                : `现在，请生成“${themeName}”的提示词模板，并输出对应 JSON，不要包含任何解释和说明性内容。`,
            },
          ],
        },
      ],
    },
    parameters: { result_format: "message", response_format: { type: "json_object" }, enable_thinking: el.assistantThinking.checked },
  };
}

function buildWorkflowPageDesignPayload(page) {
  const theme = requireConfirmedTheme();
  const aspect = getCurrentAspectMeta();
  const next = applyHarnessMetaToPage({ ...page });
  const displayTitle = getWorkflowDisplayTitle(next);
  const textPlacementBlock = buildWorkflowTextPlacementBlock(next);
  const hierarchyOutline = buildWorkflowHierarchyOutline(next);
  const hierarchyRules = [
    "内容分层规则：先把当前页内容整理成最多 4 层，再决定版式。",
    "第 1 层是页标题。",
    "第 2 层是一级标题或一级信息组。",
    "第 3 层是二级标题，或一级标题的附属说明、数据、注释。",
    "第 4 层是二级标题的附属内容，只保留必要补充。",
    "结构关系只允许总分结构或并列结构，不要把不同层级混成大段文本。",
    "textHierarchy 必须按这 4 层以内的结构写清楚，明确哪些是页标题、一级、二级和附属内容。",
  ].join("\n");
  const harnessBlock = buildHarnessPromptBlock("pageDesign", {
    chunkIds: ["core_principles", "readability_rules", "simplicity_rules", "layout_system", "page_density_strategies", "content_element_rules", "quality_checks"],
    includeHardRules: true,
    maxHardRules: 1,
    styleSeed: theme,
    densityBand: next.densityBand,
    layoutRisk: next.layoutRisk,
  });
  return {
    model: "qwen3.6-plus",
    input: {
      messages: [
        {
          role: "system",
          content: [
            "你是一位顶级 PPT 页面导演、信息设计师和提示词工程师。",
            "你的任务不是重写正文，而是基于当前页面内容、页面类型和共享风格模板，为这一页生成更适合出图的版式策略与视觉模板。",
            "可读文字必须受控，但非文字的视觉补充可以更大胆、更像信息图。",
            "输出 JSON object，不要 markdown。",
            hierarchyRules,
            harnessBlock,
          ].filter(Boolean).join("\n"),
        },
        {
          role: "user",
          content: [{
            text: [
              `全局风格主题：${theme.label}`,
              `全局风格模板：\n${formatJSON({ basic: theme.basic, cover: theme.cover, content: theme.content, data: theme.data })}`,
              `当前页面类型：${PAGE_TYPE_LABELS[next.pageType] || next.pageType}`,
              `当前页面标题：${displayTitle || `第 ${next.pageNumber} 页`}`,
              `当前页面原文：\n${next.pageContent || next.pageTitle || ""}`,
              `输出尺寸固定为 ${aspect.label}，${aspect.outputWidth}x${aspect.outputHeight}，目标是清晰的 4K PPT 画面。`,
              `当前页面密度档位：${next.densityBand}`,
              `当前页面版式风险：${next.layoutRisk}`,
              buildDecorationPromptBlock(getPageDecorationLevel(next)),
              buildNoExtraTextConstraintBlock({ scope: "decoration", requireWhitelist: true }),
              textPlacementBlock,
              hierarchyOutline ? `内容层级提纲：\n${hierarchyOutline}` : "",
              hierarchyRules,
              "请根据这一页内容单独思考版式，不要把所有页都套成同一种布局。",
              "先判断这一页属于“总分结构”还是“并列结构”，再决定卡片、分栏和视觉焦点。",
              "总分结构：页面上必须有一个总述主信息区，再拆 2 到 4 个分支信息块；并列结构：同级信息必须拆成 2 到 4 个并列模块，不要混成一整段。",
              "最多只允许 4 层内容：页标题、一级标题、二级标题或一级附属内容、二级附属内容。超出的内容必须压缩、并入或删掉装饰性表达。",
              "排版时必须按层级大小、容器、留白、对齐和对比拉开主次，不能把二级和附属内容全部塞进一个文字块。",
              "特别要求：",
              "1. 标题区必须明显更大，优先保证远距离可读。",
              "2. 正文宁可拆成 2 到 4 个大内容块，也不要做成密集小字墙。",
              "3. 数据页要放大关键数字或核心指标，不要把数据埋在角落里。",
              "4. dense 页面要优先使用分栏、多块、图表化和延续建议，必要时降低装饰性。",
              "5. pagePrompt 只输出视觉模板与构图策略，不要把完整标题和正文再重复存一遍；系统会在真正生图前，把最终确认的上屏文字白名单合并进去。",
              "6. 如果当前页包含作者、成员、机构、来源、时间、页脚说明或副标题，这些都属于必要信息，不得省略，不得降成模糊装饰。",
              "7. 允许你补充更丰富的无字视觉表达，例如技术原理示意、时间线、场景小图、数据图形、图标组、结构剖面或卡片模块，只要不新增未经提供的可读文字或事实。",
              "8. layoutSummary、textHierarchy、visualFocus、readabilityNotes 必须输出为简洁中文字符串，不要输出对象或数组。",
              "返回 JSON，字段必须包含：layoutSummary、textHierarchy、visualFocus、readabilityNotes、pagePrompt、densityBand、layoutRisk。",
            ].join("\n\n"),
          }],
        },
      ],
    },
    parameters: { result_format: "message", response_format: { type: "json_object" }, enable_thinking: el.assistantThinking.checked },
  };
}

function applyWorkflowPageDesign(page, parsed) {
  page.layoutSummary = stringifyStructuredField(parsed?.layoutSummary || parsed?.layout_summary);
  page.textHierarchy = stringifyStructuredField(parsed?.textHierarchy || parsed?.text_hierarchy);
  page.visualFocus = stringifyStructuredField(parsed?.visualFocus || parsed?.visual_focus);
  page.readabilityNotes = stringifyStructuredField(parsed?.readabilityNotes || parsed?.readability_notes);
  page.densityBand = normalizeDensityBandValue(parsed?.densityBand || parsed?.density_band || page.densityBand) || inferWorkflowDensityBand(page);
  page.layoutRisk = normalizeLayoutRiskValue(parsed?.layoutRisk || parsed?.layout_risk || page.layoutRisk) || inferWorkflowLayoutRisk(page);
  page.visualPromptTemplate = sanitizeWorkflowVisualPromptTemplate(parsed?.pagePrompt || parsed?.page_prompt || "", page) || buildWorkflowPagePrompt(page);
  page.pagePrompt = page.visualPromptTemplate;
  page.layoutStatus = "ready";
  page.layoutError = "";
}

function applySelectedWorkflowResearchCandidates(page) {
  if (!page) return;
  const selected = normalizeWorkflowResearchCandidates(page.researchCandidates).filter((item) => item.selected);
  if (!selected.length) {
    setStatus("请先勾选要并入的联网补充候选。", "error");
    return;
  }

  const draft = getWorkflowDetailDraft(page);
  const nextLines = parseWorkflowVisibleBodyText(draft.body);
  selected.forEach((item) => {
    if (!nextLines.includes(item.text)) nextLines.push(item.text);
  });
  updateWorkflowDetailDraft(page, { body: nextLines.join("\n") });
  if (el.workflowDetailVisibleTextEditor) el.workflowDetailVisibleTextEditor.value = draft.body;
  renderWorkflowDetail();
  setStatus(`已将 ${selected.length} 条联网补充放入待确认上屏内容，确认后才会生效。`, "success");
}

function renderWorkflowDetail() {
  const page = state.workflowPages.find((item) => item.id === state.workflowDetailPageId);
  if (!page) {
    el.workflowDetailModal?.classList.add("hidden");
    el.workflowDetailModal?.setAttribute("aria-hidden", "true");
    return;
  }

  const theme = getCurrentThemePreset();
  const draft = getWorkflowDetailDraft(page);
  const updateDetailDerivedContent = () => {
    const effectivePrompt = buildEffectiveWorkflowPagePrompt(page);
    const confirmedPreview = buildWorkflowVisibleTextPreview(page);
    const suggestedPreview = buildWorkflowVisibleTextPreview({
      ...page,
      displayTitleOverride: null,
      displayBodyOverride: null,
    });
    const policySummary = buildWorkflowTextPolicySummary(page);
    const hasPendingDraft = getWorkflowDetailDraft(page).dirty;
    el.workflowDetailMeta.textContent = `${PAGE_TYPE_LABELS[page.pageType] || page.pageType} · ${getWorkflowDisplayTitle(page) || page.pageTitle || `第 ${page.pageNumber} 页`} · 输出 ${getPptOutputDescription()}`;
    if (el.workflowDetailSuggestedContent) el.workflowDetailSuggestedContent.textContent = suggestedPreview || "暂无系统建议。";
    if (el.workflowDetailConfirmedContent) el.workflowDetailConfirmedContent.textContent = `${confirmedPreview}\n\n保留策略：\n${policySummary}`;
    if (el.workflowDetailContent) el.workflowDetailContent.textContent = page.pageContent || "";
    el.workflowDetailPrompt.textContent = effectivePrompt || "// 这一页的最终出图提示词还没准备好。打开详情后会先整理本页版式，再生成最终提示词。";
    el.workflowDetailUseBtn.disabled = page.status === "running" || !effectivePrompt || hasPendingDraft;
    el.workflowDetailCopyBtn.disabled = page.status === "running" || !effectivePrompt || hasPendingDraft;
    el.workflowDetailRunBtn.disabled = page.status === "running" || state.workflowRunning || page.layoutStatus === "running" || hasPendingDraft;
    el.workflowDetailRunBtn.textContent = hasPendingDraft ? "先确认内容" : (page.status === "running" ? "生成中..." : "生成本页");
    if (el.workflowDetailConfirmContentBtn) {
      el.workflowDetailConfirmContentBtn.disabled = !hasPendingDraft;
    }
    if (el.workflowDetailConfirmStatus) {
      el.workflowDetailConfirmStatus.textContent = getWorkflowDetailConfirmStatusText(page);
    }
    if (el.workflowDetailResearchStatus) {
      el.workflowDetailResearchStatus.textContent = page.researchStatus === "running"
        ? "正在联网搜索并生成补充候选..."
        : page.researchStatus === "error"
          ? (page.researchSummary || "联网补充失败，请稍后重试。")
          : page.researchStatus === "ready"
            ? (page.researchSummary || `已生成 ${normalizeWorkflowResearchCandidates(page.researchCandidates).length} 条可核验补充候选。`)
            : "还没有联网补充候选。";
    }
    if (el.workflowDetailResearchApplyBtn) {
      el.workflowDetailResearchApplyBtn.disabled = page.researchStatus === "running" || !normalizeWorkflowResearchCandidates(page.researchCandidates).some((item) => item.selected);
    }
    if (el.workflowDetailResearchQueryInput) {
      el.workflowDetailResearchQueryInput.placeholder = page.researchLastQuery || "例如：smart window electrochromic thermochromic PDLC history";
    }
    if (el.workflowDetailResearchBtn) {
      el.workflowDetailResearchBtn.disabled = page.researchStatus === "running";
      el.workflowDetailResearchBtn.textContent = page.researchStatus === "running" ? "联网补充中..." : "联网补充候选";
    }
    if (el.workflowDetailResearchList) {
      const candidates = normalizeWorkflowResearchCandidates(page.researchCandidates);
      if (!candidates.length) {
        el.workflowDetailResearchList.innerHTML = "";
      } else {
        el.workflowDetailResearchList.innerHTML = candidates.map((candidate, index) => `
          <article class="workflow-research-item">
            <label>
              <input type="checkbox" class="workflow-research-checkbox" data-index="${index}" ${candidate.selected ? "checked" : ""} />
              <div>
                <strong>${escapeHtml(candidate.text)}</strong>
                <p>${escapeHtml(candidate.why || "这条内容可作为有依据的补充信息。")}</p>
                <div class="workflow-research-sources">
                  ${candidate.sources.map((source) => `<a href="${escapeHtml(source.url)}" target="_blank" rel="noreferrer">${escapeHtml(source.title)}</a>`).join("")}
                </div>
              </div>
            </label>
          </article>
        `).join("");
        el.workflowDetailResearchList.querySelectorAll(".workflow-research-checkbox").forEach((checkbox) => {
          checkbox.addEventListener("change", (event) => {
            const index = Number(event.currentTarget.dataset.index);
            if (!Number.isInteger(index) || !page.researchCandidates[index]) return;
            page.researchCandidates[index].selected = Boolean(event.currentTarget.checked);
            saveSettings();
            updateDetailDerivedContent();
          });
        });
      }
    }
  };

  el.workflowDetailTitle.textContent = `第 ${page.pageNumber} 页详情`;
  if (el.workflowDetailDisplayTitleInput) {
    el.workflowDetailDisplayTitleInput.value = draft.title;
    el.workflowDetailDisplayTitleInput.oninput = () => {
      updateWorkflowDetailDraft(page, { title: el.workflowDetailDisplayTitleInput.value });
      updateDetailDerivedContent();
    };
  }
  if (el.workflowDetailVisibleTextEditor) {
    el.workflowDetailVisibleTextEditor.value = draft.body;
    el.workflowDetailVisibleTextEditor.oninput = () => {
      updateWorkflowDetailDraft(page, { body: el.workflowDetailVisibleTextEditor.value });
      updateDetailDerivedContent();
    };
  }
  if (el.workflowDetailConfirmContentBtn) {
    el.workflowDetailConfirmContentBtn.onclick = () => {
      const suggestedPlan = buildWorkflowSuggestedTextPlan(page);
      const suggestedBody = suggestedPlan.visibleLines.join("\n").trim();
      const nextTitle = draft.title.trim();
      const nextBody = draft.body.trim();
      page.displayTitleOverride = nextTitle === String(page.pageTitle || "").trim() ? null : draft.title;
      page.displayBodyOverride = nextBody === suggestedBody ? null : draft.body;
      primeWorkflowDetailDraft(page);
      saveSettings();
      renderWorkflowPlan();
      renderWorkflowDetail();
      updatePayloadPreview();
      setStatus(`第 ${page.pageNumber} 页的上屏内容已确认，之后生成将以这份内容为准。`, "success");
    };
  }
  if (el.workflowDetailResearchQueryInput) {
    el.workflowDetailResearchQueryInput.value = page.researchQueryOverride || "";
    el.workflowDetailResearchQueryInput.oninput = () => {
      page.researchQueryOverride = el.workflowDetailResearchQueryInput.value.trim();
      saveSettings();
      updateDetailDerivedContent();
    };
  }
  if (el.workflowDetailVisibleResetBtn) {
    el.workflowDetailVisibleResetBtn.onclick = () => {
      page.displayTitleOverride = null;
      page.displayBodyOverride = null;
      primeWorkflowDetailDraft(page, { useSuggested: true });
      saveSettings();
      renderWorkflowPlan();
      renderWorkflowDetail();
      updatePayloadPreview();
      setStatus(`第 ${page.pageNumber} 页已恢复系统建议上屏内容。`, "success");
    };
  }
  if (el.workflowDetailResearchBtn) {
    el.workflowDetailResearchBtn.onclick = () => requestWorkflowResearchSupplements(page);
  }
  if (el.workflowDetailResearchApplyBtn) {
    el.workflowDetailResearchApplyBtn.onclick = () => applySelectedWorkflowResearchCandidates(page);
  }
  el.workflowDetailLayout.textContent = buildWorkflowPageLayoutText(page);
  el.workflowDetailTheme.textContent = [
    `Basic：${theme.basic}`,
    "",
    `本页分类风格：${theme[page.pageType] || theme.content}`,
  ].join("\n");
  el.workflowDetailUseBtn.onclick = () => {
    const effectivePrompt = buildEffectiveWorkflowPagePrompt(page);
    el.prompt.value = effectivePrompt || "";
    saveSettings();
    updatePayloadPreview();
    setStatus(`已将第 ${page.pageNumber} 页提示词写入主提示词。`, "success");
  };
  el.workflowDetailCopyBtn.onclick = async () => {
    try {
      const effectivePrompt = buildEffectiveWorkflowPagePrompt(page);
      await navigator.clipboard.writeText(effectivePrompt || "");
      setStatus(`第 ${page.pageNumber} 页提示词已复制。`, "success");
    } catch (error) {
      setStatus(error.message || "复制提示词失败。", "error");
    }
  };
  el.workflowDetailRunBtn.onclick = () => generateWorkflowPage(page.id);
  el.workflowDetailCloseBtn.onclick = closeWorkflowDetail;
  updateDetailDerivedContent();
}

function buildWorkflowGenerationPayload(page) {
  const parameters = buildGenerationParameters({
    n: 1,
    orderedImages: [],
    enableSequential: false,
    enableBbox: false,
  });
  parameters.size = getPptOutputSizeValue();

  const effectivePrompt = typeof buildEffectiveWorkflowPagePrompt === "function"
    ? buildEffectiveWorkflowPagePrompt(page)
    : String(page?.pagePrompt || "").trim();

  return {
    model: el.model.value,
    input: {
      messages: [
        {
          role: "user",
          content: [{ text: effectivePrompt }],
        },
      ],
    },
    parameters,
  };
}

function renderWorkflowPlanHistory() {
  if (!el.workflowPlanHistoryPanel) return;
  const activeEntry = state.workflowPlanLibrary.find((entry) => entry.id === state.workflowPlanLibraryActiveId);
  if (!activeEntry) {
    el.workflowPlanHistoryPanel.classList.add("hidden");
    if (el.workflowPlanHistoryMeta) el.workflowPlanHistoryMeta.textContent = "选中后会显示这套方案的主题、时间和封面页摘要。";
    if (el.workflowPlanHistoryLead) el.workflowPlanHistoryLead.textContent = "";
    if (el.workflowPlanHistorySummary) el.workflowPlanHistorySummary.textContent = "";
    if (el.workflowPlanHistoryPages) el.workflowPlanHistoryPages.innerHTML = "";
    if (el.workflowPlanDeleteBtn) el.workflowPlanDeleteBtn.disabled = true;
    return;
  }

  const pages = Array.isArray(activeEntry.pages) ? activeEntry.pages : [];
  const leadPage = pages[0];
  const pageCount = pages.length;
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
  if (el.workflowPlanHistoryPages) {
    el.workflowPlanHistoryPages.innerHTML = pages.map((page, index) => {
      const preview = typeof buildWorkflowVisibleTextPreview === "function"
        ? buildWorkflowVisibleTextPreview(page)
        : String(page.pageContent || "").trim();
      const title = typeof getWorkflowDisplayTitle === "function"
        ? (getWorkflowDisplayTitle(page) || page.pageTitle || `第 ${index + 1} 页`)
        : (page.pageTitle || `第 ${index + 1} 页`);
      return `
        <article class="workflow-history-page">
          <strong>第 ${Number(page.pageNumber || index + 1)} 页 · ${escapeHtml(title)}</strong>
          <p>${escapeHtml(preview || "暂无内容")}</p>
        </article>
      `;
    }).join("");
  }
  if (el.workflowPlanDeleteBtn) el.workflowPlanDeleteBtn.disabled = false;
}

function setWorkflowDetailSectionVisibility(node, visible) {
  if (!node) return;
  node.hidden = !visible;
  const title = node.previousElementSibling;
  if (title?.tagName === "STRONG") title.hidden = !visible;
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

    const tone = page.status === "error"
      ? "error"
      : page.status === "success"
        ? "success"
        : page.status === "running"
          ? "running"
          : "idle";

    const statusMarkup = page.status === "running" || page.status === "error"
      ? `
        <div class="workflow-status" data-tone="${tone}">
          ${escapeHtml(page.error || (page.status === "running" ? "正在生成本页..." : "本页生成失败。"))}
        </div>
      `
      : "";

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
      ${statusMarkup}
      <div class="workflow-section">
        <strong>最终上屏内容</strong>
        <p class="workflow-card-brief">${escapeHtml(
          typeof buildWorkflowVisibleTextPreview === "function"
            ? buildWorkflowVisibleTextPreview(page)
            : (page.pageContent || "")
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
            imageUrl,
          });
        });

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
  const draft = getWorkflowDetailDraft(page);
  const updateDetailDerivedContent = () => {
    const effectivePrompt = buildEffectiveWorkflowPagePrompt(page);
    const hasPendingDraft = getWorkflowDetailDraft(page).dirty;

    el.workflowDetailMeta.textContent = `${PAGE_TYPE_LABELS[page.pageType] || page.pageType} · ${getWorkflowDisplayTitle(page) || page.pageTitle || `第 ${page.pageNumber} 页`} · 输出 ${getPptOutputDescription()}`;
    el.workflowDetailLayout.textContent = buildWorkflowPageLayoutText(page);
    if (el.workflowDetailTheme) el.workflowDetailTheme.textContent = [
      `Basic：${theme.basic}`,
      "",
      `本页分类风格：${theme[page.pageType] || theme.content}`,
    ].join("\n");
    if (el.workflowDetailPrompt) el.workflowDetailPrompt.textContent = effectivePrompt || "";

    if (el.workflowDetailUseBtn) {
      el.workflowDetailUseBtn.hidden = true;
      el.workflowDetailUseBtn.disabled = true;
    }
    if (el.workflowDetailCopyBtn) {
      el.workflowDetailCopyBtn.hidden = true;
      el.workflowDetailCopyBtn.disabled = true;
    }
    if (el.workflowDetailVisibleResetBtn) el.workflowDetailVisibleResetBtn.hidden = false;

    setWorkflowDetailSectionVisibility(el.workflowDetailSuggestedContent, false);
    setWorkflowDetailSectionVisibility(el.workflowDetailConfirmedContent, false);
    setWorkflowDetailSectionVisibility(el.workflowDetailContent, false);
    if (el.workflowDetailResearchQueryInput?.closest(".field")) el.workflowDetailResearchQueryInput.closest(".field").hidden = true;
    if (el.workflowDetailResearchBtn?.closest(".action-row")) el.workflowDetailResearchBtn.closest(".action-row").hidden = true;
    if (el.workflowDetailResearchStatus) el.workflowDetailResearchStatus.hidden = true;
    if (el.workflowDetailResearchList) el.workflowDetailResearchList.hidden = true;
    if (el.workflowDetailTheme?.closest(".theme-review-card")) el.workflowDetailTheme.closest(".theme-review-card").hidden = true;
    if (el.workflowDetailPrompt?.closest(".theme-review-card")) el.workflowDetailPrompt.closest(".theme-review-card").hidden = true;

    if (el.workflowDetailRunBtn) {
      el.workflowDetailRunBtn.disabled = page.status === "running" || state.workflowRunning || page.layoutStatus === "running" || hasPendingDraft;
      el.workflowDetailRunBtn.textContent = hasPendingDraft ? "先确认内容" : (page.status === "running" ? "生成中..." : "生成本页");
    }
    if (el.workflowDetailConfirmContentBtn) {
      el.workflowDetailConfirmContentBtn.disabled = !hasPendingDraft;
    }
    if (el.workflowDetailConfirmStatus) {
      el.workflowDetailConfirmStatus.textContent = hasPendingDraft
        ? "有未确认修改，先点“确认内容”后才会生效。"
        : "当前生成会严格使用这份已确认的上屏内容。";
    }
  };

  if (el.workflowDetailTitle) el.workflowDetailTitle.textContent = `第 ${page.pageNumber} 页详情`;
  if (el.workflowDetailDisplayTitleInput) {
    el.workflowDetailDisplayTitleInput.value = draft.title;
    el.workflowDetailDisplayTitleInput.oninput = () => {
      updateWorkflowDetailDraft(page, { title: el.workflowDetailDisplayTitleInput.value });
      updateDetailDerivedContent();
    };
  }
  if (el.workflowDetailVisibleTextEditor) {
    el.workflowDetailVisibleTextEditor.value = draft.body;
    el.workflowDetailVisibleTextEditor.oninput = () => {
      updateWorkflowDetailDraft(page, { body: el.workflowDetailVisibleTextEditor.value });
      updateDetailDerivedContent();
    };
  }
  if (el.workflowDetailConfirmContentBtn) {
    el.workflowDetailConfirmContentBtn.onclick = () => {
      const suggestedPlan = buildWorkflowSuggestedTextPlan(page);
      const suggestedBody = suggestedPlan.visibleLines.join("\n").trim();
      const nextTitle = draft.title.trim();
      const nextBody = draft.body.trim();
      page.displayTitleOverride = nextTitle === String(page.pageTitle || "").trim() ? null : draft.title;
      page.displayBodyOverride = nextBody === suggestedBody ? null : draft.body;
      primeWorkflowDetailDraft(page);
      saveSettings();
      renderWorkflowPlan();
      renderWorkflowDetail();
      updatePayloadPreview();
      setStatus(`第 ${page.pageNumber} 页的上屏内容已确认，之后生成将以这份内容为准。`, "success");
    };
  }
  if (el.workflowDetailVisibleResetBtn) {
    el.workflowDetailVisibleResetBtn.onclick = () => {
      page.displayTitleOverride = null;
      page.displayBodyOverride = null;
      primeWorkflowDetailDraft(page, { useSuggested: true });
      saveSettings();
      renderWorkflowPlan();
      renderWorkflowDetail();
      updatePayloadPreview();
      setStatus(`第 ${page.pageNumber} 页已恢复默认上屏内容。`, "success");
    };
  }
  if (el.workflowDetailRunBtn) el.workflowDetailRunBtn.onclick = () => generateWorkflowPage(page.id);
  if (el.workflowDetailCloseBtn) el.workflowDetailCloseBtn.onclick = closeWorkflowDetail;
  updateDetailDerivedContent();
}

function normalizeWorkflowDetailHeadings() {
  const detailCard = document.querySelector("#workflowDetailDisplayTitleInput")?.closest(".theme-review-card");
  if (detailCard) {
    const strongs = Array.from(detailCard.querySelectorAll("strong"));
    strongs.forEach((node, index) => {
      if (index === 0) {
        node.textContent = "最终上屏内容（可编辑）";
        node.hidden = false;
      } else {
        node.hidden = true;
      }
    });
  }
}

async function bootstrapPptHarnessIntegration() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}");
    if (el.pptHarnessEnabled) el.pptHarnessEnabled.checked = saved.pptHarnessEnabled !== false;
  } catch {
    if (el.pptHarnessEnabled) el.pptHarnessEnabled.checked = true;
  }

  await loadPptHarness();
  await loadPersistentLibrary();
  state.workflowPages.forEach((page) => {
    applyHarnessMetaToPage(page);
    sanitizeWorkflowPageFields(page);
    if (!page.pagePrompt && hasConfirmedThemeDefinition()) page.pagePrompt = buildWorkflowPagePrompt(page);
  });
  state.workflowPlanLibrary.forEach((entry) => (entry.pages || []).forEach((page) => {
    applyHarnessMetaToPage(page);
    sanitizeWorkflowPageFields(page);
  }));
  bindPptHarnessEvents();
  normalizeWorkflowDetailHeadings();
  renderThemeStatus();
  renderWorkflowPlanLibrary();
  renderWorkflowPlan();
  renderWorkflowDetail();
  renderManualLayoutPreview();
  updatePayloadPreview();
  saveSettings();
  if (state.pptHarnessLoadError) {
    setStatus("已就绪 · PPT Harness 回退", "error");
  }
}

bootstrapPptHarnessIntegration();

function ensureWorkflowDetailWorkspaceShell() {
  if (el.workflowDetailPanel) {
    el.workflowDetailPanel.classList.add("workflow-detail-workspace");
  }

  const detailCard = document.querySelector("#workflowDetailDisplayTitleInput")?.closest(".theme-review-card");
  if (detailCard) {
    detailCard.classList.add("workflow-detail-editor-card");
  }

  const layoutCard = document.querySelector("#workflowDetailLayout")?.closest(".theme-review-card");
  if (layoutCard) {
    layoutCard.classList.add("workflow-detail-layout-card");
  }

  const actions = document.querySelector("#workflowDetailConfirmContentBtn")?.closest(".workflow-detail-editor-actions");
  if (!actions) return {};

  let researchToggleBtn = document.querySelector("#workflowDetailResearchPopoverBtn");
  if (!researchToggleBtn) {
    researchToggleBtn = document.createElement("button");
    researchToggleBtn.id = "workflowDetailResearchPopoverBtn";
    researchToggleBtn.type = "button";
    researchToggleBtn.className = "ghost-btn workflow-detail-mini-btn";
    researchToggleBtn.textContent = "联网补充";
    actions.appendChild(researchToggleBtn);
  }

  return { actions, researchToggleBtn };
}

function ensureWorkflowResearchPopover() {
  let popover = document.querySelector("#workflowResearchPopover");
  if (!popover) {
    popover = document.createElement("section");
    popover.id = "workflowResearchPopover";
    popover.className = "workflow-research-popover hidden";
    popover.hidden = true;
    popover.innerHTML = `
      <div class="workflow-research-popover-head">
        <strong>联网补充</strong>
        <button id="workflowResearchPopoverCloseBtn" type="button" class="ghost-btn workflow-detail-mini-btn">关闭</button>
      </div>
      <label class="field workflow-research-popover-field">
        <span>搜索词</span>
        <input id="workflowResearchPopoverQuery" type="text" placeholder="例如：smart window electrochromic thermochromic PDLC history" />
      </label>
      <div class="action-row workflow-research-popover-actions">
        <button id="workflowResearchPopoverFetchBtn" type="button" class="ghost-btn workflow-detail-mini-btn">联网补充</button>
        <button id="workflowResearchPopoverCopyBtn" type="button" class="ghost-btn workflow-detail-mini-btn">复制已选</button>
      </div>
      <div id="workflowResearchPopoverStatus" class="hint-box workflow-research-status">还没有联网补充候选。</div>
      <div id="workflowResearchPopoverList" class="workflow-research-list"></div>
    `;
  }

  if (el.workflowDetailPanel && popover.parentElement !== el.workflowDetailPanel) {
    el.workflowDetailPanel.appendChild(popover);
  }

  return {
    popover,
    queryInput: popover.querySelector("#workflowResearchPopoverQuery"),
    fetchBtn: popover.querySelector("#workflowResearchPopoverFetchBtn"),
    copyBtn: popover.querySelector("#workflowResearchPopoverCopyBtn"),
    closeBtn: popover.querySelector("#workflowResearchPopoverCloseBtn"),
    status: popover.querySelector("#workflowResearchPopoverStatus"),
    list: popover.querySelector("#workflowResearchPopoverList"),
  };
}

function closeWorkflowResearchPopover() {
  state.workflowResearchPopoverOpen = false;
  state.workflowResearchPopoverPageId = "";
  const { popover } = ensureWorkflowResearchPopover();
  popover.hidden = true;
  popover.classList.add("hidden");
}

function openWorkflowResearchPopover(page) {
  if (!page) return;
  state.workflowResearchPopoverOpen = true;
  state.workflowResearchPopoverPageId = page.id;
  renderWorkflowResearchPopover(page);
}

function getWorkflowResearchStatusText(page) {
  if (!page) return "还没有联网补充候选。";
  if (page.researchStatus === "running") return "正在联网搜索并生成补充候选...";
  if (page.researchStatus === "error") return page.researchSummary || "联网补充失败，请稍后重试。";
  if (page.researchStatus === "ready") {
    return page.researchSummary || `已生成 ${normalizeWorkflowResearchCandidates(page.researchCandidates).length} 条可核验补充候选。`;
  }
  return "还没有联网补充候选。";
}

async function copyWorkflowResearchCandidates(page, specificIndex = null) {
  if (!page) return;
  const candidates = normalizeWorkflowResearchCandidates(page.researchCandidates);
  const selected = specificIndex === null
    ? candidates.filter((item) => item.selected)
    : (Number.isInteger(specificIndex) && candidates[specificIndex] ? [candidates[specificIndex]] : []);

  if (!selected.length) {
    setStatus("请先选择要复制的联网补充内容。", "error");
    return;
  }

  const text = selected.map((item) => item.text).filter(Boolean).join("\n");
  if (!text) {
    setStatus("当前没有可复制的补充文字。", "error");
    return;
  }

  try {
    await navigator.clipboard.writeText(text);
    setStatus(`已复制 ${selected.length} 条联网补充到剪贴板，可以直接粘贴到上屏内容。`, "success");
  } catch (error) {
    setStatus(error.message || "复制联网补充失败。", "error");
  }
}

function renderWorkflowResearchPopover(page) {
  const refs = ensureWorkflowResearchPopover();
  const shouldShow = Boolean(page)
    && Boolean(state.workflowResearchPopoverOpen)
    && state.workflowResearchPopoverPageId === page.id;

  refs.popover.hidden = !shouldShow;
  refs.popover.classList.toggle("hidden", !shouldShow);
  if (!shouldShow) return;

  refs.queryInput.value = page.researchQueryOverride || page.researchLastQuery || "";
  refs.queryInput.oninput = () => {
    page.researchQueryOverride = refs.queryInput.value.trim();
    saveSettings();
  };

  refs.status.textContent = getWorkflowResearchStatusText(page);
  refs.fetchBtn.disabled = page.researchStatus === "running";
  refs.fetchBtn.textContent = page.researchStatus === "running" ? "联网补充中..." : "联网补充";
  refs.copyBtn.disabled = !normalizeWorkflowResearchCandidates(page.researchCandidates).some((item) => item.selected);

  refs.fetchBtn.onclick = async () => {
    page.researchQueryOverride = refs.queryInput.value.trim();
    saveSettings();
    await requestWorkflowResearchSupplements(page);
    renderWorkflowResearchPopover(page);
  };
  refs.copyBtn.onclick = () => {
    copyWorkflowResearchCandidates(page).catch((error) => {
      setStatus(error.message || "复制联网补充失败。", "error");
    });
  };
  refs.closeBtn.onclick = () => closeWorkflowResearchPopover();

  const candidates = normalizeWorkflowResearchCandidates(page.researchCandidates);
  if (!candidates.length) {
    refs.list.innerHTML = "";
    return;
  }

  refs.list.innerHTML = candidates.map((candidate, index) => `
    <article class="workflow-research-item workflow-research-popover-item">
      <div class="workflow-research-item-head">
        <label class="workflow-research-select">
          <input type="checkbox" class="workflow-research-checkbox" data-index="${index}" ${candidate.selected ? "checked" : ""} />
          <span>选择</span>
        </label>
        <button type="button" class="ghost-btn workflow-detail-mini-btn workflow-research-copy-single" data-index="${index}">复制</button>
      </div>
      <strong>${escapeHtml(candidate.text)}</strong>
      <p>${escapeHtml(candidate.why || "这条内容可作为有依据的补充信息。")}</p>
      <div class="workflow-research-sources">
        ${candidate.sources.map((source) => `<a href="${escapeHtml(source.url)}" target="_blank" rel="noreferrer">${escapeHtml(source.title)}</a>`).join("")}
      </div>
    </article>
  `).join("");

  refs.list.querySelectorAll(".workflow-research-checkbox").forEach((checkbox) => {
    checkbox.addEventListener("change", (event) => {
      const index = Number(event.currentTarget.dataset.index);
      if (!Number.isInteger(index) || !page.researchCandidates[index]) return;
      page.researchCandidates[index].selected = Boolean(event.currentTarget.checked);
      saveSettings();
      renderWorkflowResearchPopover(page);
    });
  });

  refs.list.querySelectorAll(".workflow-research-copy-single").forEach((button) => {
    button.addEventListener("click", () => {
      const index = Number(button.dataset.index);
      copyWorkflowResearchCandidates(page, index).catch((error) => {
        setStatus(error.message || "复制联网补充失败。", "error");
      });
    });
  });
}

function renderWorkflowPlanHistory() {
  if (!el.workflowPlanHistoryPanel) return;
  const activeEntry = state.workflowPlanLibrary.find((entry) => entry.id === state.workflowPlanLibraryActiveId);
  const panelHead = el.workflowPlanHistoryPanel.querySelector(".subpanel-head");
  const leadBlock = panelHead?.firstElementChild || null;
  const historyGrid = el.workflowPlanHistoryPanel.querySelector(".workflow-history-grid");

  if (!activeEntry) {
    el.workflowPlanHistoryPanel.classList.add("hidden");
    if (el.workflowPlanHistoryPages) el.workflowPlanHistoryPages.innerHTML = "";
    if (el.workflowPlanDeleteBtn) el.workflowPlanDeleteBtn.disabled = true;
    return;
  }

  const pages = Array.isArray(state.workflowPages) && state.workflowPages.length
    ? state.workflowPages
    : (Array.isArray(activeEntry.pages) ? activeEntry.pages : []);

  el.workflowPlanHistoryPanel.classList.remove("hidden");
  el.workflowPlanHistoryPanel.classList.add("compact");
  if (leadBlock) leadBlock.hidden = true;
  if (historyGrid) historyGrid.hidden = true;
  if (el.workflowPlanHistoryMeta) el.workflowPlanHistoryMeta.hidden = true;
  if (el.workflowPlanDeleteBtn) el.workflowPlanDeleteBtn.disabled = false;

  if (!el.workflowPlanHistoryPages) return;
  el.workflowPlanHistoryPages.innerHTML = pages.map((page, index) => `
    <article class="workflow-history-page" data-page-id="${escapeHtml(page.id || "")}">
      <div class="workflow-page-meta">
        <span class="workflow-page-index">第 ${Number(page.pageNumber || index + 1)} 页</span>
        <span class="workflow-type-chip" data-type="${escapeHtml(page.pageType || "content")}">${escapeHtml(PAGE_TYPE_LABELS[page.pageType] || page.pageType || "内容页")}</span>
      </div>
      <strong>${escapeHtml(getWorkflowDisplayTitle(page) || page.pageTitle || `第 ${index + 1} 页`)}</strong>
      <p>${escapeHtml(buildWorkflowVisibleTextPreview(page) || "暂无内容")}</p>
    </article>
  `).join("");

  el.workflowPlanHistoryPages.querySelectorAll(".workflow-history-page").forEach((node) => {
    node.addEventListener("click", () => {
      const pageId = node.dataset.pageId || "";
      if (!pageId) return;
      openWorkflowDetail(pageId);
    });
  });
}

function normalizeWorkflowDetailHeadings() {
  const detailCard = document.querySelector("#workflowDetailDisplayTitleInput")?.closest(".theme-review-card");
  if (detailCard) {
    const strongs = Array.from(detailCard.querySelectorAll("strong"));
    strongs.forEach((node, index) => {
      if (index === 0) {
        node.textContent = "最终上屏内容（可编辑）";
        node.hidden = false;
        node.style.display = "";
      } else {
        node.hidden = true;
        node.style.display = "none";
      }
    });
  }
}

function renderWorkflowDetail() {
  const page = state.workflowPages.find((item) => item.id === state.workflowDetailPageId);
  if (!page) {
    closeWorkflowResearchPopover();
    el.workflowDetailModal?.classList.add("hidden");
    el.workflowDetailModal?.setAttribute("aria-hidden", "true");
    return;
  }

  const shell = ensureWorkflowDetailWorkspaceShell();
  normalizeWorkflowDetailHeadings();
  const draft = getWorkflowDetailDraft(page);

  const updateDetailDerivedContent = () => {
    const hasPendingDraft = getWorkflowDetailDraft(page).dirty;

    el.workflowDetailMeta.textContent = `${PAGE_TYPE_LABELS[page.pageType] || page.pageType} · ${getWorkflowDisplayTitle(page) || page.pageTitle || `第 ${page.pageNumber} 页`} · 输出 ${getPptOutputDescription()}`;
    el.workflowDetailLayout.textContent = buildWorkflowPageLayoutText(page);

    if (el.workflowDetailUseBtn) {
      el.workflowDetailUseBtn.hidden = true;
      el.workflowDetailUseBtn.disabled = true;
    }
    if (el.workflowDetailCopyBtn) {
      el.workflowDetailCopyBtn.hidden = true;
      el.workflowDetailCopyBtn.disabled = true;
    }
    if (el.workflowDetailVisibleResetBtn) {
      el.workflowDetailVisibleResetBtn.hidden = true;
      el.workflowDetailVisibleResetBtn.disabled = true;
    }
    if (el.workflowDetailConfirmStatus) {
      el.workflowDetailConfirmStatus.hidden = true;
    }

    setWorkflowDetailSectionVisibility(el.workflowDetailSuggestedContent, false);
    setWorkflowDetailSectionVisibility(el.workflowDetailConfirmedContent, false);
    setWorkflowDetailSectionVisibility(el.workflowDetailContent, false);

    if (el.workflowDetailResearchQueryInput?.closest(".field")) el.workflowDetailResearchQueryInput.closest(".field").hidden = true;
    if (el.workflowDetailResearchBtn?.closest(".action-row")) el.workflowDetailResearchBtn.closest(".action-row").hidden = true;
    if (el.workflowDetailResearchStatus) el.workflowDetailResearchStatus.hidden = true;
    if (el.workflowDetailResearchList) el.workflowDetailResearchList.hidden = true;
    if (el.workflowDetailResearchApplyBtn) el.workflowDetailResearchApplyBtn.hidden = true;
    if (el.workflowDetailTheme?.closest(".theme-review-card")) el.workflowDetailTheme.closest(".theme-review-card").hidden = true;
    if (el.workflowDetailPrompt?.closest(".theme-review-card")) el.workflowDetailPrompt.closest(".theme-review-card").hidden = true;

    if (el.workflowDetailRunBtn) {
      el.workflowDetailRunBtn.disabled = page.status === "running" || state.workflowRunning || page.layoutStatus === "running" || hasPendingDraft;
      el.workflowDetailRunBtn.textContent = hasPendingDraft ? "先确认内容" : (page.status === "running" ? "生成中..." : "生成本页");
    }
    if (el.workflowDetailConfirmContentBtn) {
      el.workflowDetailConfirmContentBtn.disabled = !hasPendingDraft;
    }
    if (shell.researchToggleBtn) {
      shell.researchToggleBtn.disabled = page.researchStatus === "running";
      shell.researchToggleBtn.textContent = page.researchStatus === "running" ? "联网补充中..." : "联网补充";
    }

    renderWorkflowResearchPopover(page);
  };

  if (el.workflowDetailTitle) el.workflowDetailTitle.textContent = `第 ${page.pageNumber} 页详情`;
  if (el.workflowDetailDisplayTitleInput) {
    el.workflowDetailDisplayTitleInput.value = draft.title;
    el.workflowDetailDisplayTitleInput.oninput = () => {
      updateWorkflowDetailDraft(page, { title: el.workflowDetailDisplayTitleInput.value });
      updateDetailDerivedContent();
    };
  }
  if (el.workflowDetailVisibleTextEditor) {
    el.workflowDetailVisibleTextEditor.value = draft.body;
    el.workflowDetailVisibleTextEditor.oninput = () => {
      updateWorkflowDetailDraft(page, { body: el.workflowDetailVisibleTextEditor.value });
      updateDetailDerivedContent();
    };
  }
  if (el.workflowDetailConfirmContentBtn) {
    el.workflowDetailConfirmContentBtn.onclick = () => {
      const suggestedPlan = buildWorkflowSuggestedTextPlan(page);
      const suggestedBody = suggestedPlan.visibleLines.join("\n").trim();
      const nextTitle = draft.title.trim();
      const nextBody = draft.body.trim();
      page.displayTitleOverride = nextTitle === String(page.pageTitle || "").trim() ? null : draft.title;
      page.displayBodyOverride = nextBody === suggestedBody ? null : draft.body;
      primeWorkflowDetailDraft(page);
      saveSettings();
      renderWorkflowPlan();
      renderWorkflowDetail();
      updatePayloadPreview();
      setStatus(`第 ${page.pageNumber} 页的上屏内容已确认，之后生成将以这份内容为准。`, "success");
    };
  }

  if (shell.researchToggleBtn) {
    shell.researchToggleBtn.onclick = () => {
      if (state.workflowResearchPopoverOpen && state.workflowResearchPopoverPageId === page.id) {
        closeWorkflowResearchPopover();
      } else {
        openWorkflowResearchPopover(page);
      }
      updateDetailDerivedContent();
    };
  }

  if (el.workflowDetailRunBtn) el.workflowDetailRunBtn.onclick = () => generateWorkflowPage(page.id);
  if (el.workflowDetailCloseBtn) {
    el.workflowDetailCloseBtn.onclick = () => {
      closeWorkflowResearchPopover();
      closeWorkflowDetail();
    };
  }

  updateDetailDerivedContent();
}

function enhanceWorkflowWorkspaceUi() {
  ensureWorkflowDetailWorkspaceShell();
  normalizeWorkflowDetailHeadings();
  ensureWorkflowResearchPopover();
  renderWorkflowPlanHistory();
  renderWorkflowPlan();
  renderWorkflowDetail();
}

enhanceWorkflowWorkspaceUi();

function hideSmartWorkflowChrome() {
  const smartPanel = document.querySelector('[data-main-panel="smart"]');
  const sourceSubpanel = smartPanel?.querySelector('.smart-flow-grid > .subpanel');
  const metaGrid = sourceSubpanel?.querySelector('.grid.two');
  const flowInfoField = metaGrid?.querySelector('.field:not(label)');

  if (sourceSubpanel?.querySelector('.subpanel-head')) {
    sourceSubpanel.querySelector('.subpanel-head').hidden = true;
  }
  if (el.workflowGateHint) {
    el.workflowGateHint.hidden = true;
    el.workflowGateHint.style.display = "none";
  }
  if (el.workflowPlanSummary) {
    el.workflowPlanSummary.hidden = true;
    el.workflowPlanSummary.style.display = "none";
  }
  if (metaGrid) {
    metaGrid.classList.add("workflow-compact-meta");
  }
  if (flowInfoField) {
    flowInfoField.hidden = true;
    flowInfoField.style.display = "none";
  }
}

function hideWorkflowDetailLegacySections() {
  [
    el.workflowDetailSuggestedContent,
    el.workflowDetailConfirmedContent,
    el.workflowDetailContent,
    el.workflowDetailConfirmStatus,
    el.workflowDetailResearchStatus,
    el.workflowDetailResearchList,
    el.workflowDetailResearchApplyBtn,
    el.workflowDetailUseBtn,
    el.workflowDetailCopyBtn,
    el.workflowDetailVisibleResetBtn,
  ].forEach((node) => {
    if (!node) return;
    node.hidden = true;
    node.style.display = "none";
  });

  const queryField = el.workflowDetailResearchQueryInput?.closest(".field");
  if (queryField) {
    queryField.hidden = true;
    queryField.style.display = "none";
  }

  const researchRow = el.workflowDetailResearchBtn?.closest(".action-row");
  if (researchRow) {
    researchRow.hidden = true;
    researchRow.style.display = "none";
  }

  const themeCard = el.workflowDetailTheme?.closest(".theme-review-card");
  if (themeCard) {
    themeCard.hidden = true;
    themeCard.style.display = "none";
  }

  const promptCard = el.workflowDetailPrompt?.closest(".theme-review-card");
  if (promptCard) {
    promptCard.hidden = true;
    promptCard.style.display = "none";
  }
}

const __previousEnsureWorkflowResearchPopover = ensureWorkflowResearchPopover;
ensureWorkflowResearchPopover = function ensureWorkflowResearchPopoverMerged() {
  const refs = __previousEnsureWorkflowResearchPopover();
  const host = document.querySelector("#workflowDetailConfirmContentBtn")?.closest(".workflow-detail-editor-actions")
    || document.querySelector("#workflowDetailDisplayTitleInput")?.closest(".theme-review-card")
    || el.workflowDetailPanel;
  if (host && refs.popover.parentElement !== host) {
    host.appendChild(refs.popover);
  }
  return refs;
};

function renderWorkflowPlanHistory() {
  if (!el.workflowPlanHistoryPanel) return;
  el.workflowPlanHistoryPanel.hidden = true;
  el.workflowPlanHistoryPanel.classList.add("hidden");
  el.workflowPlanHistoryPanel.style.display = "none";
  if (el.workflowPlanDeleteBtn) {
    el.workflowPlanDeleteBtn.hidden = true;
    el.workflowPlanDeleteBtn.style.display = "none";
  }
}

const __previousRenderWorkflowPlanLibrary = renderWorkflowPlanLibrary;
renderWorkflowPlanLibrary = function renderWorkflowPlanLibraryCompact() {
  __previousRenderWorkflowPlanLibrary();
  renderWorkflowPlanHistory();
  hideSmartWorkflowChrome();
};

const __previousRenderWorkflowDetail = renderWorkflowDetail;
renderWorkflowDetail = function renderWorkflowDetailCompact() {
  __previousRenderWorkflowDetail();
  hideWorkflowDetailLegacySections();
  const page = state.workflowPages.find((item) => item.id === state.workflowDetailPageId);
  if (page) {
    renderWorkflowResearchPopover(page);
  }
};

hideSmartWorkflowChrome();
renderWorkflowPlanHistory();
renderWorkflowPlanLibrary();
renderWorkflowDetail();

function normalizeWorkflowDetailHeadings() {
  const detailCard = document.querySelector("#workflowDetailDisplayTitleInput")?.closest(".theme-review-card");
  if (detailCard) {
    const strongs = Array.from(detailCard.querySelectorAll("strong"));
    strongs.forEach((node, index) => {
      node.hidden = index !== 0;
      node.style.display = index === 0 ? "" : "none";
      if (index === 0) node.textContent = "最终上屏内容";
    });
  }

  const layoutCard = document.querySelector("#workflowDetailLayout")?.closest(".theme-review-card");
  if (layoutCard) {
    const strongs = Array.from(layoutCard.querySelectorAll("strong"));
    strongs.forEach((node, index) => {
      node.hidden = index !== 0;
      node.style.display = index === 0 ? "" : "none";
      if (index === 0) node.textContent = "格式安排";
    });
  }
}

const __workflowEnsureResearchPopoverMergedV2 = ensureWorkflowResearchPopover;
ensureWorkflowResearchPopover = function ensureWorkflowResearchPopoverDocked() {
  const refs = __workflowEnsureResearchPopoverMergedV2();
  const host = document.querySelector("#workflowDetailConfirmContentBtn")?.closest(".workflow-detail-editor-actions")
    || document.querySelector("#workflowDetailDisplayTitleInput")?.closest(".theme-review-card")
    || el.workflowDetailPanel;
  if (host && refs.popover.parentElement !== host) {
    host.appendChild(refs.popover);
  }
  return refs;
};

function ensureWorkflowWorkbenchCards() {
  const grid = el.workflowDetailPanel?.querySelector(".workflow-detail-grid");
  const editorCard = document.querySelector("#workflowDetailDisplayTitleInput")?.closest(".theme-review-card");
  const layoutCard = document.querySelector("#workflowDetailLayout")?.closest(".theme-review-card");
  if (!grid || !editorCard || !layoutCard) return {};

  grid.classList.add("workflow-detail-workbench-grid");
  editorCard.classList.add("workflow-detail-editor-card");
  layoutCard.classList.add("workflow-detail-layout-card");

  let boardCard = document.querySelector("#workflowDetailBoardCard");
  if (!boardCard) {
    boardCard = document.createElement("article");
    boardCard.id = "workflowDetailBoardCard";
    boardCard.className = "theme-review-card workflow-detail-board-card";
    boardCard.innerHTML = `
      <div class="workflow-detail-board-shell">
        <aside class="workflow-sketch-toolbar" id="workflowDetailSketchToolbar">
          <button type="button" class="workflow-sketch-tool active" data-tool="select" title="选择">选</button>
          <button type="button" class="workflow-sketch-tool" data-tool="pen" title="画笔">笔</button>
          <button type="button" class="workflow-sketch-tool" data-tool="eraser" title="橡皮">擦</button>
          <button type="button" class="workflow-sketch-tool" data-tool="line" title="直线">线</button>
          <button type="button" class="workflow-sketch-tool" data-tool="rect" title="矩形框选">框</button>
          <button type="button" class="workflow-sketch-tool" data-tool="circle" title="圆形">圆</button>
          <button type="button" class="workflow-sketch-tool" data-tool="arrow" title="箭头">箭</button>
          <div class="workflow-sketch-toolbar-divider"></div>
          <label class="workflow-sketch-color" title="颜色">
            <input id="workflowDetailSketchColor" type="color" value="#498094" />
          </label>
          <input id="workflowDetailSketchSize" type="range" min="1" max="10" value="3" title="线宽" />
          <div class="workflow-sketch-toolbar-divider"></div>
          <button type="button" class="workflow-sketch-tool" data-action="undo" title="撤销">撤</button>
          <button type="button" class="workflow-sketch-tool" data-action="redo" title="重做">返</button>
          <button type="button" class="workflow-sketch-tool" data-action="clear" title="清空">清</button>
        </aside>
        <div class="workflow-detail-board-stage">
          <img id="workflowDetailSketchBackdrop" class="workflow-detail-sketch-backdrop hidden" alt="当前页结果图" />
          <canvas id="workflowDetailSketchCanvas" class="workflow-detail-sketch-canvas"></canvas>
        </div>
      </div>
      <div class="workflow-detail-board-actions">
        <button id="workflowDetailBoardGenerateBtn" type="button" class="primary-btn">生成本页</button>
        <button id="workflowDetailBoardReviseBtn" type="button" class="ghost-btn">打开改图</button>
      </div>
    `;
    grid.insertBefore(boardCard, editorCard);
  }

  let notesCard = document.querySelector("#workflowDetailNotesCard");
  if (!notesCard) {
    notesCard = document.createElement("article");
    notesCard.id = "workflowDetailNotesCard";
    notesCard.className = "theme-review-card workflow-detail-notes-card";
    notesCard.innerHTML = `
      <strong>额外要求 / 修改要求</strong>
      <label class="field workflow-detail-field">
        <textarea id="workflowDetailWorkspaceNotes" rows="5" class="workflow-detail-notes-editor" placeholder="例如：右上角补一个参数卡片；把圆形改成地球；整体更像信息图；保留一块区域给图表。"></textarea>
      </label>
    `;
    grid.insertBefore(notesCard, layoutCard);
  }

  let layoutScroller = document.querySelector("#workflowDetailLayoutScroller");
  if (!layoutScroller) {
    layoutScroller = document.createElement("div");
    layoutScroller.id = "workflowDetailLayoutScroller";
    layoutScroller.className = "workflow-layout-scroller";
    layoutCard.appendChild(layoutScroller);
  }

  if (el.workflowDetailLayout) {
    el.workflowDetailLayout.hidden = true;
    el.workflowDetailLayout.style.display = "none";
  }

  return { grid, editorCard, layoutCard, boardCard, notesCard, layoutScroller };
}

function renderWorkflowSketchCanvas(page) {
  const canvas = document.querySelector("#workflowDetailSketchCanvas");
  if (!canvas || !page) return;
  const stage = canvas.parentElement;
  const backdrop = document.querySelector("#workflowDetailSketchBackdrop");
  const displayWidth = Math.max(400, stage.clientWidth || 640);
  const displayHeight = Math.round(displayWidth * 9 / 16);
  const dpr = window.devicePixelRatio || 1;

  canvas.width = Math.round(displayWidth * dpr);
  canvas.height = Math.round(displayHeight * dpr);
  canvas.style.width = `${displayWidth}px`;
  canvas.style.height = `${displayHeight}px`;

  if (backdrop) {
    const imageUrl = Array.isArray(page.resultImages) && page.resultImages.length ? page.resultImages[0] : "";
    if (imageUrl) {
      backdrop.src = imageUrl;
      backdrop.hidden = false;
      backdrop.classList.remove("hidden");
    } else {
      backdrop.hidden = true;
      backdrop.classList.add("hidden");
      backdrop.removeAttribute("src");
    }
  }

  const ctx = canvas.getContext("2d");
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, displayWidth, displayHeight);
  ctx.fillStyle = "rgba(252, 253, 255, 0.82)";
  ctx.fillRect(0, 0, displayWidth, displayHeight);
  ctx.strokeStyle = "rgba(19, 34, 56, 0.08)";
  ctx.lineWidth = 1;

  for (let x = 0; x <= displayWidth; x += displayWidth / 8) {
    ctx.beginPath();
    ctx.moveTo(x, 0);
    ctx.lineTo(x, displayHeight);
    ctx.stroke();
  }
  for (let y = 0; y <= displayHeight; y += displayHeight / 4.5) {
    ctx.beginPath();
    ctx.moveTo(0, y);
    ctx.lineTo(displayWidth, y);
    ctx.stroke();
  }

  ctx.save();
  ctx.scale(displayWidth, displayHeight);
  normalizeWorkflowSketchShapes(page.detailSketchShapes).forEach((shape) => drawWorkflowSketchShape(ctx, shape));
  if (workflowDetailSketchUi.pageId === page.id && workflowDetailSketchUi.draftShape) {
    drawWorkflowSketchShape(ctx, workflowDetailSketchUi.draftShape);
  }
  ctx.restore();
}

function setupWorkflowSketchInteractions(page) {
  const canvas = document.querySelector("#workflowDetailSketchCanvas");
  const toolbar = document.querySelector("#workflowDetailSketchToolbar");
  const colorInput = document.querySelector("#workflowDetailSketchColor");
  const sizeInput = document.querySelector("#workflowDetailSketchSize");
  const notesInput = document.querySelector("#workflowDetailWorkspaceNotes");
  const generateBtn = document.querySelector("#workflowDetailBoardGenerateBtn");
  const reviseBtn = document.querySelector("#workflowDetailBoardReviseBtn");

  if (!canvas || !toolbar || !page) return;
  workflowDetailSketchUi.pageId = page.id;
  page.detailSketchShapes = normalizeWorkflowSketchShapes(page.detailSketchShapes);
  page.detailSketchRedo = normalizeWorkflowSketchShapes(page.detailSketchRedo);

  toolbar.querySelectorAll(".workflow-sketch-tool[data-tool]").forEach((button) => {
    button.classList.toggle("active", button.dataset.tool === workflowDetailSketchUi.tool);
    button.onclick = () => {
      workflowDetailSketchUi.tool = button.dataset.tool;
      setupWorkflowSketchInteractions(page);
    };
  });

  toolbar.querySelectorAll(".workflow-sketch-tool[data-action]").forEach((button) => {
    button.onclick = () => {
      if (button.dataset.action === "undo" && page.detailSketchShapes.length) {
        page.detailSketchRedo.push(page.detailSketchShapes.pop());
      }
      if (button.dataset.action === "redo" && page.detailSketchRedo.length) {
        page.detailSketchShapes.push(page.detailSketchRedo.pop());
      }
      if (button.dataset.action === "clear") {
        page.detailSketchRedo = [];
        page.detailSketchShapes = [];
      }
      saveSettings();
      renderWorkflowSketchCanvas(page);
      updatePayloadPreview();
    };
  });

  if (colorInput) {
    colorInput.value = workflowDetailSketchUi.color;
    colorInput.oninput = () => {
      workflowDetailSketchUi.color = colorInput.value;
    };
  }
  if (sizeInput) {
    sizeInput.value = String(workflowDetailSketchUi.size);
    sizeInput.oninput = () => {
      workflowDetailSketchUi.size = Number(sizeInput.value) || 3;
    };
  }
  if (notesInput) {
    notesInput.value = page.detailWorkspaceNotes || "";
    notesInput.oninput = () => {
      page.detailWorkspaceNotes = notesInput.value;
      saveSettings();
      updatePayloadPreview();
    };
  }

  if (generateBtn) {
    generateBtn.disabled = page.status === "running" || state.workflowRunning || page.layoutStatus === "running" || getWorkflowDetailDraft(page).dirty;
    generateBtn.onclick = () => generateWorkflowPage(page.id);
  }

  if (reviseBtn) {
    reviseBtn.disabled = !Array.isArray(page.resultImages) || !page.resultImages.length;
    reviseBtn.onclick = async () => {
      if (!page.resultImages?.length) {
        setStatus("请先生成本页，再打开改图。", "error");
        return;
      }
      state.imagePreview = {
        title: `第 ${page.pageNumber} 页结果`,
        url: page.resultImages[0],
      };
      if (el.imagePreviewPrompt) {
        el.imagePreviewPrompt.value = String(page.detailWorkspaceNotes || "").trim();
      }
      await sendImagePreviewToRevise({ submitImmediately: false });
    };
  }

  canvas.onpointerdown = (event) => {
    const tool = workflowDetailSketchUi.tool;
    if (tool === "select") return;
    const point = getWorkflowSketchCanvasPoint(event, canvas);
    if (tool === "eraser") {
      for (let index = page.detailSketchShapes.length - 1; index >= 0; index -= 1) {
        if (workflowShapeHitTest(page.detailSketchShapes[index], point)) {
          page.detailSketchRedo = [];
          page.detailSketchShapes.splice(index, 1);
          saveSettings();
          renderWorkflowSketchCanvas(page);
          updatePayloadPreview();
          break;
        }
      }
      return;
    }

    workflowDetailSketchUi.drawing = true;
    workflowDetailSketchUi.start = point;
    workflowDetailSketchUi.draftShape = createWorkflowSketchShape(tool, point, point, workflowDetailSketchUi.color, workflowDetailSketchUi.size);
    if (tool === "pen") workflowDetailSketchUi.draftShape.points = [point];
    canvas.setPointerCapture?.(event.pointerId);
    renderWorkflowSketchCanvas(page);
  };

  canvas.onpointermove = (event) => {
    if (!workflowDetailSketchUi.drawing || !workflowDetailSketchUi.draftShape) return;
    const point = getWorkflowSketchCanvasPoint(event, canvas);
    workflowDetailSketchUi.draftShape.x2 = point.x;
    workflowDetailSketchUi.draftShape.y2 = point.y;
    if (workflowDetailSketchUi.draftShape.type === "pen") {
      workflowDetailSketchUi.draftShape.points.push(point);
    }
    renderWorkflowSketchCanvas(page);
  };

  const finishDrawing = () => {
    if (!workflowDetailSketchUi.drawing || !workflowDetailSketchUi.draftShape) return;
    const shape = workflowDetailSketchUi.draftShape;
    workflowDetailSketchUi.drawing = false;
    workflowDetailSketchUi.start = null;
    workflowDetailSketchUi.draftShape = null;
    if (!isWorkflowSketchShapeTiny(shape)) {
      page.detailSketchRedo = [];
      page.detailSketchShapes.push(shape);
      saveSettings();
      updatePayloadPreview();
    }
    renderWorkflowSketchCanvas(page);
  };

  canvas.onpointerup = finishDrawing;
  canvas.onpointerleave = () => {
    if (workflowDetailSketchUi.drawing && workflowDetailSketchUi.tool === "pen") {
      finishDrawing();
      return;
    }
    if (!workflowDetailSketchUi.drawing) return;
    workflowDetailSketchUi.drawing = false;
    workflowDetailSketchUi.start = null;
    workflowDetailSketchUi.draftShape = null;
    renderWorkflowSketchCanvas(page);
  };

  renderWorkflowSketchCanvas(page);
}

function renderWorkflowDetail() {
  const page = state.workflowPages.find((item) => item.id === state.workflowDetailPageId);
  if (!page) {
    closeWorkflowResearchPopover();
    el.workflowDetailModal?.classList.add("hidden");
    el.workflowDetailModal?.setAttribute("aria-hidden", "true");
    return;
  }

  const shell = ensureWorkflowDetailWorkspaceShell();
  ensureWorkflowWorkbenchCards();
  normalizeWorkflowDetailHeadings();
  const draft = getWorkflowDetailDraft(page);

  const updateDetailDerivedContent = () => {
    const hasPendingDraft = getWorkflowDetailDraft(page).dirty;
    el.workflowDetailMeta.textContent = `${PAGE_TYPE_LABELS[page.pageType] || page.pageType} · ${getWorkflowDisplayTitle(page) || page.pageTitle || `第 ${page.pageNumber} 页`} · 输出 ${getPptOutputDescription()}`;
    renderWorkflowLayoutScroller(page);

    if (el.workflowDetailUseBtn) {
      el.workflowDetailUseBtn.hidden = true;
      el.workflowDetailUseBtn.disabled = true;
    }
    if (el.workflowDetailCopyBtn) {
      el.workflowDetailCopyBtn.hidden = true;
      el.workflowDetailCopyBtn.disabled = true;
    }
    if (el.workflowDetailVisibleResetBtn) {
      el.workflowDetailVisibleResetBtn.hidden = true;
      el.workflowDetailVisibleResetBtn.disabled = true;
    }
    if (el.workflowDetailConfirmStatus) el.workflowDetailConfirmStatus.hidden = true;

    setWorkflowDetailSectionVisibility(el.workflowDetailSuggestedContent, false);
    setWorkflowDetailSectionVisibility(el.workflowDetailConfirmedContent, false);
    setWorkflowDetailSectionVisibility(el.workflowDetailContent, false);

    if (el.workflowDetailResearchQueryInput?.closest(".field")) el.workflowDetailResearchQueryInput.closest(".field").hidden = true;
    if (el.workflowDetailResearchBtn?.closest(".action-row")) el.workflowDetailResearchBtn.closest(".action-row").hidden = true;
    if (el.workflowDetailResearchStatus) el.workflowDetailResearchStatus.hidden = true;
    if (el.workflowDetailResearchList) el.workflowDetailResearchList.hidden = true;
    if (el.workflowDetailResearchApplyBtn) el.workflowDetailResearchApplyBtn.hidden = true;
    if (el.workflowDetailTheme?.closest(".theme-review-card")) el.workflowDetailTheme.closest(".theme-review-card").hidden = true;
    if (el.workflowDetailPrompt?.closest(".theme-review-card")) el.workflowDetailPrompt.closest(".theme-review-card").hidden = true;

    if (el.workflowDetailRunBtn) {
      el.workflowDetailRunBtn.hidden = true;
      el.workflowDetailRunBtn.disabled = page.status === "running" || state.workflowRunning || page.layoutStatus === "running" || hasPendingDraft;
      el.workflowDetailRunBtn.textContent = hasPendingDraft ? "先确认内容" : (page.status === "running" ? "生成中..." : "生成本页");
    }
    if (el.workflowDetailConfirmContentBtn) {
      el.workflowDetailConfirmContentBtn.disabled = !hasPendingDraft;
    }
    if (shell.researchToggleBtn) {
      shell.researchToggleBtn.disabled = page.researchStatus === "running";
      shell.researchToggleBtn.textContent = page.researchStatus === "running" ? "联网补充中..." : "联网补充";
    }

    setupWorkflowSketchInteractions(page);
    renderWorkflowResearchPopover(page);
  };

  if (el.workflowDetailTitle) el.workflowDetailTitle.textContent = `第 ${page.pageNumber} 页详情`;
  if (el.workflowDetailDisplayTitleInput) {
    el.workflowDetailDisplayTitleInput.value = draft.title;
    el.workflowDetailDisplayTitleInput.oninput = () => {
      updateWorkflowDetailDraft(page, { title: el.workflowDetailDisplayTitleInput.value });
      updateDetailDerivedContent();
    };
  }
  if (el.workflowDetailVisibleTextEditor) {
    el.workflowDetailVisibleTextEditor.value = draft.body;
    el.workflowDetailVisibleTextEditor.oninput = () => {
      updateWorkflowDetailDraft(page, { body: el.workflowDetailVisibleTextEditor.value });
      updateDetailDerivedContent();
    };
  }
  if (el.workflowDetailConfirmContentBtn) {
    el.workflowDetailConfirmContentBtn.onclick = () => {
      const suggestedPlan = buildWorkflowSuggestedTextPlan(page);
      const suggestedBody = suggestedPlan.visibleLines.join("\n").trim();
      const nextTitle = draft.title.trim();
      const nextBody = draft.body.trim();
      page.displayTitleOverride = nextTitle === String(page.pageTitle || "").trim() ? null : draft.title;
      page.displayBodyOverride = nextBody === suggestedBody ? null : draft.body;
      primeWorkflowDetailDraft(page);
      saveSettings();
      renderWorkflowPlan();
      renderWorkflowDetail();
      updatePayloadPreview();
      setStatus(`第 ${page.pageNumber} 页的上屏内容已确认，之后生成会按这份内容执行。`, "success");
    };
  }

  if (shell.researchToggleBtn) {
    shell.researchToggleBtn.onclick = () => {
      if (state.workflowResearchPopoverOpen && state.workflowResearchPopoverPageId === page.id) {
        closeWorkflowResearchPopover();
      } else {
        openWorkflowResearchPopover(page);
      }
      updateDetailDerivedContent();
    };
  }

  if (el.workflowDetailCloseBtn) {
    el.workflowDetailCloseBtn.onclick = () => {
      closeWorkflowResearchPopover();
      closeWorkflowDetail();
    };
  }

  updateDetailDerivedContent();
}

hideSmartWorkflowChrome();
renderWorkflowPlanHistory();
renderWorkflowPlanLibrary();
renderWorkflowDetail();

const workflowDetailSketchUi = {
  pageId: "",
  tool: "rect",
  color: "#498094",
  size: 3,
  drawing: false,
  start: null,
  draftShape: null,
};

function cloneWorkflowSketchShapes(shapes) {
  if (!Array.isArray(shapes)) return [];
  return shapes.map((shape) => ({
    ...shape,
    points: Array.isArray(shape.points) ? shape.points.map((point) => ({ x: Number(point.x) || 0, y: Number(point.y) || 0 })) : undefined,
  }));
}

function normalizeWorkflowSketchShapes(shapes) {
  return cloneWorkflowSketchShapes(shapes)
    .filter((shape) => shape && typeof shape === "object" && shape.type)
    .map((shape) => ({
      id: String(shape.id || uid()),
      type: String(shape.type || "rect"),
      x1: Number(shape.x1) || 0,
      y1: Number(shape.y1) || 0,
      x2: Number(shape.x2) || 0,
      y2: Number(shape.y2) || 0,
      color: String(shape.color || "#498094"),
      size: Math.max(1, Number(shape.size) || 3),
      points: Array.isArray(shape.points) ? shape.points.map((point) => ({
        x: Number(point.x) || 0,
        y: Number(point.y) || 0,
      })) : [],
    }));
}

const __prevSanitizeWorkflowPageFieldsV2 = sanitizeWorkflowPageFields;
sanitizeWorkflowPageFields = function sanitizeWorkflowPageFieldsWithWorkbench(page) {
  const next = __prevSanitizeWorkflowPageFieldsV2(page);
  if (!next || typeof next !== "object") return next;
  if (!Object.prototype.hasOwnProperty.call(next, "detailWorkspaceNotes")) next.detailWorkspaceNotes = "";
  next.detailWorkspaceNotes = String(next.detailWorkspaceNotes || "").trim();
  next.detailSketchShapes = normalizeWorkflowSketchShapes(next.detailSketchShapes);
  next.detailSketchRedo = normalizeWorkflowSketchShapes(next.detailSketchRedo);
  return next;
};

const __prevSerializeWorkflowPageV2 = serializeWorkflowPage;
serializeWorkflowPage = function serializeWorkflowPageWithWorkbench(page) {
  const next = __prevSerializeWorkflowPageV2(page);
  next.detailWorkspaceNotes = String(page?.detailWorkspaceNotes || "").trim();
  next.detailSketchShapes = normalizeWorkflowSketchShapes(page?.detailSketchShapes);
  return next;
};

const __prevHydrateWorkflowPagesV2 = hydrateWorkflowPages;
hydrateWorkflowPages = function hydrateWorkflowPagesWithWorkbench(savedPages) {
  const pages = __prevHydrateWorkflowPagesV2(savedPages);
  return pages.map((page, index) => {
    const saved = Array.isArray(savedPages) ? (savedPages[index] || {}) : {};
    page.detailWorkspaceNotes = String(saved.detailWorkspaceNotes || saved.detail_workspace_notes || "").trim();
    page.detailSketchShapes = normalizeWorkflowSketchShapes(saved.detailSketchShapes || saved.detail_sketch_shapes);
    page.detailSketchRedo = [];
    return sanitizeWorkflowPageFields(page);
  });
};

function summarizeWorkflowSketchForPrompt(page) {
  const shapes = normalizeWorkflowSketchShapes(page?.detailSketchShapes);
  if (!shapes.length) return "";

  const lines = [
    "手工画板补充：以下图形表示版式意图或改造意图，请尽量遵守。",
    "如果用户额外要求解释了某个图形的含义，就按要求理解；没有说明时，把它们当作布局区域、焦点位置或需要转换的视觉形状。",
  ];

  shapes.slice(0, 12).forEach((shape, index) => {
    const x1 = Math.round(Math.min(shape.x1, shape.x2) * 100);
    const y1 = Math.round(Math.min(shape.y1, shape.y2) * 100);
    const x2 = Math.round(Math.max(shape.x1, shape.x2) * 100);
    const y2 = Math.round(Math.max(shape.y1, shape.y2) * 100);
    const typeLabel = ({
      rect: "矩形框",
      circle: "圆形",
      line: "直线",
      arrow: "箭头",
      pen: "自由线条",
    })[shape.type] || "标记";
    lines.push(`${index + 1}. ${typeLabel}：约位于画板 ${x1}%,${y1}% 到 ${x2}%,${y2}% 的范围。`);
  });

  return lines.join("\n");
}

const __prevBuildEffectiveWorkflowPagePromptV2 = buildEffectiveWorkflowPagePrompt;
buildEffectiveWorkflowPagePrompt = function buildEffectiveWorkflowPagePromptWithWorkbench(page) {
  const base = __prevBuildEffectiveWorkflowPagePromptV2(page);
  const notes = String(page?.detailWorkspaceNotes || "").trim();
  const sketch = summarizeWorkflowSketchForPrompt(page);
  const extraBlocks = [];
  if (notes) extraBlocks.push(`用户额外要求：\n${notes}`);
  if (sketch) extraBlocks.push(sketch);
  return [base, ...extraBlocks].filter(Boolean).join("\n\n");
};

function buildWorkflowLayoutSections(page) {
  const next = applyHarnessMetaToPage({ ...page });
  return [
    {
      title: "页面设置",
      summary: `${PAGE_TYPE_LABELS[next.pageType] || next.pageType} · ${getPptOutputDescription()}`,
      body: [`页面密度：${next.densityBand}`, `版式风险：${next.layoutRisk}`].join("\n"),
    },
    {
      title: "布局结构",
      summary: next.layoutSummary || "等待版式整理",
      body: next.layoutSummary || "这一页还没有生成单独的版式安排。拆分页完成后，系统会自动补做这一页的版式整理。",
    },
    {
      title: "文字层级",
      summary: next.textHierarchy || "默认标题优先、正文次级",
      body: next.textHierarchy || "默认标题更大、正文保持远距离可读，关键数字和结论优先放大。",
    },
    {
      title: "视觉焦点",
      summary: next.visualFocus || "等待 AI 版式建议",
      body: next.visualFocus || "当前还没有单独的视觉焦点描述。",
    },
    {
      title: "可读性",
      summary: next.readabilityNotes || "默认大标题 + 稀疏正文",
      body: next.readabilityNotes || "默认要求：标题更大、正文不要密集小字、关键结论和数字优先放大。",
    },
  ];
}

function ensureWorkflowWorkbenchCards() {
  const grid = el.workflowDetailPanel?.querySelector(".workflow-detail-grid");
  const editorCard = document.querySelector("#workflowDetailDisplayTitleInput")?.closest(".theme-review-card");
  const layoutCard = document.querySelector("#workflowDetailLayout")?.closest(".theme-review-card");
  if (!grid || !editorCard || !layoutCard) return {};

  grid.classList.add("workflow-detail-workbench-grid");
  editorCard.classList.add("workflow-detail-editor-card");
  layoutCard.classList.add("workflow-detail-layout-card");

  let boardCard = document.querySelector("#workflowDetailBoardCard");
  if (!boardCard) {
    boardCard = document.createElement("article");
    boardCard.id = "workflowDetailBoardCard";
    boardCard.className = "theme-review-card workflow-detail-board-card";
    boardCard.innerHTML = `
      <div class="workflow-detail-board-shell">
        <aside class="workflow-sketch-toolbar" id="workflowDetailSketchToolbar">
          <button type="button" class="workflow-sketch-tool active" data-tool="select" title="选择">↖</button>
          <button type="button" class="workflow-sketch-tool" data-tool="pen" title="画笔">✎</button>
          <button type="button" class="workflow-sketch-tool" data-tool="eraser" title="橡皮">⌫</button>
          <button type="button" class="workflow-sketch-tool" data-tool="line" title="直线">／</button>
          <button type="button" class="workflow-sketch-tool" data-tool="rect" title="矩形">▭</button>
          <button type="button" class="workflow-sketch-tool" data-tool="circle" title="圆形">◯</button>
          <button type="button" class="workflow-sketch-tool" data-tool="arrow" title="箭头">→</button>
          <div class="workflow-sketch-toolbar-divider"></div>
          <label class="workflow-sketch-color">
            <input id="workflowDetailSketchColor" type="color" value="#498094" />
          </label>
          <input id="workflowDetailSketchSize" type="range" min="1" max="10" value="3" />
          <div class="workflow-sketch-toolbar-divider"></div>
          <button type="button" class="workflow-sketch-tool" data-action="undo" title="撤销">↶</button>
          <button type="button" class="workflow-sketch-tool" data-action="redo" title="重做">↷</button>
          <button type="button" class="workflow-sketch-tool" data-action="clear" title="清空">✕</button>
        </aside>
        <div class="workflow-detail-board-stage">
          <canvas id="workflowDetailSketchCanvas" class="workflow-detail-sketch-canvas"></canvas>
        </div>
      </div>
      <div class="workflow-detail-board-actions">
        <button id="workflowDetailBoardGenerateBtn" type="button" class="primary-btn">生成本页</button>
        <button id="workflowDetailBoardReviseBtn" type="button" class="ghost-btn">应用修改</button>
      </div>
    `;
    grid.insertBefore(boardCard, editorCard);
  }

  let notesCard = document.querySelector("#workflowDetailNotesCard");
  if (!notesCard) {
    notesCard = document.createElement("article");
    notesCard.id = "workflowDetailNotesCard";
    notesCard.className = "theme-review-card workflow-detail-notes-card";
    notesCard.innerHTML = `
      <strong>额外要求</strong>
      <label class="field workflow-detail-field">
        <textarea id="workflowDetailWorkspaceNotes" rows="5" class="workflow-detail-notes-editor" placeholder="例如：这个矩形框表示右上角要放一个技术卡片；圆形改造成地球；整体更像信息图。"></textarea>
      </label>
    `;
    grid.insertBefore(notesCard, layoutCard);
  }

  let layoutScroller = document.querySelector("#workflowDetailLayoutScroller");
  if (!layoutScroller) {
    layoutScroller = document.createElement("div");
    layoutScroller.id = "workflowDetailLayoutScroller";
    layoutScroller.className = "workflow-layout-scroller";
    layoutCard.appendChild(layoutScroller);
  }

  if (el.workflowDetailLayout) {
    el.workflowDetailLayout.hidden = true;
    el.workflowDetailLayout.style.display = "none";
  }

  return { grid, editorCard, layoutCard, boardCard, notesCard, layoutScroller };
}

function renderWorkflowLayoutScroller(page) {
  const refs = ensureWorkflowWorkbenchCards();
  if (!refs.layoutScroller) return;
  const sections = buildWorkflowLayoutSections(page);
  refs.layoutScroller.innerHTML = sections.map((section, index) => `
    <details class="workflow-layout-section" ${index === 0 ? "open" : ""}>
      <summary>
        <span>${escapeHtml(section.title)}</span>
        <small>${escapeHtml(section.summary)}</small>
      </summary>
      <div class="workflow-layout-section-body">${escapeHtml(section.body || "").replace(/\n/g, "<br />")}</div>
    </details>
  `).join("");
}

function getWorkflowSketchCanvasPoint(event, canvas) {
  const rect = canvas.getBoundingClientRect();
  const x = (event.clientX - rect.left) / rect.width;
  const y = (event.clientY - rect.top) / rect.height;
  return {
    x: Math.max(0, Math.min(1, x)),
    y: Math.max(0, Math.min(1, y)),
  };
}

function createWorkflowSketchShape(tool, start, end, color, size) {
  if (tool === "pen") {
    return {
      id: uid(),
      type: "pen",
      color,
      size,
      x1: start.x,
      y1: start.y,
      x2: end.x,
      y2: end.y,
      points: [start, end],
    };
  }
  return {
    id: uid(),
    type: tool,
    color,
    size,
    x1: start.x,
    y1: start.y,
    x2: end.x,
    y2: end.y,
    points: [],
  };
}

function isWorkflowSketchShapeTiny(shape) {
  if (!shape) return true;
  if (shape.type === "pen") return !Array.isArray(shape.points) || shape.points.length < 2;
  return Math.abs(shape.x2 - shape.x1) < 0.01 && Math.abs(shape.y2 - shape.y1) < 0.01;
}

function drawWorkflowSketchShape(ctx, shape) {
  ctx.save();
  ctx.strokeStyle = shape.color || "#498094";
  ctx.lineWidth = Math.max(1, Number(shape.size) || 3);
  ctx.lineCap = "round";
  ctx.lineJoin = "round";

  if (shape.type === "pen" && Array.isArray(shape.points) && shape.points.length > 1) {
    ctx.beginPath();
    ctx.moveTo(shape.points[0].x, shape.points[0].y);
    shape.points.slice(1).forEach((point) => ctx.lineTo(point.x, point.y));
    ctx.stroke();
    ctx.restore();
    return;
  }

  if (shape.type === "rect") {
    const x = Math.min(shape.x1, shape.x2);
    const y = Math.min(shape.y1, shape.y2);
    const w = Math.abs(shape.x2 - shape.x1);
    const h = Math.abs(shape.y2 - shape.y1);
    ctx.strokeRect(x, y, w, h);
    ctx.restore();
    return;
  }

  if (shape.type === "circle") {
    const cx = (shape.x1 + shape.x2) / 2;
    const cy = (shape.y1 + shape.y2) / 2;
    const rx = Math.abs(shape.x2 - shape.x1) / 2;
    const ry = Math.abs(shape.y2 - shape.y1) / 2;
    ctx.beginPath();
    ctx.ellipse(cx, cy, rx, ry, 0, 0, Math.PI * 2);
    ctx.stroke();
    ctx.restore();
    return;
  }

  ctx.beginPath();
  ctx.moveTo(shape.x1, shape.y1);
  ctx.lineTo(shape.x2, shape.y2);
  ctx.stroke();

  if (shape.type === "arrow") {
    const angle = Math.atan2(shape.y2 - shape.y1, shape.x2 - shape.x1);
    const arrowLen = 0.018;
    ctx.beginPath();
    ctx.moveTo(shape.x2, shape.y2);
    ctx.lineTo(shape.x2 - arrowLen * Math.cos(angle - Math.PI / 6), shape.y2 - arrowLen * Math.sin(angle - Math.PI / 6));
    ctx.moveTo(shape.x2, shape.y2);
    ctx.lineTo(shape.x2 - arrowLen * Math.cos(angle + Math.PI / 6), shape.y2 - arrowLen * Math.sin(angle + Math.PI / 6));
    ctx.stroke();
  }
  ctx.restore();
}

function renderWorkflowSketchCanvas(page) {
  const canvas = document.querySelector("#workflowDetailSketchCanvas");
  if (!canvas || !page) return;
  const stage = canvas.parentElement;
  const displayWidth = Math.max(400, stage.clientWidth || 640);
  const displayHeight = Math.round(displayWidth * 9 / 16);
  const dpr = window.devicePixelRatio || 1;

  canvas.width = Math.round(displayWidth * dpr);
  canvas.height = Math.round(displayHeight * dpr);
  canvas.style.width = `${displayWidth}px`;
  canvas.style.height = `${displayHeight}px`;

  const ctx = canvas.getContext("2d");
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, displayWidth, displayHeight);
  ctx.fillStyle = "#fcfdff";
  ctx.fillRect(0, 0, displayWidth, displayHeight);
  ctx.strokeStyle = "rgba(19, 34, 56, 0.08)";
  ctx.lineWidth = 1;

  for (let x = 0; x <= displayWidth; x += displayWidth / 8) {
    ctx.beginPath();
    ctx.moveTo(x, 0);
    ctx.lineTo(x, displayHeight);
    ctx.stroke();
  }
  for (let y = 0; y <= displayHeight; y += displayHeight / 4.5) {
    ctx.beginPath();
    ctx.moveTo(0, y);
    ctx.lineTo(displayWidth, y);
    ctx.stroke();
  }

  ctx.save();
  ctx.scale(displayWidth, displayHeight);
  normalizeWorkflowSketchShapes(page.detailSketchShapes).forEach((shape) => drawWorkflowSketchShape(ctx, shape));
  if (workflowDetailSketchUi.pageId === page.id && workflowDetailSketchUi.draftShape) {
    drawWorkflowSketchShape(ctx, workflowDetailSketchUi.draftShape);
  }
  ctx.restore();
}

function workflowShapeHitTest(shape, point) {
  const left = Math.min(shape.x1, shape.x2) - 0.02;
  const top = Math.min(shape.y1, shape.y2) - 0.02;
  const right = Math.max(shape.x1, shape.x2) + 0.02;
  const bottom = Math.max(shape.y1, shape.y2) + 0.02;
  return point.x >= left && point.x <= right && point.y >= top && point.y <= bottom;
}

function setupWorkflowSketchInteractions(page) {
  const refs = ensureWorkflowWorkbenchCards();
  const canvas = document.querySelector("#workflowDetailSketchCanvas");
  const toolbar = document.querySelector("#workflowDetailSketchToolbar");
  const colorInput = document.querySelector("#workflowDetailSketchColor");
  const sizeInput = document.querySelector("#workflowDetailSketchSize");
  const notesInput = document.querySelector("#workflowDetailWorkspaceNotes");
  const generateBtn = document.querySelector("#workflowDetailBoardGenerateBtn");
  const reviseBtn = document.querySelector("#workflowDetailBoardReviseBtn");

  if (!canvas || !toolbar || !page) return;
  workflowDetailSketchUi.pageId = page.id;
  page.detailSketchShapes = normalizeWorkflowSketchShapes(page.detailSketchShapes);
  page.detailSketchRedo = normalizeWorkflowSketchShapes(page.detailSketchRedo);

  toolbar.querySelectorAll(".workflow-sketch-tool[data-tool]").forEach((button) => {
    button.classList.toggle("active", button.dataset.tool === workflowDetailSketchUi.tool);
    button.onclick = () => {
      workflowDetailSketchUi.tool = button.dataset.tool;
      setupWorkflowSketchInteractions(page);
    };
  });

  toolbar.querySelectorAll(".workflow-sketch-tool[data-action]").forEach((button) => {
    button.onclick = () => {
      if (button.dataset.action === "undo" && page.detailSketchShapes.length) {
        page.detailSketchRedo.push(page.detailSketchShapes.pop());
      }
      if (button.dataset.action === "redo" && page.detailSketchRedo.length) {
        page.detailSketchShapes.push(page.detailSketchRedo.pop());
      }
      if (button.dataset.action === "clear") {
        page.detailSketchRedo = [];
        page.detailSketchShapes = [];
      }
      saveSettings();
      renderWorkflowSketchCanvas(page);
    };
  });

  if (colorInput) {
    colorInput.value = workflowDetailSketchUi.color;
    colorInput.oninput = () => {
      workflowDetailSketchUi.color = colorInput.value;
    };
  }
  if (sizeInput) {
    sizeInput.value = String(workflowDetailSketchUi.size);
    sizeInput.oninput = () => {
      workflowDetailSketchUi.size = Number(sizeInput.value) || 3;
    };
  }
  if (notesInput) {
    notesInput.value = page.detailWorkspaceNotes || "";
    notesInput.oninput = () => {
      page.detailWorkspaceNotes = notesInput.value;
      saveSettings();
      updatePayloadPreview();
    };
  }

  if (generateBtn) {
    generateBtn.disabled = page.status === "running" || state.workflowRunning || page.layoutStatus === "running" || getWorkflowDetailDraft(page).dirty;
    generateBtn.onclick = () => generateWorkflowPage(page.id);
  }

  if (reviseBtn) {
    reviseBtn.disabled = !Array.isArray(page.resultImages) || !page.resultImages.length;
    reviseBtn.onclick = async () => {
      if (!page.resultImages?.length) {
        setStatus("请先生成本页，再把结果送去改图。", "error");
        return;
      }
      state.imagePreview = {
        title: `第 ${page.pageNumber} 页结果`,
        url: page.resultImages[0],
      };
      if (el.imagePreviewPrompt) {
        el.imagePreviewPrompt.value = String(page.detailWorkspaceNotes || "").trim();
      }
      await sendImagePreviewToRevise({ submitImmediately: false });
    };
  }

  canvas.onpointerdown = (event) => {
    const tool = workflowDetailSketchUi.tool;
    if (tool === "select") return;
    const point = getWorkflowSketchCanvasPoint(event, canvas);
    if (tool === "eraser") {
      for (let index = page.detailSketchShapes.length - 1; index >= 0; index -= 1) {
        if (workflowShapeHitTest(page.detailSketchShapes[index], point)) {
          page.detailSketchRedo = [];
          page.detailSketchShapes.splice(index, 1);
          saveSettings();
          renderWorkflowSketchCanvas(page);
          updatePayloadPreview();
          break;
        }
      }
      return;
    }

    workflowDetailSketchUi.drawing = true;
    workflowDetailSketchUi.start = point;
    workflowDetailSketchUi.draftShape = createWorkflowSketchShape(tool, point, point, workflowDetailSketchUi.color, workflowDetailSketchUi.size);
    if (tool === "pen") {
      workflowDetailSketchUi.draftShape.points = [point];
    }
    canvas.setPointerCapture?.(event.pointerId);
    renderWorkflowSketchCanvas(page);
  };

  canvas.onpointermove = (event) => {
    if (!workflowDetailSketchUi.drawing || !workflowDetailSketchUi.draftShape) return;
    const point = getWorkflowSketchCanvasPoint(event, canvas);
    const shape = workflowDetailSketchUi.draftShape;
    shape.x2 = point.x;
    shape.y2 = point.y;
    if (shape.type === "pen") {
      shape.points.push(point);
    }
    renderWorkflowSketchCanvas(page);
  };

  const finishDrawing = () => {
    if (!workflowDetailSketchUi.drawing || !workflowDetailSketchUi.draftShape) return;
    const shape = workflowDetailSketchUi.draftShape;
    workflowDetailSketchUi.drawing = false;
    workflowDetailSketchUi.start = null;
    workflowDetailSketchUi.draftShape = null;
    if (!isWorkflowSketchShapeTiny(shape)) {
      page.detailSketchRedo = [];
      page.detailSketchShapes.push(shape);
      saveSettings();
      updatePayloadPreview();
    }
    renderWorkflowSketchCanvas(page);
  };

  canvas.onpointerup = finishDrawing;
  canvas.onpointerleave = () => {
    if (workflowDetailSketchUi.drawing && workflowDetailSketchUi.tool === "pen") {
      finishDrawing();
      return;
    }
    if (!workflowDetailSketchUi.drawing) return;
    workflowDetailSketchUi.drawing = false;
    workflowDetailSketchUi.start = null;
    workflowDetailSketchUi.draftShape = null;
    renderWorkflowSketchCanvas(page);
  };

  renderWorkflowSketchCanvas(page);
}

function renderWorkflowDetail() {
  const page = state.workflowPages.find((item) => item.id === state.workflowDetailPageId);
  if (!page) {
    closeWorkflowResearchPopover();
    el.workflowDetailModal?.classList.add("hidden");
    el.workflowDetailModal?.setAttribute("aria-hidden", "true");
    return;
  }

  const shell = ensureWorkflowDetailWorkspaceShell();
  ensureWorkflowWorkbenchCards();
  normalizeWorkflowDetailHeadings();
  const draft = getWorkflowDetailDraft(page);

  const updateDetailDerivedContent = () => {
    const hasPendingDraft = getWorkflowDetailDraft(page).dirty;

    el.workflowDetailMeta.textContent = `${PAGE_TYPE_LABELS[page.pageType] || page.pageType} · ${getWorkflowDisplayTitle(page) || page.pageTitle || `第 ${page.pageNumber} 页`} · 输出 ${getPptOutputDescription()}`;
    renderWorkflowLayoutScroller(page);

    if (el.workflowDetailUseBtn) {
      el.workflowDetailUseBtn.hidden = true;
      el.workflowDetailUseBtn.disabled = true;
    }
    if (el.workflowDetailCopyBtn) {
      el.workflowDetailCopyBtn.hidden = true;
      el.workflowDetailCopyBtn.disabled = true;
    }
    if (el.workflowDetailVisibleResetBtn) {
      el.workflowDetailVisibleResetBtn.hidden = true;
      el.workflowDetailVisibleResetBtn.disabled = true;
    }
    if (el.workflowDetailConfirmStatus) {
      el.workflowDetailConfirmStatus.hidden = true;
    }

    setWorkflowDetailSectionVisibility(el.workflowDetailSuggestedContent, false);
    setWorkflowDetailSectionVisibility(el.workflowDetailConfirmedContent, false);
    setWorkflowDetailSectionVisibility(el.workflowDetailContent, false);

    if (el.workflowDetailResearchQueryInput?.closest(".field")) el.workflowDetailResearchQueryInput.closest(".field").hidden = true;
    if (el.workflowDetailResearchBtn?.closest(".action-row")) el.workflowDetailResearchBtn.closest(".action-row").hidden = true;
    if (el.workflowDetailResearchStatus) el.workflowDetailResearchStatus.hidden = true;
    if (el.workflowDetailResearchList) el.workflowDetailResearchList.hidden = true;
    if (el.workflowDetailResearchApplyBtn) el.workflowDetailResearchApplyBtn.hidden = true;
    if (el.workflowDetailTheme?.closest(".theme-review-card")) el.workflowDetailTheme.closest(".theme-review-card").hidden = true;
    if (el.workflowDetailPrompt?.closest(".theme-review-card")) el.workflowDetailPrompt.closest(".theme-review-card").hidden = true;

    if (el.workflowDetailRunBtn) {
      el.workflowDetailRunBtn.hidden = true;
      el.workflowDetailRunBtn.disabled = page.status === "running" || state.workflowRunning || page.layoutStatus === "running" || hasPendingDraft;
      el.workflowDetailRunBtn.textContent = hasPendingDraft ? "先确认内容" : (page.status === "running" ? "生成中..." : "生成本页");
    }
    if (el.workflowDetailConfirmContentBtn) {
      el.workflowDetailConfirmContentBtn.disabled = !hasPendingDraft;
    }
    if (shell.researchToggleBtn) {
      shell.researchToggleBtn.disabled = page.researchStatus === "running";
      shell.researchToggleBtn.textContent = page.researchStatus === "running" ? "联网补充中..." : "联网补充";
    }

    setupWorkflowSketchInteractions(page);
    renderWorkflowResearchPopover(page);
  };

  if (el.workflowDetailTitle) el.workflowDetailTitle.textContent = `第 ${page.pageNumber} 页详情`;
  if (el.workflowDetailDisplayTitleInput) {
    el.workflowDetailDisplayTitleInput.value = draft.title;
    el.workflowDetailDisplayTitleInput.oninput = () => {
      updateWorkflowDetailDraft(page, { title: el.workflowDetailDisplayTitleInput.value });
      updateDetailDerivedContent();
    };
  }
  if (el.workflowDetailVisibleTextEditor) {
    el.workflowDetailVisibleTextEditor.value = draft.body;
    el.workflowDetailVisibleTextEditor.oninput = () => {
      updateWorkflowDetailDraft(page, { body: el.workflowDetailVisibleTextEditor.value });
      updateDetailDerivedContent();
    };
  }
  if (el.workflowDetailConfirmContentBtn) {
    el.workflowDetailConfirmContentBtn.onclick = () => {
      const suggestedPlan = buildWorkflowSuggestedTextPlan(page);
      const suggestedBody = suggestedPlan.visibleLines.join("\n").trim();
      const nextTitle = draft.title.trim();
      const nextBody = draft.body.trim();
      page.displayTitleOverride = nextTitle === String(page.pageTitle || "").trim() ? null : draft.title;
      page.displayBodyOverride = nextBody === suggestedBody ? null : draft.body;
      primeWorkflowDetailDraft(page);
      saveSettings();
      renderWorkflowPlan();
      renderWorkflowDetail();
      updatePayloadPreview();
      setStatus(`第 ${page.pageNumber} 页的上屏内容已确认，之后生成将以这份内容为准。`, "success");
    };
  }

  if (shell.researchToggleBtn) {
    shell.researchToggleBtn.onclick = () => {
      if (state.workflowResearchPopoverOpen && state.workflowResearchPopoverPageId === page.id) {
        closeWorkflowResearchPopover();
      } else {
        openWorkflowResearchPopover(page);
      }
      updateDetailDerivedContent();
    };
  }

  if (el.workflowDetailCloseBtn) {
    el.workflowDetailCloseBtn.onclick = () => {
      closeWorkflowResearchPopover();
      closeWorkflowDetail();
    };
  }

  updateDetailDerivedContent();
}

renderWorkflowDetail();
