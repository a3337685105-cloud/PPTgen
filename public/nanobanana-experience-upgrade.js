(function bootstrapNanoBananaExperienceUpgrade() {
  const SURVEY_STORAGE_KEY = "pptgen.preferenceSurvey.v1";
  const PREFERENCE_MARKER = "【用户排版偏好】";
  const DEFAULT_PREFERENCES = {
    layoutVariety: "balanced",
    detailLevel: "polished",
    visualDensity: "balanced",
    compositionFocus: "balanced",
    dataNarrative: "balanced",
    pageMood: "modern",
  };
  const PREFERENCE_OPTIONS = {
    layoutVariety: ["uniform", "balanced", "diverse"],
    detailLevel: ["minimal", "polished", "rich"],
    visualDensity: ["airy", "balanced", "dense"],
    compositionFocus: ["imageLead", "balanced", "textLead"],
    dataNarrative: ["clean", "balanced", "expressive"],
    pageMood: ["steady", "modern", "dramatic"],
  };
  const PREFERENCE_LABELS = {
    layoutVariety: {
      uniform: "统一稳定",
      balanced: "平衡变化",
      diverse: "尽量多样",
    },
    detailLevel: {
      minimal: "偏简约",
      polished: "精致平衡",
      rich: "偏精细",
    },
    visualDensity: {
      airy: "留白更多",
      balanced: "均衡",
      dense: "信息更满",
    },
    compositionFocus: {
      imageLead: "视觉主导",
      balanced: "图文平衡",
      textLead: "内容主导",
    },
    dataNarrative: {
      clean: "清晰克制",
      balanced: "适度信息图",
      expressive: "视觉化强调",
    },
    pageMood: {
      steady: "稳重统一",
      modern: "现代清爽",
      dramatic: "更有冲击",
    },
  };

  const workflowProgressState = {
    lastMessage: "",
    lastTone: "idle",
  };

  const reviseBboxState = {
    syncing: false,
    manualOptOut: false,
    announcedReady: false,
  };

  function normalizePreferenceValue(key, value) {
    const text = String(value || "").trim();
    return PREFERENCE_OPTIONS[key]?.includes(text) ? text : DEFAULT_PREFERENCES[key];
  }

  function readStoredPreferences() {
    try {
      const parsed = JSON.parse(localStorage.getItem(SURVEY_STORAGE_KEY) || "{}");
      return {
        layoutVariety: normalizePreferenceValue("layoutVariety", parsed.layoutVariety),
        detailLevel: normalizePreferenceValue("detailLevel", parsed.detailLevel),
        visualDensity: normalizePreferenceValue("visualDensity", parsed.visualDensity),
        compositionFocus: normalizePreferenceValue("compositionFocus", parsed.compositionFocus),
        dataNarrative: normalizePreferenceValue("dataNarrative", parsed.dataNarrative),
        pageMood: normalizePreferenceValue("pageMood", parsed.pageMood),
      };
    } catch {
      return { ...DEFAULT_PREFERENCES };
    }
  }

  function getPreferenceNodes() {
    return {
      layoutVariety: document.getElementById("prefLayoutVariety"),
      detailLevel: document.getElementById("prefDetailLevel"),
      visualDensity: document.getElementById("prefVisualDensity"),
      compositionFocus: document.getElementById("prefCompositionFocus"),
      dataNarrative: document.getElementById("prefDataNarrative"),
      pageMood: document.getElementById("prefPageMood"),
      summary: document.getElementById("pptPreferenceSummary"),
    };
  }

  function getCurrentPreferences() {
    const nodes = getPreferenceNodes();
    return {
      layoutVariety: normalizePreferenceValue("layoutVariety", nodes.layoutVariety?.value),
      detailLevel: normalizePreferenceValue("detailLevel", nodes.detailLevel?.value),
      visualDensity: normalizePreferenceValue("visualDensity", nodes.visualDensity?.value),
      compositionFocus: normalizePreferenceValue("compositionFocus", nodes.compositionFocus?.value),
      dataNarrative: normalizePreferenceValue("dataNarrative", nodes.dataNarrative?.value),
      pageMood: normalizePreferenceValue("pageMood", nodes.pageMood?.value),
    };
  }

  function getPreferenceLabel(key, value) {
    return PREFERENCE_LABELS[key]?.[value] || value;
  }

  function applyPreferencesToUi(preferences) {
    const nodes = getPreferenceNodes();
    Object.keys(DEFAULT_PREFERENCES).forEach((key) => {
      if (nodes[key]) nodes[key].value = normalizePreferenceValue(key, preferences[key]);
    });
    renderPreferenceSummary();
  }

  function renderPreferenceSummary() {
    const nodes = getPreferenceNodes();
    if (!nodes.summary) return;
    const preferences = getCurrentPreferences();
    nodes.summary.textContent = [
      `当前偏好：${getPreferenceLabel("layoutVariety", preferences.layoutVariety)}`,
      getPreferenceLabel("detailLevel", preferences.detailLevel),
      getPreferenceLabel("visualDensity", preferences.visualDensity),
      getPreferenceLabel("compositionFocus", preferences.compositionFocus),
      getPreferenceLabel("dataNarrative", preferences.dataNarrative),
      getPreferenceLabel("pageMood", preferences.pageMood),
    ].join("、");
  }

  function savePreferences() {
    localStorage.setItem(SURVEY_STORAGE_KEY, JSON.stringify(getCurrentPreferences()));
    renderPreferenceSummary();
    if (typeof saveSettings === "function") {
      try {
        saveSettings();
      } catch {}
    }
    if (typeof renderWorkflowPlan === "function") renderWorkflowPlan();
    if (typeof renderWorkflowDetail === "function") renderWorkflowDetail();
    if (typeof renderManualLayoutPreview === "function") renderManualLayoutPreview();
    if (typeof updatePayloadPreview === "function") updatePayloadPreview();
  }

  function bindPreferenceSurvey() {
    const nodes = getPreferenceNodes();
    Object.keys(DEFAULT_PREFERENCES).forEach((key) => {
      nodes[key]?.addEventListener("change", savePreferences);
    });
  }

  function getLayoutVarietyRule(page, preferences) {
    const previousPage = page && Array.isArray(state?.workflowPages)
      ? state.workflowPages.find((item) => Number(item.pageNumber) === Number(page.pageNumber) - 1)
      : null;
    const previousSummary = previousPage
      ? String(previousPage.layoutSummary || previousPage.visualFocus || previousPage.pageType || "").trim()
      : "";
    const guides = {
      uniform: "延续统一网格、统一标题位置和稳定边距，不要频繁切换骨架。",
      balanced: "整体保持同一品牌秩序，但允许在左右分栏、上下分带、模块卡片和大数字焦点之间自然切换。",
      diverse: "主动避免与邻页同构图，优先在左右分屏、上下分区、中心聚焦、卡片矩阵、时间线和大数字重点之间切换。",
    };
    const lines = [
      `版式变化偏好：${getPreferenceLabel("layoutVariety", preferences.layoutVariety)}。${guides[preferences.layoutVariety]}`,
    ];
    if (previousSummary) {
      lines.push(`上一页已整理出的版式摘要：${previousSummary}。当前页请避免和上一页使用完全相同的主体位置、栏数和信息组织方式。`);
    }
    return lines.join("\n");
  }

  function getDetailLevelRule(preferences) {
    const guides = {
      minimal: "视觉细节偏好：偏简约。减少多余容器、描边和图形层，优先大留白、干净背景和少而准的图形装饰。",
      polished: "视觉细节偏好：精致平衡。允许适量容器、线性图形、局部材质和轻层次，但整体要克制，不能显得繁琐。",
      rich: "视觉细节偏好：偏精细。允许更丰富的容器、边框、材质层次和信息图图形，但仍然要先保证标题与正文可读。",
    };
    return guides[preferences.detailLevel];
  }

  function getVisualDensityRule(preferences) {
    const guides = {
      airy: "留白与信息量偏好：留白更多。每页优先突出 1 到 2 个重点，宁可删减辅助装饰，也不要把内容压满。",
      balanced: "留白与信息量偏好：均衡。保留清晰留白，同时用卡片、分栏和图示提升信息表达效率。",
      dense: "留白与信息量偏好：信息更满。允许更饱满的信息模块，但必须通过分组、栅格和层级控制避免小字墙。",
    };
    return guides[preferences.visualDensity];
  }

  function getCompositionFocusRule(preferences) {
    const guides = {
      imageLead: "图文主次偏好：视觉主导。优先建立明确主视觉或大图容器，让标题与正文围绕视觉主轴排布，不要平均分摊所有元素。",
      balanced: "图文主次偏好：图文平衡。让文字结构和主视觉共同承担信息表达，不偏向极端大图，也不做纯文字墙。",
      textLead: "图文主次偏好：内容主导。更强调标题层级、条目分组、信息卡片和清晰阅读路径，图形只做辅助，不要让主视觉抢占过多面积。",
    };
    return guides[preferences.compositionFocus];
  }

  function getDataNarrativeRule(page, preferences) {
    const pageType = String(page?.pageType || "").trim().toLowerCase();
    const prefix = pageType === "data" ? "当前页是数据页。" : "若当前页出现数字、对比或图表信息。";
    const guides = {
      clean: "数据页表达偏好：清晰克制。优先稳定图表区、清楚数字层级和干净容器，少做戏剧化信息图。",
      balanced: "数据页表达偏好：适度信息图。允许把数字转译成数据卡、信息图结构和辅助示意，但仍要保持读数与结论清晰。",
      expressive: "数据页表达偏好：视觉化强调。可以强化大数字、趋势图形、图示结构和更鲜明的对比，但必须避免牺牲数字可读性。",
    };
    return `${prefix}${guides[preferences.dataNarrative]}`;
  }

  function getPageMoodRule(preferences) {
    const guides = {
      steady: "整体气质偏好：稳重统一。减少过度戏剧化角度和夸张光效，优先成熟、专业、稳定的页面节奏。",
      modern: "整体气质偏好：现代清爽。保持轻盈、利落、清晰，避免背景过重或元素堆积。",
      dramatic: "整体气质偏好：更有冲击。可以适度放大主视觉、对比和视觉张力，但不要牺牲内容可读性。",
    };
    return guides[preferences.pageMood];
  }

  function buildPreferencePromptBlock(page) {
    const preferences = getCurrentPreferences();
    return [
      PREFERENCE_MARKER,
      getLayoutVarietyRule(page, preferences),
      getDetailLevelRule(preferences),
      getVisualDensityRule(preferences),
      getCompositionFocusRule(preferences),
      getDataNarrativeRule(page, preferences),
      getPageMoodRule(preferences),
    ].join("\n");
  }

  function appendPreferenceBlock(text, page) {
    const base = String(text || "").trim();
    const block = buildPreferencePromptBlock(page);
    if (!block) return base;
    if (base.includes(PREFERENCE_MARKER)) return base;
    return [base, block].filter(Boolean).join("\n\n");
  }

  function appendPreferenceTextToPayload(payload, page) {
    if (!payload?.input?.messages?.length) return payload;
    return {
      ...payload,
      input: {
        ...payload.input,
        messages: payload.input.messages.map((message, messageIndex, messages) => {
          if (messageIndex !== messages.length - 1) return message;
          const content = Array.isArray(message.content) ? message.content : [];
          return {
            ...message,
            content: content.map((item, itemIndex) => {
              if (itemIndex !== content.length - 1 || typeof item?.text !== "string") return item;
              return {
                ...item,
                text: appendPreferenceBlock(item.text, page),
              };
            }),
          };
        }),
      },
    };
  }

  function bindPreferencePromptWrappers() {
    if (typeof buildSettingsSnapshot === "function" && !buildSettingsSnapshot.__nanoPreferenceWrapped) {
      const originalBuildSettingsSnapshot = buildSettingsSnapshot;
      buildSettingsSnapshot = function buildSettingsSnapshotWithPreferences(...args) {
        return {
          ...originalBuildSettingsSnapshot.apply(this, args),
          pptPreferences: getCurrentPreferences(),
        };
      };
      buildSettingsSnapshot.__nanoPreferenceWrapped = true;
    }

    if (typeof applyLibrarySettings === "function" && !applyLibrarySettings.__nanoPreferenceWrapped) {
      const originalApplyLibrarySettings = applyLibrarySettings;
      applyLibrarySettings = function applyLibrarySettingsWithPreferences(settings = {}, ...args) {
        const result = originalApplyLibrarySettings.call(this, settings, ...args);
        if (settings?.pptPreferences) {
          applyPreferencesToUi(settings.pptPreferences);
        }
        return result;
      };
      applyLibrarySettings.__nanoPreferenceWrapped = true;
    }

    if (typeof buildWorkflowAssistantPayload === "function" && !buildWorkflowAssistantPayload.__nanoPreferenceWrapped) {
      const originalBuildWorkflowAssistantPayload = buildWorkflowAssistantPayload;
      buildWorkflowAssistantPayload = function buildWorkflowAssistantPayloadWithPreferences(...args) {
        return appendPreferenceTextToPayload(originalBuildWorkflowAssistantPayload.apply(this, args), null);
      };
      buildWorkflowAssistantPayload.__nanoPreferenceWrapped = true;
    }

    if (typeof buildWorkflowPageDesignPayload === "function" && !buildWorkflowPageDesignPayload.__nanoPreferenceWrapped) {
      const originalBuildWorkflowPageDesignPayload = buildWorkflowPageDesignPayload;
      buildWorkflowPageDesignPayload = function buildWorkflowPageDesignPayloadWithPreferences(page, ...args) {
        return appendPreferenceTextToPayload(originalBuildWorkflowPageDesignPayload.call(this, page, ...args), page);
      };
      buildWorkflowPageDesignPayload.__nanoPreferenceWrapped = true;
    }

    if (typeof buildWorkflowPagePrompt === "function" && !buildWorkflowPagePrompt.__nanoPreferenceWrapped) {
      const originalBuildWorkflowPagePrompt = buildWorkflowPagePrompt;
      buildWorkflowPagePrompt = function buildWorkflowPagePromptWithPreferences(page, ...args) {
        return appendPreferenceBlock(originalBuildWorkflowPagePrompt.call(this, page, ...args), page);
      };
      buildWorkflowPagePrompt.__nanoPreferenceWrapped = true;
    }

    if (typeof buildEffectiveWorkflowPagePrompt === "function" && !buildEffectiveWorkflowPagePrompt.__nanoPreferenceWrapped) {
      const originalBuildEffectiveWorkflowPagePrompt = buildEffectiveWorkflowPagePrompt;
      buildEffectiveWorkflowPagePrompt = function buildEffectiveWorkflowPagePromptWithPreferences(page, ...args) {
        return appendPreferenceBlock(originalBuildEffectiveWorkflowPagePrompt.call(this, page, ...args), page);
      };
      buildEffectiveWorkflowPagePrompt.__nanoPreferenceWrapped = true;
    }

    if (typeof buildPromptText === "function" && !buildPromptText.__nanoPreferenceWrapped) {
      const originalBuildPromptText = buildPromptText;
      buildPromptText = function buildPromptTextWithPreferences(...args) {
        return appendPreferenceBlock(originalBuildPromptText.apply(this, args), null);
      };
      buildPromptText.__nanoPreferenceWrapped = true;
    }
  }

  function getWorkflowProgressNodes() {
    return {
      panel: document.getElementById("workflowProgressPanel"),
      text: document.getElementById("workflowProgressText"),
      stats: document.getElementById("workflowProgressStats"),
      bar: document.getElementById("workflowProgress"),
      visual: document.getElementById("workflowProgressVisual"),
    };
  }

  function getWorkflowPageState(page) {
    if (page?.status === "success") return "generated";
    if (page?.status === "running" || page?.layoutStatus === "running") return "running";
    if (page?.status === "error" || page?.layoutStatus === "error") return "error";
    if (page?.layoutStatus === "ready") return "ready";
    return "idle";
  }

  function getWorkflowPageStateLabel(pageState) {
    return {
      idle: "待整理",
      running: "整理中",
      ready: "可生成",
      error: "待重试",
      generated: "已出图",
    }[pageState] || pageState;
  }

  function renderWorkflowProgressPanel() {
    const nodes = getWorkflowProgressNodes();
    if (!nodes.panel || !nodes.text || !nodes.stats || !nodes.bar || !nodes.visual) return;

    const pages = Array.isArray(state?.workflowPages) ? state.workflowPages : [];
    const total = pages.length;
    const ready = pages.filter((page) => page.layoutStatus === "ready").length;
    const layoutRunning = pages.filter((page) => page.layoutStatus === "running").length;
    const layoutError = pages.filter((page) => page.layoutStatus === "error").length;
    const generated = pages.filter((page) => page.status === "success").length;
    const generating = pages.filter((page) => page.status === "running").length;
    const shouldShow = Boolean(state?.workflowPlanAbortController) || total > 0;

    nodes.panel.classList.toggle("hidden", !shouldShow);
    if (!shouldShow) return;

    const readyToGenerate = pages.filter((page) => getWorkflowPageState(page) === "ready").length;
    nodes.text.textContent = workflowProgressState.lastMessage
      || (total ? `已准备好 ${ready}/${total} 页版式` : "正在开始拆分页");
    nodes.stats.textContent = total
      ? `已排版 ${ready}/${total} 页 · 可立即生成 ${readyToGenerate} 页 · 已出图 ${generated} 页${layoutError ? ` · 待重试 ${layoutError} 页` : ""}${layoutRunning ? ` · 仍在整理 ${layoutRunning} 页` : ""}${generating ? ` · 正在出图 ${generating} 页` : ""}`
      : "开始后这里会展示拆页、逐页版式整理和当前可立即生成的页数。";

    const determinate = total > 0;
    nodes.bar.classList.remove("hidden");
    nodes.bar.dataset.determinate = determinate ? "true" : "false";
    if (determinate) {
      const percent = Math.round(((ready + layoutError) / total) * 100);
      nodes.bar.style.setProperty("--progress-value", `${percent}%`);
    } else {
      nodes.bar.style.removeProperty("--progress-value");
    }

    nodes.visual.innerHTML = "";
    pages.forEach((page) => {
      const pageState = getWorkflowPageState(page);
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "workflow-progress-chip";
      chip.dataset.state = pageState;
      chip.innerHTML = `<strong>第 ${page.pageNumber} 页</strong><span>${getWorkflowPageStateLabel(pageState)}</span>`;
      chip.addEventListener("click", () => {
        if (typeof openWorkflowDetail === "function") openWorkflowDetail(page.id);
      });
      nodes.visual.appendChild(chip);
    });
  }

  function decorateWorkflowCards() {
    const cards = Array.from(document.querySelectorAll("#workflowPlanCards .workflow-card"));
    if (!cards.length || !Array.isArray(state?.workflowPages)) return;

    cards.forEach((card, index) => {
      const pageId = card.dataset.pageId || "";
      const page = state.workflowPages.find((item) => String(item.id || "") === pageId) || state.workflowPages[index];
      if (!page) return;

      const pageMeta = card.querySelector(".workflow-page-meta");
      let chip = card.querySelector(".workflow-layout-chip");
      if (!chip && pageMeta) {
        chip = document.createElement("span");
        chip.className = "workflow-layout-chip";
        pageMeta.appendChild(chip);
      }

      const pageState = getWorkflowPageState(page);
      if (chip) {
        chip.dataset.state = pageState;
        chip.textContent = pageState === "ready"
          ? "已排版，可生成"
          : pageState === "running"
            ? "排版处理中"
            : pageState === "error"
              ? "排版待重试"
              : pageState === "generated"
                ? "已生成结果"
                : "待整理";
      }

      const runBtn = card.querySelector(".run-workflow-page");
      if (runBtn) {
        runBtn.textContent = page.status === "running"
          ? "生成中..."
          : pageState === "ready" || pageState === "generated"
            ? "立即生成"
            : pageState === "running"
              ? "排版中..."
              : pageState === "error"
                ? "重试本页"
                : "生成本页";
        runBtn.disabled = page.status === "running" || state?.workflowRunning || page.layoutStatus === "running";
      }
    });
  }

  function bindWorkflowProgressWrappers() {
    if (typeof setStatus === "function" && !setStatus.__nanoExperienceWrapped) {
      const originalSetStatus = setStatus;
      setStatus = function setStatusWithProgress(message, tone = "idle", ...args) {
        workflowProgressState.lastMessage = String(message || "").trim();
        workflowProgressState.lastTone = tone;
        const result = originalSetStatus.call(this, message, tone, ...args);
        renderWorkflowProgressPanel();
        return result;
      };
      setStatus.__nanoExperienceWrapped = true;
    }

    if (typeof renderWorkflowPlan === "function" && !renderWorkflowPlan.__nanoExperienceWrapped) {
      const originalRenderWorkflowPlan = renderWorkflowPlan;
      renderWorkflowPlan = function renderWorkflowPlanWithProgress(...args) {
        const result = originalRenderWorkflowPlan.apply(this, args);
        decorateWorkflowCards();
        renderWorkflowProgressPanel();
        return result;
      };
      renderWorkflowPlan.__nanoExperienceWrapped = true;
    }

    if (typeof ensureWorkflowPageDesign === "function" && !ensureWorkflowPageDesign.__nanoExperienceWrapped) {
      const originalEnsureWorkflowPageDesign = ensureWorkflowPageDesign;
      ensureWorkflowPageDesign = async function ensureWorkflowPageDesignWithProgress(page, ...args) {
        const previousState = getWorkflowPageState(page);
        try {
          const result = await originalEnsureWorkflowPageDesign.call(this, page, ...args);
          if (page?.layoutStatus === "ready" && previousState !== "ready") {
            workflowProgressState.lastMessage = `第 ${page.pageNumber} 页已排版完成，现在可以直接生成这一页。`;
            renderWorkflowProgressPanel();
          }
          return result;
        } catch (error) {
          if (page?.layoutStatus === "error") {
            workflowProgressState.lastMessage = `第 ${page.pageNumber} 页排版失败，稍后可以单独重试这一页。`;
            renderWorkflowProgressPanel();
          }
          throw error;
        }
      };
      ensureWorkflowPageDesign.__nanoExperienceWrapped = true;
    }
  }

  function isReviseMode() {
    return state?.activeTabs?.main === "revise";
  }

  function hasImages() {
    return Array.isArray(state?.images) && state.images.length > 0;
  }

  function getBboxQuickButton() {
    return document.getElementById("bboxQuickToggleBtn");
  }

  function markReviseBboxIntent(checked) {
    if (!isReviseMode()) return;
    reviseBboxState.manualOptOut = !checked;
    if (checked) {
      reviseBboxState.announcedReady = true;
    }
  }

  function bindBboxPreferenceTracking() {
    [el?.bboxEnabled, el?.bboxEnabledImagePanel].filter(Boolean).forEach((node) => {
      if (node.dataset.nanoExperienceTracked === "true") return;
      node.dataset.nanoExperienceTracked = "true";
      node.addEventListener("change", () => {
        markReviseBboxIntent(Boolean(node.checked));
        syncBboxQuickToggleUi();
      });
    });
  }

  function shouldAutoEnableReviseBbox() {
    return isReviseMode() && hasImages() && !el?.bboxEnabled?.checked && !reviseBboxState.manualOptOut;
  }

  function decorateReviseCanvasHints() {
    if (!isReviseMode()) return;
    const button = getBboxQuickButton();
    const target = typeof getEditTargetImage === "function" ? getEditTargetImage() : null;
    document.querySelectorAll("#imageList .image-card").forEach((card) => {
      const info = card.querySelector(".box-info");
      if (!info) return;
      if (!card.querySelector(".target-badge")) return;
      if (!target) {
        info.textContent = "先选一张底稿，再决定是整图修改还是拖框局部修改。";
        return;
      }
      if (!el?.bboxEnabled?.checked) {
        info.textContent = button?.disabled
          ? "先上传底图后再开启框选。"
          : "当前底稿已就绪。点击“开启框选”后，直接在图上拖框即可局部修改。";
      }
    });
  }

  function syncBboxQuickToggleUi() {
    const button = getBboxQuickButton();
    if (!button || !el?.bboxEnabled || !el?.bboxEnabledImagePanel) return;

    const active = Boolean(el.bboxEnabled.checked);
    button.hidden = !isReviseMode();
    button.disabled = !hasImages();
    button.dataset.active = active ? "true" : "false";
    button.textContent = !hasImages() ? "先上传底图" : (active ? "关闭框选" : "开启框选");
    button.title = active
      ? "关闭后这次修改会把底图整张作为编辑输入，不再只限定框选区域。"
      : "开启后可直接在左侧底图上拖拽绘制局部修改区域。";
    decorateReviseCanvasHints();
  }

  function enableReviseBbox(options = {}) {
    const { announce = false, rerender = false } = options;
    if (!el?.bboxEnabled || !el?.bboxEnabledImagePanel || reviseBboxState.syncing) return false;
    reviseBboxState.syncing = true;
    try {
      syncBboxToggles(true);
      reviseBboxState.manualOptOut = false;
      if (typeof saveSettings === "function") saveSettings();
      if (rerender && typeof renderImages === "function") {
        renderImages();
      } else if (typeof renderEditTargetControls === "function") {
        renderEditTargetControls();
      }
      if (typeof updatePayloadPreview === "function") updatePayloadPreview();
      if (announce && !reviseBboxState.announcedReady && typeof setStatus === "function") {
        reviseBboxState.announcedReady = true;
        setStatus("外挂改图已进入框选模式，直接在底图上拖框即可局部修改。", "success");
      }
      return true;
    } finally {
      reviseBboxState.syncing = false;
      syncBboxQuickToggleUi();
    }
  }

  function bindBboxQuickToggle() {
    const button = getBboxQuickButton();
    if (!button || button.dataset.bound === "true") return;
    button.dataset.bound = "true";
    button.addEventListener("click", () => {
      if (!el?.bboxEnabled || !el?.bboxEnabledImagePanel || !hasImages()) return;
      const next = !el.bboxEnabled.checked;
      syncBboxToggles(next);
      markReviseBboxIntent(next);
      if (typeof saveSettings === "function") saveSettings();
      if (typeof renderEditTargetControls === "function") renderEditTargetControls();
      if (typeof renderImages === "function") renderImages();
      if (typeof updatePayloadPreview === "function") updatePayloadPreview();
      syncBboxQuickToggleUi();
      if (typeof setStatus === "function") {
        setStatus(
          next
            ? "已开启框选改图，现在可以直接在底图上拖拽框选区域。"
            : "已关闭框选改图，当前底图会整张作为编辑输入。",
          "success",
        );
      }
    });
  }

  function bindReviseBoardAssist() {
    const imageList = document.getElementById("imageList");
    if (imageList && imageList.dataset.nanoBoardAssist !== "true") {
      imageList.dataset.nanoBoardAssist = "true";
      imageList.addEventListener("click", (event) => {
        if (!event.target.closest(".set-target")) return;
        queueMicrotask(() => {
          if (shouldAutoEnableReviseBbox()) {
            enableReviseBbox({ announce: true, rerender: true });
          } else {
            syncBboxQuickToggleUi();
          }
        });
      });
    }

    if (el?.editTargetSelect && el.editTargetSelect.dataset.nanoBoardAssist !== "true") {
      el.editTargetSelect.dataset.nanoBoardAssist = "true";
      el.editTargetSelect.addEventListener("change", () => {
        queueMicrotask(() => {
          if (shouldAutoEnableReviseBbox()) {
            enableReviseBbox({ announce: true, rerender: true });
          } else {
            syncBboxQuickToggleUi();
          }
        });
      });
    }
  }

  function bindBboxRenderWrappers() {
    if (typeof renderImages === "function" && !renderImages.__nanoExperienceWrapped) {
      const originalRenderImages = renderImages;
      renderImages = function renderImagesWithQuickToggle(...args) {
        const result = originalRenderImages.apply(this, args);
        if (!reviseBboxState.syncing && shouldAutoEnableReviseBbox()) {
          enableReviseBbox({ announce: true, rerender: true });
        } else {
          syncBboxQuickToggleUi();
        }
        return result;
      };
      renderImages.__nanoExperienceWrapped = true;
    }

    if (typeof renderEditTargetControls === "function" && !renderEditTargetControls.__nanoExperienceWrapped) {
      const originalRenderEditTargetControls = renderEditTargetControls;
      renderEditTargetControls = function renderEditTargetControlsWithQuickToggle(...args) {
        const result = originalRenderEditTargetControls.apply(this, args);
        syncBboxQuickToggleUi();
        return result;
      };
      renderEditTargetControls.__nanoExperienceWrapped = true;
    }

    if (typeof applyTabState === "function" && !applyTabState.__nanoExperienceWrapped) {
      const originalApplyTabState = applyTabState;
      applyTabState = function applyTabStateWithQuickToggle(...args) {
        const result = originalApplyTabState.apply(this, args);
        if (!reviseBboxState.syncing && shouldAutoEnableReviseBbox()) {
          enableReviseBbox({ announce: false, rerender: true });
        } else {
          syncBboxQuickToggleUi();
        }
        renderWorkflowProgressPanel();
        return result;
      };
      applyTabState.__nanoExperienceWrapped = true;
    }
  }

  function init() {
    applyPreferencesToUi(readStoredPreferences());
    bindPreferenceSurvey();
    bindPreferencePromptWrappers();
    bindWorkflowProgressWrappers();
    bindBboxPreferenceTracking();
    bindBboxQuickToggle();
    bindReviseBoardAssist();
    bindBboxRenderWrappers();
    syncBboxQuickToggleUi();
    renderWorkflowProgressPanel();
    if (typeof renderWorkflowPlan === "function") renderWorkflowPlan();
  }

  init();
})();
