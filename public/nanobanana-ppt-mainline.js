(function bootstrapNanoBananaPptMainline() {
  const MAIN_MODEL = "gemini-3-pro-image-preview";
  const EDIT_MODEL = "wan2.7-image-pro";
  const APP_TITLE = "Nano Banana PPT Studio";
  const MAINLINE_NOTICE = "当前主线：Qwen 继续负责主题生成、拆分页和逐页版式整理；Nano Banana Pro 负责最终 PPT 生图。";
  const EDIT_ADDON_NOTICE = "外挂改图模块：这里继续保留 Wan 改图能力，包括底图选择、bbox 定点改图和历史回退，不影响上面的 Nano Banana PPT 主线。";

  function getNode(target) {
    if (!target) return null;
    return typeof target === "string" ? document.querySelector(target) : target;
  }

  function setText(target, text) {
    const node = getNode(target);
    if (node) node.textContent = text;
  }

  function hideElement(target) {
    const node = getNode(target);
    if (!node) return;
    node.hidden = true;
    node.setAttribute("aria-hidden", "true");
    node.style.display = "none";
  }

  function showElement(target) {
    const node = getNode(target);
    if (!node) return;
    node.hidden = false;
    node.removeAttribute("aria-hidden");
    node.style.display = "";
  }

  function getFieldByControlId(id) {
    return document.getElementById(id)?.closest(".field") || null;
  }

  function relabelField(id, labelText) {
    const field = getFieldByControlId(id);
    const label = field?.querySelector("span");
    if (label) label.textContent = labelText;
    return field;
  }

  function ensureHintBox(id, text, anchor) {
    const existing = document.getElementById(id);
    if (existing) {
      existing.textContent = text;
      return existing;
    }
    const target = getNode(anchor);
    if (!target || !target.parentNode) return null;

    const node = document.createElement("div");
    node.id = id;
    node.className = "hint-box api-key-status";
    node.textContent = text;
    target.parentNode.insertBefore(node, target.nextSibling);
    return node;
  }

  function updateBranding() {
    document.title = APP_TITLE;
    setText(".eyebrow", "Google Gemini / Nano Banana Pro / Wan 改图外挂 / Qwen 3.6 / PPT Harness");
    setText(".hero-copy h1", APP_TITLE);
    setText(
      ".hero-copy .hero-text",
      "保留原来的 Qwen 拆分页和逐页版式流程，把最终 PPT 生图固定切到 Nano Banana Pro；同时把原来的 Wan 改图能力作为外挂模块继续保留，方便做局部回修。",
    );

    const badgeTexts = ["Nano Banana PPT", "Wan 改图外挂", "PPT Harness"];
    document.querySelectorAll(".hero-badge span").forEach((node, index) => {
      if (badgeTexts[index]) node.textContent = badgeTexts[index];
    });

    setText(".settings-panel h2", "主线设置");
    setText(
      ".settings-panel .panel-head p",
      "Qwen 继续承担主题生成、拆分页和 Harness 辅助链路；最终 PPT 生图固定走 Nano Banana Pro。改图模块会在 revise 流程里单独走 Wan。",
    );

    relabelField("apiKey", "DashScope / Qwen / Wan API Key");
    relabelField("googleApiKey", "Google API Key（Nano Banana）");

    const apiKeyInput = document.getElementById("apiKey");
    if (apiKeyInput) apiKeyInput.placeholder = "sk-...（用于 Qwen / Harness / Wan 改图）";

    const googleApiKeyInput = document.getElementById("googleApiKey");
    if (googleApiKeyInput) googleApiKeyInput.placeholder = "AIza...（用于 Nano Banana PPT 生图）";

    const flowCopy = document.querySelector(".flow-nav-panel .panel-head p");
    if (flowCopy) {
      flowCopy.textContent = "保留原来的拆分页和手动排版主流程，再把改图模块作为外挂保留下来。";
    }

    setText('[data-main-tab="smart"]', "1. 自动拆页与逐页生图");
    setText('[data-main-tab="manual"]', "2. 手动排版与单页精修");
    setText('[data-main-tab="revise"]', "3. 外挂改图（Wan）");

    const harnessToggle = document.querySelector('label[for="pptHarnessEnabled"] span');
    if (harnessToggle) harnessToggle.textContent = "PPT Harness（已迁移）";

    const statusBar = document.querySelector("#statusBar");
    if (statusBar) {
      statusBar.textContent = "主线已就绪：Qwen 负责拆页与版式，Nano Banana Pro 负责最终生图，Wan 外挂改图模块也已恢复。";
    }

    const reviseTitle = document.querySelector('.flow-panel[data-main-panel="revise"] .panel-head h2');
    if (reviseTitle) reviseTitle.textContent = "外挂改图模块：沿用 Wan 改图链路";

    const reviseDesc = document.querySelector('.flow-panel[data-main-panel="revise"] .panel-head p');
    if (reviseDesc) {
      reviseDesc.textContent = "这条支线继续沿用原来的 Wan 改图能力，适合对已生成页面或底图做局部回修。";
    }
  }

  function wrapSyncImageProviderUi() {
    if (typeof syncImageProviderUi !== "function" || syncImageProviderUi.__nanoMainlineWrapped) return;
    const original = syncImageProviderUi;
    const wrapped = function wrappedSyncImageProviderUi(...args) {
      const result = original.apply(this, args);
      const testBtn = document.getElementById("testApiKeyBtn");
      if (testBtn && !testBtn.disabled) {
        testBtn.dataset.idleLabel = "测试主线与外挂配置";
        testBtn.textContent = "测试主线与外挂配置";
        testBtn.title = "会测试 Nano Banana PPT 生图链路，以及 DashScope 侧的 Qwen / Wan 可用性。";
      }
      return result;
    };
    wrapped.__nanoMainlineWrapped = true;
    syncImageProviderUi = wrapped;
  }

  function syncModeSpecificControls() {
    const currentMain = state?.activeTabs?.main || "smart";
    const reviseMode = currentMain === "revise";

    relabelField("model", reviseMode ? "改图模型（当前固定）" : "PPT 生图模型（当前固定）");

    const model = document.getElementById("model");
    if (model) {
      model.innerHTML = reviseMode
        ? '<option value="wan2.7-image-pro">wan2.7-image-pro · Wan 改图外挂</option>'
        : '<option value="gemini-3-pro-image-preview">gemini-3-pro-image-preview · Nano Banana Pro</option>';
      model.value = reviseMode ? EDIT_MODEL : MAIN_MODEL;
      model.disabled = true;
      model.title = reviseMode
        ? "当前外挂改图模块固定走 wan2.7-image-pro，并保留 bbox 定点改图。"
        : "PPT 主线当前固定使用 Nano Banana Pro；切到外挂改图时会自动改走 wan2.7-image-pro。";
    }

    const requestModeField = getFieldByControlId("requestMode");
    const requestMode = document.getElementById("requestMode");
    if (requestMode) {
      if (reviseMode) {
        showElement(requestModeField);
        requestMode.disabled = false;
      } else {
        requestMode.value = "sync";
        requestMode.disabled = true;
        hideElement(requestModeField);
      }
    }

    const regionField = getFieldByControlId("region");
    const region = document.getElementById("region");
    if (region) {
      if (reviseMode) {
        showElement(regionField);
        region.disabled = false;
      } else {
        region.value = "beijing";
        region.disabled = true;
        hideElement(regionField);
      }
    }

    wrapSyncImageProviderUi();
    if (typeof syncImageProviderUi === "function") {
      syncImageProviderUi();
    }
  }

  function ensureMainlineNotice() {
    ensureHintBox("nanoBananaMainlineNotice", MAINLINE_NOTICE, ".settings-panel .settings-grid");
  }

  function ensureEditAddonNotice() {
    const panelHead = document.querySelector('.flow-panel[data-main-panel="revise"] .panel-head');
    if (!panelHead) return;
    let notice = document.getElementById("wanEditAddonNotice");
    if (!notice) {
      notice = document.createElement("div");
      notice.id = "wanEditAddonNotice";
      notice.className = "hint-box";
      panelHead.parentNode.insertBefore(notice, panelHead.nextSibling);
    }
    notice.textContent = EDIT_ADDON_NOTICE;
  }

  function tuneRevisePanel() {
    ensureEditAddonNotice();
    setText("#promptLabel", "改图需求 / Wan 指令");

    const prompt = document.getElementById("prompt");
    if (prompt) {
      prompt.placeholder = "例如：保留版式，只提升质感；或仅修改右上角主体，并继续使用 bbox 定点改图。";
    }

    const studioTitle = document.getElementById("studioTitle");
    if (studioTitle) studioTitle.textContent = "外挂改图模块";

    const studioHint = document.getElementById("studioHint");
    if (studioHint) {
      studioHint.textContent = "这里固定走 Wan 改图链路，支持 bbox 定点改图，不会影响 Nano Banana 的 PPT 生图主线。";
    }

    const sendBtn = document.getElementById("sendBtn");
    if (sendBtn) sendBtn.textContent = "提交到 Wan 改图";
  }

  function tuneMainlinePanel() {
    setText("#promptLabel", "本次创作需求 / 主提示词");

    const prompt = document.getElementById("prompt");
    if (prompt) {
      prompt.placeholder = "例如：做一页关于城市低碳交通的 PPT 主视觉，保留标题区，装饰仅限图形，不要额外文字。";
    }

    const studioTitle = document.getElementById("studioTitle");
    if (studioTitle) studioTitle.textContent = "Nano Banana PPT 主线";

    const studioHint = document.getElementById("studioHint");
    if (studioHint) {
      studioHint.textContent = "这里固定走 Nano Banana Pro 的 PPT 生图主线；拆分页与 Harness 仍由 Qwen 负责。";
    }

    const sendBtn = document.getElementById("sendBtn");
    if (sendBtn) sendBtn.textContent = "提交到 Nano Banana";
  }

  function wrapApplyTabState() {
    if (typeof applyTabState !== "function" || applyTabState.__nanoMainlineWrapped) return;
    const original = applyTabState;
    const wrapped = function wrappedApplyTabState(...args) {
      const result = original.apply(this, args);
      syncModeSpecificControls();
      if (state?.activeTabs?.main === "revise") {
        tuneRevisePanel();
      } else {
        tuneMainlinePanel();
      }
      return result;
    };
    wrapped.__nanoMainlineWrapped = true;
    applyTabState = wrapped;
  }

  function coerceMainlineState() {
    const reviseMode = (state?.activeTabs?.main || "smart") === "revise";
    let dirty = false;

    const model = document.getElementById("model");
    const expectedModel = reviseMode ? EDIT_MODEL : MAIN_MODEL;
    if (model && model.value !== expectedModel) {
      model.value = expectedModel;
      dirty = true;
    }

    const requestMode = document.getElementById("requestMode");
    if (!reviseMode && requestMode && requestMode.value !== "sync") {
      requestMode.value = "sync";
      dirty = true;
    }

    if (dirty && typeof saveSettings === "function") {
      try {
        saveSettings();
      } catch (error) {
        console.warn("Failed to persist Nano Banana mainline defaults:", error);
      }
    }
  }

  function initMainline() {
    if (typeof ensureGoogleApiKeyField === "function") {
      ensureGoogleApiKeyField();
    }
    updateBranding();
    syncModeSpecificControls();
    wrapApplyTabState();
    ensureMainlineNotice();
    ensureEditAddonNotice();
    coerceMainlineState();
    if (typeof applyTabState === "function") {
      applyTabState();
    } else {
      syncModeSpecificControls();
      tuneMainlinePanel();
    }
    if (typeof renderApiKeyStatus === "function") {
      renderApiKeyStatus();
    }
  }

  initMainline();
})();
