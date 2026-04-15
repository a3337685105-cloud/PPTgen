(function applyNanoBananaOverride() {
  const PPT_MODEL = "gemini-3-pro-image-preview";
  const EDIT_MODEL = "wan2.7-image-pro";

  function isAddonEditMode() {
    return state?.activeTabs?.main === "revise";
  }

  function getCurrentGenerationModel(fallbackModel = el?.model?.value) {
    return isAddonEditMode() ? EDIT_MODEL : (fallbackModel || PPT_MODEL);
  }

  function getCurrentGenerationLabel(model = getCurrentGenerationModel()) {
    return model === EDIT_MODEL ? "Wan 外挂改图" : "Nano Banana PPT 生图";
  }

  function ensurePayloadHasContent(payload) {
    const messages = Array.isArray(payload?.input?.messages) ? payload.input.messages : [];
    return messages.some((message) => {
      const items = Array.isArray(message?.content) ? message.content : [];
      return items.some((item) => {
        if (typeof item?.text === "string" && item.text.trim()) return true;
        if (typeof item?.image === "string" && item.image.trim()) return true;
        return false;
      });
    });
  }

  function buildWanAddonParameters({
    orderedImages = [],
    enableSequential = el.enableSequential.checked,
    enableBbox = el.bboxEnabled.checked,
  } = {}) {
    const imageCount = Number(el.imageCount?.value ?? 1);
    if (enableSequential ? imageCount < 1 || imageCount > 12 : imageCount < 1 || imageCount > 4) {
      throw new Error(enableSequential ? "组图模式下 n 的范围应为 1 到 12。" : "非组图模式下 n 的范围应为 1 到 4。");
    }

    const parameters = {
      size: buildSizeValue(),
      n: imageCount,
      watermark: el.watermark.checked,
    };

    if (enableSequential) parameters.enable_sequential = true;
    if (el.seed?.value?.trim()) parameters.seed = Number(el.seed.value.trim());
    if (el.thinkingMode?.checked && !enableSequential && orderedImages.length === 0) {
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

  function normalizeGenerationPayload(payload) {
    if (!payload || typeof payload !== "object") return payload;
    const model = getCurrentGenerationModel(payload.model);

    if (model === EDIT_MODEL) {
      const orderedImages = typeof getOrderedImagesForPayload === "function" ? getOrderedImagesForPayload() : [];
      return {
        ...payload,
        model,
        parameters: buildWanAddonParameters({
          orderedImages,
          enableSequential: el.enableSequential.checked,
          enableBbox: el.bboxEnabled.checked,
        }),
      };
    }

    return {
      ...payload,
      model: PPT_MODEL,
    };
  }

  function validateGenerationPayload(payload) {
    if (!payload || typeof payload !== "object") return "当前没有可发送的请求体。";
    if (!payload.model) return "当前请求缺少模型设置。";
    const messages = Array.isArray(payload.input?.messages) ? payload.input.messages : [];
    if (!messages.length) return "当前请求缺少消息数组。";
    if (!ensurePayloadHasContent(payload)) return "当前请求里没有可用的文本或图片内容。";
    return "";
  }

  function buildApiKeyStatusSummary() {
    const dashScopeKey = getDashScopeApiKey();
    const googleKey = getGoogleApiKey();
    const dashScopeLabel = dashScopeKey ? `DashScope / Wan / Qwen ${maskApiKey(dashScopeKey)}` : "DashScope / Wan / Qwen 未填写";
    const googleLabel = googleKey ? `Nano Banana ${maskApiKey(googleKey)}` : "Nano Banana 未填写";
    return `当前链路：Qwen 负责拆页与版式，Nano Banana 负责 PPT 生图，Wan 负责外挂改图 · ${dashScopeLabel} · ${googleLabel}`;
  }

  renderApiKeyStatus = function renderApiKeyStatusOverride(message, tone = "idle") {
    if (!el.apiKeyStatus) return;
    el.apiKeyStatus.textContent = message || buildApiKeyStatusSummary();
    el.apiKeyStatus.dataset.tone = tone;
  };

  syncApiKeyFeedback = function syncApiKeyFeedbackOverride(response, data, options = {}) {
    const provider = options.provider || (isGeminiImageModel(options.model) ? "gemini" : "dashscope");
    const providerLabel = provider === "gemini" ? "Google" : "DashScope";
    const providerKey = provider === "gemini" ? getGoogleApiKey() : getDashScopeApiKey();
    const message = String(data?.message || data?.error?.message || "");

    if (response?.status === 401 || response?.status === 403 || /invalid api-?key|api key not valid|permission denied|unauthenticated/i.test(message)) {
      renderApiKeyStatus(`${providerLabel} Key 无效`, "error");
      return;
    }

    if (response?.ok && providerKey) {
      renderApiKeyStatus();
    }
  };

  if (typeof buildPayload === "function" && !buildPayload.__nanoWrapped) {
    const originalBuildPayload = buildPayload;
    const wrappedBuildPayload = function wrappedBuildPayload(...args) {
      return normalizeGenerationPayload(originalBuildPayload.apply(this, args));
    };
    wrappedBuildPayload.__nanoWrapped = true;
    buildPayload = wrappedBuildPayload;
  }

  if (typeof buildWorkflowGenerationPayload === "function" && !buildWorkflowGenerationPayload.__nanoWrapped) {
    const originalBuildWorkflowGenerationPayload = buildWorkflowGenerationPayload;
    const wrappedBuildWorkflowGenerationPayload = function wrappedBuildWorkflowGenerationPayload(...args) {
      const payload = originalBuildWorkflowGenerationPayload.apply(this, args);
      return {
        ...payload,
        model: PPT_MODEL,
      };
    };
    wrappedBuildWorkflowGenerationPayload.__nanoWrapped = true;
    buildWorkflowGenerationPayload = wrappedBuildWorkflowGenerationPayload;
  }

  testApiKey = async function testApiKeyOverride() {
    const googleKey = getGoogleApiKey();
    const dashScopeKey = getDashScopeApiKey();
    const currentMain = state?.activeTabs?.main || "smart";

    if (currentMain === "revise" && !dashScopeKey) {
      renderApiKeyStatus("外挂改图模块走 Wan，请先填写 DashScope / Wan API Key。", "error");
      setStatus("外挂改图模块走 Wan，请先填写 DashScope / Wan API Key。", "error");
      return;
    }

    if (currentMain !== "revise" && !googleKey) {
      renderApiKeyStatus("PPT 生图主线走 Nano Banana，请先填写 Google API Key。", "error");
      setStatus("PPT 生图主线走 Nano Banana，请先填写 Google API Key。", "error");
      return;
    }

    if (!googleKey && !dashScopeKey) {
      renderApiKeyStatus("请先填写至少一个可用 Key。", "error");
      setStatus("请先填写至少一个可用 Key。", "error");
      return;
    }

    saveSettings();
    syncImageProviderUi();

    setButtonLoading(el.testApiKeyBtn, true, "测试主线与外挂配置...");
    renderApiKeyStatus("正在测试 Nano Banana 与 DashScope / Wan / Qwen 链路...", "running");
    setStatus("正在测试 Nano Banana 和 Wan / Qwen 链路...", "running");

    try {
      const checks = [];

      if (googleKey) {
        checks.push(
          fetch("/api/test-image-key", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              apiKey: dashScopeKey,
              googleApiKey: googleKey,
              region: el.region.value,
              model: PPT_MODEL,
            }),
          }).then(async (response) => ({
            role: "nano",
            response,
            data: await response.json(),
          })),
        );
      }

      if (dashScopeKey) {
        checks.push(
          fetch("/api/test-image-key", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              apiKey: dashScopeKey,
              googleApiKey: googleKey,
              region: el.region.value,
              model: EDIT_MODEL,
            }),
          }).then(async (response) => ({
            role: "dashscope",
            response,
            data: await response.json(),
          })),
        );
      }

      const results = await Promise.all(checks);
      results.forEach(({ role, response, data }) => {
        syncApiKeyFeedback(response, data, { model: role === "nano" ? PPT_MODEL : EDIT_MODEL });
      });

      const failures = results.filter(({ response, data }) => !response.ok || data?.code);
      if (failures.length) {
        const message = failures
          .map(({ role, data }) => `${role === "nano" ? "Nano Banana" : "DashScope / Wan / Qwen"}：${data?.message || "测试失败"}`)
          .join("；");
        renderApiKeyStatus(`配置未通过 · ${message}`, "error");
        setStatus(message, "error");
        return;
      }

      const successParts = [];
      if (results.some((item) => item.role === "nano")) successParts.push("Nano Banana PPT 生图可用");
      if (results.some((item) => item.role === "dashscope")) successParts.push("DashScope / Wan / Qwen 链路可用");
      renderApiKeyStatus(successParts.join(" · "), "success");
      setStatus(successParts.join("，"), "success");
    } catch (error) {
      renderApiKeyStatus(`测试失败 · ${error.message || "网络失败"}`, "error");
      setStatus(error.message || "配置测试失败。", "error");
    } finally {
      setButtonLoading(el.testApiKeyBtn, false);
      syncImageProviderUi();
    }
  };

  sendRequest = async function sendRequestOverride() {
    saveSettings();
    updatePayloadPreview();

    let payload;
    try {
      payload = buildPayload();
    } catch (error) {
      setStatus(error.message, "error");
      return;
    }

    const validationError = validateGenerationPayload(payload);
    if (validationError) {
      setStatus(validationError, "error");
      return;
    }

    if (payload.model === EDIT_MODEL && !getDashScopeApiKey()) {
      setStatus("外挂改图模块走 Wan，请先填写 DashScope / Wan API Key。", "error");
      return;
    }
    if (payload.model !== EDIT_MODEL && !getGoogleApiKey()) {
      setStatus("PPT 生图主线走 Nano Banana，请先填写 Google API Key。", "error");
      return;
    }

    if (payload.model === EDIT_MODEL && !state.images.length) {
      setStatus("外挂改图模块至少需要一张底图或参考图。", "error");
      return;
    }

    if (payload.model !== EDIT_MODEL && el.bboxEnabled.checked) {
      setStatus("Nano Banana PPT 生图暂不支持当前 bbox 定点改图，请切到外挂改图模块。", "error");
      return;
    }

    if (payload.model === EDIT_MODEL && el.bboxEnabled.checked) {
      const target = getEditTargetImage();
      if (!target) {
        setStatus("你已启用定点改图，但还没有选择改图底图。", "error");
        return;
      }
      if (!target.boxes.length) {
        setStatus("你已启用定点改图，但当前底图还没有绘制任何 bbox 区域。", "error");
        return;
      }
    }

    setStatus(
      payload.model === EDIT_MODEL
        ? "正在调用 Wan 外挂改图模块..."
        : "正在调用 Nano Banana 生成最终 PPT 页面图...",
      "running",
    );

    try {
      const { response, data } = await requestGeneration(payload);
      renderResponse(data);
      if (!response.ok || data.code) {
        stopPolling();
        state.currentTaskId = null;
        setStatus(data.message || "调用失败。", "error");
        return;
      }

      if (shouldUseAsyncImageGeneration(payload.model)) {
        const taskId = data.output?.task_id;
        if (!taskId) {
          setStatus("异步任务已创建，但没有返回 task_id。", "error");
          return;
        }
        setStatus(`任务已创建，task_id：${taskId}，开始轮询。`, "running");
        startPolling(taskId);
        return;
      }

      stopPolling();
      state.currentTaskId = null;
      setStatus(payload.model === EDIT_MODEL ? "Wan 改图完成。" : "Nano Banana PPT 生图完成。", "success");
    } catch (error) {
      stopPolling();
      state.currentTaskId = null;
      setStatus(error.message || "调用失败。", "error");
    }
  };

  generateWorkflowPage = async function generateWorkflowPageOverride(pageId, options = {}) {
    const { stopOnError = false } = options;
    const page = state.workflowPages.find((item) => item.id === pageId);
    if (!page) {
      setStatus("未找到对应页面。", "error");
      return;
    }

    if (!getGoogleApiKey()) {
      setStatus("逐页生图走 Nano Banana，请先填写 Google API Key。", "error");
      return;
    }

    page.status = "running";
    page.error = "正在生成本页...";
    renderWorkflowPlan();
    setStatus(`正在调用 Nano Banana 生成第 ${page.pageNumber} 页终图...`, "running");

    try {
      if (workflowPageNeedsDesign(page)) {
        setStatus(`正在整理第 ${page.pageNumber} 页的版式和提示词...`, "running");
        await ensureWorkflowPageDesign(page, {
          force: page.layoutStatus === "error" || !hasWorkflowStructuredLayout(page),
        });
      }

      const payload = buildWorkflowGenerationPayload(page);
      const validationError = validateGenerationPayload(payload);
      if (validationError) {
        throw new Error(`第 ${page.pageNumber} 页请求不完整：${validationError}`);
      }

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
      if (shouldUseAsyncImageGeneration(payload.model)) {
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
  };

  function rebindButton(key, handler) {
    const current = el[key];
    if (!current) return;
    const clone = current.cloneNode(true);
    current.replaceWith(clone);
    el[key] = clone;
    clone.addEventListener("click", handler);
  }

  rebindButton("testApiKeyBtn", testApiKey);
  rebindButton("sendBtn", sendRequest);
  syncImageProviderUi();
  if (typeof updatePayloadPreview === "function") updatePayloadPreview();
  renderApiKeyStatus();
})();
