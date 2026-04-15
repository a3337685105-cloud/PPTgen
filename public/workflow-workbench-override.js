(function workflowWorkbenchOverride() {
  if (typeof el === "undefined" || typeof state === "undefined") return;

  function hideSmartWorkflowChrome() {
    const smartPanel = document.querySelector('[data-main-panel="smart"]');
    const sourceSubpanel = smartPanel?.querySelector('.smart-flow-grid > .subpanel');
    const metaGrid = sourceSubpanel?.querySelector('.grid.two');
    const flowInfoField = metaGrid?.querySelector('.field:not(label)');

    smartPanel?.querySelector('.panel-head > p')?.setAttribute('hidden', 'hidden');
    sourceSubpanel?.querySelector('.subpanel-head')?.setAttribute('hidden', 'hidden');

    if (el.workflowGateHint) {
      el.workflowGateHint.hidden = true;
      el.workflowGateHint.style.display = 'none';
    }
    if (el.workflowPlanSummary) {
      el.workflowPlanSummary.hidden = true;
      el.workflowPlanSummary.style.display = 'none';
    }
    if (metaGrid) metaGrid.classList.add('workflow-compact-meta');
    if (flowInfoField) {
      flowInfoField.hidden = true;
      flowInfoField.style.display = 'none';
    }
  }

  function normalizeWorkflowDetailHeadings() {
    const detailCard = document.querySelector('#workflowDetailDisplayTitleInput')?.closest('.theme-review-card');
    if (detailCard) {
      const strongs = Array.from(detailCard.querySelectorAll('strong'));
      strongs.forEach((node, index) => {
        node.hidden = index !== 0;
        node.style.display = index === 0 ? '' : 'none';
        if (index === 0) node.textContent = '最终上屏内容';
      });
    }

    const layoutCard = document.querySelector('#workflowDetailLayout')?.closest('.theme-review-card');
    if (layoutCard) {
      const strongs = Array.from(layoutCard.querySelectorAll('strong'));
      strongs.forEach((node, index) => {
        node.hidden = index !== 0;
        node.style.display = index === 0 ? '' : 'none';
        if (index === 0) node.textContent = '格式安排';
      });
    }
  }

  const prevEnsureWorkflowResearchPopover = ensureWorkflowResearchPopover;
  ensureWorkflowResearchPopover = function ensureWorkflowResearchPopoverDocked() {
    const refs = prevEnsureWorkflowResearchPopover();
    const host = document.querySelector('#workflowDetailConfirmContentBtn')?.closest('.workflow-detail-editor-actions')
      || document.querySelector('#workflowDetailDisplayTitleInput')?.closest('.theme-review-card')
      || el.workflowDetailPanel;
    if (host && refs.popover.parentElement !== host) {
      host.appendChild(refs.popover);
    }
    return refs;
  };

  function ensureWorkflowPromptPopover() {
    const actions = document.querySelector('#workflowDetailConfirmContentBtn')?.closest('.workflow-detail-editor-actions');
    if (!actions) return {};

    let toggleBtn = document.querySelector('#workflowDetailPromptPopoverBtn');
    if (!toggleBtn) {
      toggleBtn = document.createElement('button');
      toggleBtn.id = 'workflowDetailPromptPopoverBtn';
      toggleBtn.type = 'button';
      toggleBtn.className = 'ghost-btn workflow-detail-mini-btn';
      toggleBtn.textContent = '查看原始提示词';
      actions.appendChild(toggleBtn);
    }

    let popover = document.querySelector('#workflowDetailPromptPopover');
    if (!popover) {
      popover = document.createElement('section');
      popover.id = 'workflowDetailPromptPopover';
      popover.className = 'workflow-prompt-popover hidden';
      popover.hidden = true;
      popover.innerHTML = `
        <div class="workflow-prompt-popover-head">
          <strong>原始提示词</strong>
          <button type="button" class="ghost-btn workflow-detail-mini-btn" data-action="close-prompt-popover">关闭</button>
        </div>
        <pre id="workflowDetailPromptPopoverContent" class="workflow-prompt-popover-body"></pre>
      `;
      actions.appendChild(popover);
    }

    const closeBtn = popover.querySelector('[data-action="close-prompt-popover"]');
    if (closeBtn && !closeBtn.dataset.bound) {
      closeBtn.dataset.bound = 'true';
      closeBtn.addEventListener('click', () => {
        state.workflowPromptPopoverOpen = false;
        state.workflowPromptPopoverPageId = '';
        popover.hidden = true;
        popover.classList.add('hidden');
      });
    }

    return {
      actions,
      toggleBtn,
      popover,
      content: popover.querySelector('#workflowDetailPromptPopoverContent'),
    };
  }

  function getWorkflowOriginalPrompt(page) {
    const prompt = String(page?.visualPromptTemplate || page?.pagePrompt || '').trim();
    return prompt || buildWorkflowPagePrompt(page);
  }

  function renderWorkflowPromptPopover(page) {
    const refs = ensureWorkflowPromptPopover();
    if (!refs.popover || !refs.content || !page) return refs;
    refs.content.textContent = getWorkflowOriginalPrompt(page) || '这一页还没有原始提示词。';
    refs.toggleBtn.disabled = false;
    refs.popover.hidden = !state.workflowPromptPopoverOpen || state.workflowPromptPopoverPageId !== page.id;
    refs.popover.classList.toggle('hidden', refs.popover.hidden);
    return refs;
  }

  function normalizeWorkflowResultImageUrl(input) {
    const value = String(input || '').trim();
    if (!value || value === 'undefined' || value === 'null') return '';
    if (/^[a-zA-Z]:[\\/]/.test(value)) {
      const normalizedPath = value.replace(/\\/g, '/');
      const generatedImagesIndex = normalizedPath.lastIndexOf('/generated-images/');
      return generatedImagesIndex >= 0 ? normalizedPath.slice(generatedImagesIndex) : '';
    }
    return value;
  }

  function getWorkflowPageResultImageCandidates(page) {
    const collect = (input) => Array.isArray(input)
      ? input.map((item) => normalizeWorkflowResultImageUrl(item)).filter(Boolean)
      : [];
    const uniqueCandidates = [];
    const seen = new Set();
    const pushCandidate = (input) => {
      const value = normalizeWorkflowResultImageUrl(input);
      if (!value || seen.has(value)) return;
      seen.add(value);
      uniqueCandidates.push(value);
    };
    const pushSavedResultCandidates = (savedResults) => {
      if (!savedResults || typeof savedResults !== 'object') return;
      Object.values(savedResults).forEach((item) => {
        pushCandidate(item?.localUrl);
        pushCandidate(item?.savedPath);
      });
      Object.values(savedResults).forEach((item) => {
        pushCandidate(item?.url);
        pushCandidate(item?.imageUrl);
      });
      Object.keys(savedResults).forEach((item) => pushCandidate(item));
    };
    const pushPageCandidates = (targetPage) => {
      if (!targetPage || typeof targetPage !== 'object') return;
      const firstResult = collect(targetPage?.resultImages)[0] || collect(targetPage?.result_images)[0] || '';
      if (typeof getWorkflowPreferredImageUrl === 'function') {
        pushCandidate(getWorkflowPreferredImageUrl(targetPage, firstResult));
      }
      if (typeof getWorkflowSavedLocalImageUrl === 'function') {
        pushCandidate(getWorkflowSavedLocalImageUrl(targetPage, firstResult));
        pushCandidate(getWorkflowSavedLocalImageUrl(targetPage));
      }
      pushSavedResultCandidates(targetPage?.savedResults);
      pushCandidate(targetPage?.detailBackdropUrl || targetPage?.detail_backdrop_url);
      collect(targetPage?.resultImages).forEach((item) => pushCandidate(item));
      collect(targetPage?.result_images).forEach((item) => pushCandidate(item));
    };
    const getCardResultUrl = () => {
      const root = el.workflowPlanCards || document;
      const pageId = String(page?.id || '').trim();
      const pageNumber = String(page?.pageNumber || '').trim();
      let card = null;

      if (pageId) {
        try {
          const escaped = typeof CSS !== 'undefined' && typeof CSS.escape === 'function'
            ? CSS.escape(pageId)
            : pageId.replace(/["\\]/g, '\\$&');
          card = root.querySelector(`.workflow-card[data-page-id="${escaped}"]`);
        } catch {}
      }

      if (!card && pageNumber) {
        const cards = Array.from(root.querySelectorAll('.workflow-card'));
        card = cards.find((node) => {
          const explicit = String(node.getAttribute('data-page-number') || '').trim();
          if (explicit && explicit === pageNumber) return true;
          const meta = String(node.querySelector('.workflow-page-index')?.textContent || '').trim();
          return meta.includes(pageNumber);
        }) || null;
      }

      if (!card) return '';
      return normalizeWorkflowResultImageUrl(
        card.getAttribute('data-result-image-url')
        || card.querySelector('.result-images img, .result-card img')?.getAttribute('src')
        || card.querySelector('.result-images img, .result-card img')?.src
      );
    };

    pushPageCandidates(page);

    const livePage = Array.isArray(state.workflowPages)
      ? state.workflowPages.find((item) => item.id === page?.id)
      : null;
    if (livePage && livePage !== page) {
      pushPageCandidates(livePage);
    }

    if (typeof getWorkflowCardResultImageUrl === 'function' && page?.id) {
      pushCandidate(getWorkflowCardResultImageUrl(page.id));
    }
    pushCandidate(getCardResultUrl());

    return uniqueCandidates;
  }

  function getWorkflowPageResultImageUrl(page) {
    return getWorkflowPageResultImageCandidates(page)[0] || '';
  }

  function applyWorkflowPageResultBackdrop(page, stage, backdrop, preferredUrl = '') {
    const candidates = [];
    const seen = new Set();
    const pushCandidate = (input) => {
      const value = normalizeWorkflowResultImageUrl(input);
      if (!value || seen.has(value)) return;
      seen.add(value);
      candidates.push(value);
    };
    const clearBackdrop = () => {
      if (backdrop) {
        backdrop.onload = null;
        backdrop.onerror = null;
        backdrop.hidden = true;
        backdrop.classList.add('hidden');
        backdrop.removeAttribute('src');
        backdrop.style.opacity = '0';
      }
      if (stage) {
        stage.dataset.backdropUrl = '';
        stage.style.backgroundImage = '';
        stage.style.backgroundSize = '';
        stage.style.backgroundPosition = '';
        stage.style.backgroundRepeat = '';
      }
    };
    const syncPageBackdrop = (url) => {
      if (!page || !url) return;
      if (page.detailBackdropUrl !== url) {
        page.detailBackdropUrl = url;
      }
      if (!Array.isArray(page.resultImages)) page.resultImages = [];
      if (!page.resultImages.includes(url)) {
        page.resultImages = [url, ...page.resultImages].filter(Boolean);
      }
    };
    const applyCandidateAt = (index) => {
      const nextUrl = candidates[index] || '';
      if (!nextUrl) {
        clearBackdrop();
        return '';
      }

      syncPageBackdrop(nextUrl);
      if (stage) {
        stage.dataset.backdropUrl = nextUrl;
        stage.style.backgroundImage = `url("${nextUrl}")`;
        stage.style.backgroundSize = 'cover';
        stage.style.backgroundPosition = 'center center';
        stage.style.backgroundRepeat = 'no-repeat';
      }
      if (backdrop) {
        backdrop.onload = () => {
          backdrop.style.opacity = '1';
        };
        backdrop.onerror = () => {
          const activeIndex = Number(backdrop.dataset.backdropIndex || '-1');
          if (activeIndex !== index) return;
          const fallbackUrl = applyCandidateAt(index + 1);
          if (fallbackUrl && typeof saveSettings === 'function') {
            saveSettings();
          }
        };
        backdrop.dataset.backdropIndex = String(index);
        backdrop.hidden = false;
        backdrop.classList.remove('hidden');
        backdrop.style.opacity = '1';
        if (backdrop.getAttribute('src') !== nextUrl) {
          backdrop.setAttribute('src', nextUrl);
        }
      }
      return nextUrl;
    };

    pushCandidate(preferredUrl);
    getWorkflowPageResultImageCandidates(page).forEach((item) => pushCandidate(item));

    return applyCandidateAt(0);
  }

  if (typeof buildEffectiveWorkflowPagePrompt === 'function') {
    const prevBuildEffectiveWorkflowPagePrompt = buildEffectiveWorkflowPagePrompt;
    buildEffectiveWorkflowPagePrompt = function buildEffectiveWorkflowPagePromptWithStrongerComposition(page) {
      const base = prevBuildEffectiveWorkflowPagePrompt(page);
      const next = typeof applyHarnessMetaToPage === 'function' ? applyHarnessMetaToPage({ ...page }) : (page || {});
      const extraBlocks = [];
      const hierarchyBlock = [
        '内容分层规则：先把这一页内容整理成最多 4 层，再决定排版。',
        '第 1 层是页标题。',
        '第 2 层是一级标题或一级信息组。',
        '第 3 层是二级标题，或一级标题的附属说明、数据、注释。',
        '第 4 层是二级标题的附属内容，只保留必要补充。',
        '结构关系只允许总分结构或并列结构，不要把不同层级混成一整段。',
        '必须通过层级、分组、容器、对比和留白把这 4 层以内的关系明确排出来。',
      ].join('\n');

      if (next.pageType === 'data') {
        extraBlocks.push([
          '数据页强化排版要求：禁止只做简单标题加正文堆砌，必须转成更像信息图的版式。',
          '至少满足以下三项中的两项：1. 放大一个核心数字或结论；2. 把内容拆成 2 到 4 个分组卡片；3. 用趋势、对比、流程、场景矩阵或案例模块承载信息。',
          '不同信息组必须有明确容器、层级和留白，不能做成单栏长文。',
        ].join('\n'));
      } else if (next.pageType === 'content') {
        extraBlocks.push([
          '内容页强化排版要求：避免整段正文直接平铺。',
          '优先拆成标题区 + 2 到 4 个信息模块、卡片、对比组或时间线，形成清晰主次。',
        ].join('\n'));
      }

      if (next.densityBand === 'dense' || next.layoutRisk === 'high') {
        extraBlocks.push('当前页内容偏密，必须优先做分组、分栏、卡片化或图示化处理，不允许把全部文字原样堆成大段。');
      }

      return [base, hierarchyBlock, ...extraBlocks].filter(Boolean).join('\n\n');
    };
  }

  renderWorkflowLayoutScroller = function renderWorkflowLayoutScrollerReadable(page) {
    const refs = ensureWorkflowWorkbenchCards();
    if (!refs.layoutScroller) return;
    const sections = buildWorkflowLayoutSections(page);
    refs.layoutScroller.innerHTML = sections.map((section, index) => `
      <details class="workflow-layout-section" ${index === 0 ? 'open' : ''}>
        <summary>
          <span>${escapeHtml(section.title)}</span>
        </summary>
        <div class="workflow-layout-section-body">${escapeHtml(section.body || section.summary || '').replace(/\n/g, '<br />')}</div>
      </details>
    `).join('');
  };

  function clampWorkflowSketchPoint(value) {
    const numeric = Number(value);
    if (!Number.isFinite(numeric)) return 0;
    return Math.max(0, Math.min(1, numeric));
  }

  function toWorkflowSketchStrokeWidth(size, ctx) {
    const pixels = Math.max(1, Number(size) || 3);
    const width = ctx?.canvas?.clientWidth || ctx?.canvas?.width || 1600;
    const height = ctx?.canvas?.clientHeight || ctx?.canvas?.height || 900;
    const basis = Math.max(width, height, 1);
    return pixels / basis;
  }

  drawWorkflowSketchShape = function drawWorkflowSketchShapeFinal(ctx, inputShape) {
    const shape = {
      ...inputShape,
      x1: clampWorkflowSketchPoint(inputShape?.x1),
      y1: clampWorkflowSketchPoint(inputShape?.y1),
      x2: clampWorkflowSketchPoint(inputShape?.x2),
      y2: clampWorkflowSketchPoint(inputShape?.y2),
      points: Array.isArray(inputShape?.points)
        ? inputShape.points.map((point) => ({
          x: clampWorkflowSketchPoint(point?.x),
          y: clampWorkflowSketchPoint(point?.y),
        }))
        : [],
    };

    ctx.save();
    ctx.strokeStyle = shape.color || "#498094";
    ctx.lineWidth = toWorkflowSketchStrokeWidth(shape.size, ctx);
    ctx.lineCap = "round";
    ctx.lineJoin = "round";

    if (shape.type === "pen" && shape.points.length > 1) {
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
  };

  renderWorkflowPlanHistory = function renderWorkflowPlanHistoryHidden() {
    if (!el.workflowPlanHistoryPanel) return;
    el.workflowPlanHistoryPanel.hidden = true;
    el.workflowPlanHistoryPanel.classList.add('hidden');
    el.workflowPlanHistoryPanel.style.display = 'none';
    if (el.workflowPlanDeleteBtn) {
      el.workflowPlanDeleteBtn.hidden = true;
      el.workflowPlanDeleteBtn.style.display = 'none';
    }
  };

  const prevRenderWorkflowPlanLibrary = renderWorkflowPlanLibrary;
  renderWorkflowPlanLibrary = function renderWorkflowPlanLibraryCompact() {
    prevRenderWorkflowPlanLibrary();
    renderWorkflowPlanHistory();
    hideSmartWorkflowChrome();
  };

  ensureWorkflowWorkbenchCards = function ensureWorkflowWorkbenchCardsFinal() {
    const grid = el.workflowDetailPanel?.querySelector('.workflow-detail-grid');
    const editorCard = document.querySelector('#workflowDetailDisplayTitleInput')?.closest('.theme-review-card');
    const layoutCard = document.querySelector('#workflowDetailLayout')?.closest('.theme-review-card');
    if (!grid || !editorCard || !layoutCard) return {};

    Array.from(document.querySelectorAll('#workflowDetailBoardCard')).slice(1).forEach((node) => node.remove());
    Array.from(document.querySelectorAll('#workflowDetailNotesCard')).slice(1).forEach((node) => node.remove());

    grid.classList.add('workflow-detail-workbench-grid');
    editorCard.classList.add('workflow-detail-editor-card');
    layoutCard.classList.add('workflow-detail-layout-card');

    const toolLabels = {
      select: '选择',
      pen: '画笔',
      eraser: '橡皮',
      line: '直线',
      rect: '框选',
      circle: '圆形',
      arrow: '箭头',
      undo: '撤销',
      redo: '重做',
      clear: '清空',
    };

    let boardCard = document.querySelector('#workflowDetailBoardCard');
    if (!boardCard) {
      boardCard = document.createElement('article');
      boardCard.id = 'workflowDetailBoardCard';
      boardCard.className = 'theme-review-card workflow-detail-board-card';
      boardCard.innerHTML = `
        <div class="workflow-detail-board-shell">
          <aside class="workflow-sketch-toolbar" id="workflowDetailSketchToolbar">
            <button type="button" class="workflow-sketch-tool active" data-tool="select" title="选择">选择</button>
            <button type="button" class="workflow-sketch-tool" data-tool="pen" title="画笔">画笔</button>
            <button type="button" class="workflow-sketch-tool" data-tool="eraser" title="橡皮">橡皮</button>
            <button type="button" class="workflow-sketch-tool" data-tool="line" title="直线">直线</button>
            <button type="button" class="workflow-sketch-tool" data-tool="rect" title="矩形框选">框选</button>
            <button type="button" class="workflow-sketch-tool" data-tool="circle" title="圆形">圆形</button>
            <button type="button" class="workflow-sketch-tool" data-tool="arrow" title="箭头">箭头</button>
            <div class="workflow-sketch-toolbar-divider"></div>
            <label class="workflow-sketch-color" title="颜色">
              <input id="workflowDetailSketchColor" type="color" value="#498094" />
            </label>
            <input id="workflowDetailSketchSize" type="range" min="1" max="10" value="3" title="线宽" />
            <div class="workflow-sketch-toolbar-divider"></div>
            <button type="button" class="workflow-sketch-tool" data-action="undo" title="撤销">撤销</button>
            <button type="button" class="workflow-sketch-tool" data-action="redo" title="重做">重做</button>
            <button type="button" class="workflow-sketch-tool" data-action="clear" title="清空">清空</button>
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

    const stage = boardCard.querySelector('.workflow-detail-board-stage');
    let canvas = boardCard.querySelector('#workflowDetailSketchCanvas');
    if (stage && !canvas) {
      canvas = document.createElement('canvas');
      canvas.id = 'workflowDetailSketchCanvas';
      canvas.className = 'workflow-detail-sketch-canvas';
      stage.appendChild(canvas);
    }
    if (stage && !boardCard.querySelector('#workflowDetailSketchBackdrop')) {
      const backdrop = document.createElement('img');
      backdrop.id = 'workflowDetailSketchBackdrop';
      backdrop.className = 'workflow-detail-sketch-backdrop hidden';
      backdrop.alt = '当前页结果图';
      stage.insertBefore(backdrop, canvas || stage.firstChild);
    }

    boardCard.querySelectorAll('.workflow-sketch-tool[data-tool]').forEach((button) => {
      const tool = button.dataset.tool || '';
      if (toolLabels[tool]) button.textContent = toolLabels[tool];
    });
    boardCard.querySelectorAll('.workflow-sketch-tool[data-action]').forEach((button) => {
      const action = button.dataset.action || '';
      if (toolLabels[action]) button.textContent = toolLabels[action];
    });

    let notesCard = document.querySelector('#workflowDetailNotesCard');
    if (!notesCard) {
      notesCard = document.createElement('article');
      notesCard.id = 'workflowDetailNotesCard';
      notesCard.className = 'theme-review-card workflow-detail-notes-card';
      notesCard.innerHTML = `
        <strong>额外要求 / 修改要求</strong>
        <label class="field workflow-detail-field">
          <textarea id="workflowDetailWorkspaceNotes" rows="5" class="workflow-detail-notes-editor" placeholder="例如：右上角补一个参数卡片；把圆形改成地球；整体更像信息图；保留一块区域给图表。"></textarea>
        </label>
      `;
      grid.insertBefore(notesCard, layoutCard);
    }

    let layoutScroller = document.querySelector('#workflowDetailLayoutScroller');
    if (!layoutScroller) {
      layoutScroller = document.createElement('div');
      layoutScroller.id = 'workflowDetailLayoutScroller';
      layoutScroller.className = 'workflow-layout-scroller';
      layoutCard.appendChild(layoutScroller);
    }

    if (el.workflowDetailLayout) {
      el.workflowDetailLayout.hidden = true;
      el.workflowDetailLayout.style.display = 'none';
    }

    return {
      grid,
      editorCard,
      layoutCard,
      boardCard,
      notesCard,
      layoutScroller,
      stage: boardCard.querySelector('.workflow-detail-board-stage'),
      canvas: boardCard.querySelector('#workflowDetailSketchCanvas'),
      backdrop: boardCard.querySelector('#workflowDetailSketchBackdrop'),
    };
  };

  renderWorkflowSketchCanvas = function renderWorkflowSketchCanvasFinal(page) {
    const refs = ensureWorkflowWorkbenchCards();
    const canvas = refs.canvas;
    if (!canvas || !page) return;
    const stage = refs.stage || canvas.parentElement;
    const backdrop = refs.backdrop;
    const backdropUrl = applyWorkflowPageResultBackdrop(page, stage, backdrop);
    const hasBackdrop = Boolean(backdropUrl);
    const displayWidth = Math.max(400, Math.round(stage.clientWidth || 640));
    const displayHeight = Math.max(225, Math.round(stage.clientHeight || (displayWidth * 9 / 16)));
    const dpr = window.devicePixelRatio || 1;

    canvas.width = Math.round(displayWidth * dpr);
    canvas.height = Math.round(displayHeight * dpr);
    canvas.style.width = `${displayWidth}px`;
    canvas.style.height = `${displayHeight}px`;

    const ctx = canvas.getContext('2d');
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, displayWidth, displayHeight);
    if (!hasBackdrop) {
      ctx.fillStyle = 'rgba(252, 253, 255, 0.82)';
      ctx.fillRect(0, 0, displayWidth, displayHeight);
    }
    ctx.strokeStyle = hasBackdrop ? 'rgba(19, 34, 56, 0.12)' : 'rgba(19, 34, 56, 0.08)';
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

    window.requestAnimationFrame(() => {
      if (state.workflowDetailPageId !== page.id) return;
      const nextWidth = Math.round(stage.clientWidth || 0);
      const nextHeight = Math.round(stage.clientHeight || 0);
      if (!nextWidth || !nextHeight) return;
      const currentWidth = Math.round(parseFloat(canvas.style.width || '0'));
      const currentHeight = Math.round(parseFloat(canvas.style.height || '0'));
      if (Math.abs(nextWidth - currentWidth) > 2 || Math.abs(nextHeight - currentHeight) > 2) {
        renderWorkflowSketchCanvas(page);
      }
    });
  };

  setupWorkflowSketchInteractions = function setupWorkflowSketchInteractionsFinal(page) {
    const refs = ensureWorkflowWorkbenchCards();
    const canvas = refs.canvas;
    const toolbar = refs.boardCard?.querySelector('#workflowDetailSketchToolbar');
    const colorInput = refs.boardCard?.querySelector('#workflowDetailSketchColor');
    const sizeInput = refs.boardCard?.querySelector('#workflowDetailSketchSize');
    const notesInput = document.querySelector('#workflowDetailWorkspaceNotes');
    const generateBtn = refs.boardCard?.querySelector('#workflowDetailBoardGenerateBtn');
    const reviseBtn = refs.boardCard?.querySelector('#workflowDetailBoardReviseBtn');

    if (!canvas || !toolbar || !page) return;
    workflowDetailSketchUi.pageId = page.id;
    page.detailSketchShapes = normalizeWorkflowSketchShapes(page.detailSketchShapes);
    page.detailSketchRedo = normalizeWorkflowSketchShapes(page.detailSketchRedo);

    toolbar.querySelectorAll('.workflow-sketch-tool[data-tool]').forEach((button) => {
      button.classList.toggle('active', button.dataset.tool === workflowDetailSketchUi.tool);
      button.onclick = () => {
        workflowDetailSketchUi.tool = button.dataset.tool;
        setupWorkflowSketchInteractions(page);
      };
    });

    toolbar.querySelectorAll('.workflow-sketch-tool[data-action]').forEach((button) => {
      button.onclick = () => {
        if (button.dataset.action === 'undo' && page.detailSketchShapes.length) {
          page.detailSketchRedo.push(page.detailSketchShapes.pop());
        }
        if (button.dataset.action === 'redo' && page.detailSketchRedo.length) {
          page.detailSketchShapes.push(page.detailSketchRedo.pop());
        }
        if (button.dataset.action === 'clear') {
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
      colorInput.oninput = () => { workflowDetailSketchUi.color = colorInput.value; };
    }
    if (sizeInput) {
      sizeInput.value = String(workflowDetailSketchUi.size);
      sizeInput.oninput = () => { workflowDetailSketchUi.size = Number(sizeInput.value) || 3; };
    }
    if (notesInput) {
      notesInput.value = page.detailWorkspaceNotes || '';
      notesInput.oninput = () => {
        page.detailWorkspaceNotes = notesInput.value;
        saveSettings();
        updatePayloadPreview();
      };
    }

    if (generateBtn) {
      generateBtn.disabled = page.status === 'running' || state.workflowRunning || page.layoutStatus === 'running' || getWorkflowDetailDraft(page).dirty;
      generateBtn.onclick = () => generateWorkflowPage(page.id);
    }

    if (reviseBtn) {
      reviseBtn.disabled = !getWorkflowPageResultImageUrl(page);
      reviseBtn.onclick = async () => {
        const latestResultImageUrl = getWorkflowPageResultImageUrl(page);
        if (!latestResultImageUrl) {
          setStatus('请先生成本页，再打开改图。', 'error');
          return;
        }
        state.imagePreview = {
          title: `第 ${page.pageNumber} 页结果`,
          url: latestResultImageUrl,
        };
        if (el.imagePreviewPrompt) {
          el.imagePreviewPrompt.value = String(page.detailWorkspaceNotes || '').trim();
        }
        await sendImagePreviewToRevise({ submitImmediately: false });
      };
    }

    canvas.onpointerdown = (event) => {
      const tool = workflowDetailSketchUi.tool;
      if (tool === 'select') return;
      const point = getWorkflowSketchCanvasPoint(event, canvas);
      if (tool === 'eraser') {
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
      if (tool === 'pen') workflowDetailSketchUi.draftShape.points = [point];
      canvas.setPointerCapture?.(event.pointerId);
      renderWorkflowSketchCanvas(page);
    };

    canvas.onpointermove = (event) => {
      if (!workflowDetailSketchUi.drawing || !workflowDetailSketchUi.draftShape) return;
      const point = getWorkflowSketchCanvasPoint(event, canvas);
      workflowDetailSketchUi.draftShape.x2 = point.x;
      workflowDetailSketchUi.draftShape.y2 = point.y;
      if (workflowDetailSketchUi.draftShape.type === 'pen') {
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
      if (workflowDetailSketchUi.drawing && workflowDetailSketchUi.tool === 'pen') {
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
  };

  renderWorkflowDetail = function renderWorkflowDetailWorkbench() {
    const page = state.workflowPages.find((item) => item.id === state.workflowDetailPageId);
    if (!page) {
      closeWorkflowResearchPopover();
      el.workflowDetailModal?.classList.add('hidden');
      el.workflowDetailModal?.setAttribute('aria-hidden', 'true');
      return;
    }

    const shell = ensureWorkflowDetailWorkspaceShell();
    const promptPopover = ensureWorkflowPromptPopover();
    const refs = ensureWorkflowWorkbenchCards();
    normalizeWorkflowDetailHeadings();
    const previousBackdropUrl = normalizeWorkflowResultImageUrl(page.detailBackdropUrl || page.detail_backdrop_url);
    const stage = refs.stage;
    const backdropNode = refs.backdrop;
    const resolvedBackdropUrl = applyWorkflowPageResultBackdrop(page, stage, backdropNode);
    if (resolvedBackdropUrl && previousBackdropUrl !== resolvedBackdropUrl) {
      saveSettings();
    }
    const draft = getWorkflowDetailDraft(page);

    const updateDetailDerivedContent = () => {
      const hasPendingDraft = getWorkflowDetailDraft(page).dirty;
      el.workflowDetailMeta.textContent = `${PAGE_TYPE_LABELS[page.pageType] || page.pageType} · ${getWorkflowDisplayTitle(page) || page.pageTitle || `第 ${page.pageNumber} 页`} · 装饰 ${getDecorationLevelLabel(getPageDecorationLevel(page))} · 输出 ${getPptOutputDescription()}`;
      renderWorkflowLayoutScroller(page);

      [
        el.workflowDetailUseBtn,
        el.workflowDetailCopyBtn,
        el.workflowDetailVisibleResetBtn,
        el.workflowDetailConfirmStatus,
        el.workflowDetailResearchStatus,
        el.workflowDetailResearchList,
        el.workflowDetailResearchApplyBtn,
      ].forEach((node) => {
        if (!node) return;
        node.hidden = true;
        node.style.display = 'none';
      });

      setWorkflowDetailSectionVisibility(el.workflowDetailSuggestedContent, false);
      setWorkflowDetailSectionVisibility(el.workflowDetailConfirmedContent, false);
      setWorkflowDetailSectionVisibility(el.workflowDetailContent, false);

      el.workflowDetailResearchQueryInput?.closest('.field')?.setAttribute('hidden', 'hidden');
      el.workflowDetailResearchBtn?.closest('.action-row')?.setAttribute('hidden', 'hidden');
      el.workflowDetailTheme?.closest('.theme-review-card')?.setAttribute('hidden', 'hidden');
      el.workflowDetailPrompt?.closest('.theme-review-card')?.setAttribute('hidden', 'hidden');

      if (el.workflowDetailRunBtn) {
        el.workflowDetailRunBtn.hidden = true;
        el.workflowDetailRunBtn.disabled = page.status === 'running' || state.workflowRunning || page.layoutStatus === 'running' || hasPendingDraft;
      }
      if (el.workflowDetailConfirmContentBtn) {
        el.workflowDetailConfirmContentBtn.disabled = !hasPendingDraft || page.layoutStatus === 'running';
        el.workflowDetailConfirmContentBtn.textContent = page.layoutStatus === 'running' ? '重排中...' : '确认内容';
      }
      if (shell.researchToggleBtn) {
        shell.researchToggleBtn.disabled = page.researchStatus === 'running';
        shell.researchToggleBtn.textContent = page.researchStatus === 'running' ? '联网补充中...' : '联网补充';
      }
      if (promptPopover.toggleBtn) {
        promptPopover.toggleBtn.disabled = page.layoutStatus === 'running' && !getWorkflowOriginalPrompt(page);
        promptPopover.toggleBtn.textContent = '查看原始提示词';
      }

      setupWorkflowSketchInteractions(page);
      renderWorkflowResearchPopover(page);
      renderWorkflowPromptPopover(page);
    };

    if (el.workflowDetailTitle) el.workflowDetailTitle.textContent = `第 ${page.pageNumber} 页详情`;
    if (el.workflowDetailDisplayTitleInput) {
      el.workflowDetailDisplayTitleInput.value = draft.title;
      el.workflowDetailDisplayTitleInput.oninput = () => {
        updateWorkflowDetailDraft(page, { title: el.workflowDetailDisplayTitleInput.value });
        updateDetailDerivedContent();
      };
    }
    if (el.workflowDetailDecorationLevel) {
      el.workflowDetailDecorationLevel.value = getPageDecorationLevel(page);
      el.workflowDetailDecorationLevel.onchange = () => {
        page.decorationLevel = normalizeDecorationLevel(el.workflowDetailDecorationLevel.value);
        saveSettings();
        renderWorkflowPlan();
        renderWorkflowDetail();
        updatePayloadPreview();
        setStatus(`第 ${page.pageNumber} 页的装饰强度已调整为${getDecorationLevelLabel(page.decorationLevel)}。`, 'success');
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
      el.workflowDetailConfirmContentBtn.onclick = async () => {
        const suggestedPlan = buildWorkflowSuggestedTextPlan(page);
        const suggestedBody = suggestedPlan.visibleLines.join('\n').trim();
        const nextTitle = draft.title.trim();
        const nextBody = draft.body.trim();
        page.displayTitleOverride = nextTitle === String(page.pageTitle || '').trim() ? null : draft.title;
        page.displayBodyOverride = nextBody === suggestedBody ? null : draft.body;
        primeWorkflowDetailDraft(page);
        saveSettings();
        renderWorkflowPlan();
        renderWorkflowDetail();
        updatePayloadPreview();
        setStatus(`第 ${page.pageNumber} 页内容已确认，正在根据新内容重新排版。`, 'running');
        try {
          await ensureWorkflowPageDesign(page, { force: true });
          updatePayloadPreview();
          setStatus(`第 ${page.pageNumber} 页内容已确认，并已按新内容重新排版。`, 'success');
        } catch (error) {
          updatePayloadPreview();
          setStatus(`第 ${page.pageNumber} 页内容已确认，但重新排版失败：${error.message || '请稍后重试。'}`, 'error');
        }
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
    if (promptPopover.toggleBtn) {
      promptPopover.toggleBtn.onclick = () => {
        const isSamePageOpen = state.workflowPromptPopoverOpen && state.workflowPromptPopoverPageId === page.id;
        state.workflowPromptPopoverOpen = !isSamePageOpen;
        state.workflowPromptPopoverPageId = isSamePageOpen ? '' : page.id;
        updateDetailDerivedContent();
      };
    }
    if (el.workflowDetailCloseBtn) {
      el.workflowDetailCloseBtn.onclick = () => {
        closeWorkflowResearchPopover();
        state.workflowPromptPopoverOpen = false;
        state.workflowPromptPopoverPageId = '';
        closeWorkflowDetail();
      };
    }

    updateDetailDerivedContent();
  };

  hideSmartWorkflowChrome();
  renderWorkflowPlanHistory();
  renderWorkflowPlanLibrary();
  renderWorkflowDetail();
}());
