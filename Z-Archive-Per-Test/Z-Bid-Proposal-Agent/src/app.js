const form = document.getElementById('proposalForm');
const resetBtn = document.getElementById('resetBtn');
const composeBtn = document.getElementById('composeBtn');
const summaryOutput = document.getElementById('summaryOutput');
const resultLinks = document.getElementById('resultLinks');
const outlineOutput = document.getElementById('outlineOutput');
const outlineStatus = document.getElementById('outlineStatus');
const composeProgress = document.getElementById('composeProgress');
const composeStatus = document.getElementById('composeStatus');
const composeSteps = document.getElementById('composeSteps');
const composeProgressBar = document.getElementById('composeProgressBar');
const composeProgressPercent = document.getElementById('composeProgressPercent');
const composeProgressModel = document.getElementById('composeProgressModel');
const composeLiveDetail = document.getElementById('composeLiveDetail');

const fileInputs = ['technicalSpecification', 'scoringCriteria', 'templateFile'];
const defaultTemplatePath = 'templates/default-proposal-template.docx';
const currentJobStorageKey = 'bidProposalAgent.currentJobId';
let currentJobId = null;
let composePollTimer = null;

function formatFile(file) {
    if (!file) {
        return '未选择文件';
    }

    const sizeKb = Math.max(1, Math.round(file.size / 1024));
    return `${file.name} (${sizeKb} KB)`;
}

function updateFileLabel(input) {
    const label = document.querySelector(`[data-file-name="${input.id}"]`);
    if (label) {
        label.textContent = formatFile(input.files[0]);
    }
}

function collectSummary() {
    const technicalSpecification = document.getElementById('technicalSpecification').files[0];
    const scoringCriteria = document.getElementById('scoringCriteria').files[0];
    const templateFile = document.getElementById('templateFile').files[0];
    const outputType = document.getElementById('outputType').value;
    const projectName = document.getElementById('projectName').value.trim();
    const headingRules = collectHeadingRules();
    const pagePlan = collectPagePlan();

    return {
        projectName: projectName || '未填写',
        outputType,
        requiredInputs: {
            technicalSpecification: formatFile(technicalSpecification),
            scoringCriteria: formatFile(scoringCriteria)
        },
        optionalInputs: {
            proposalTemplate: templateFile ? formatFile(templateFile) : `未上传，使用默认模板：${defaultTemplatePath}`
        },
        templateSource: templateFile ? 'uploaded' : 'default',
        templateStrategy: templateFile
            ? '优先提取模板中的章节结构、标题层级、表格样式和固定措辞，生成内容时保持模板格式。'
            : '使用项目内置 Word 默认方案模板作为兜底，继承默认标题层级和段落间距，减少无模板场景下的格式调整。',
        postProcessing: {
            headingRules,
            pagePlan
        }
    };
}

function buildFormData(summary) {
    const formData = new FormData();
    formData.append('config', JSON.stringify(summary));

    fileInputs.forEach((id) => {
        const input = document.getElementById(id);
        if (input.files[0]) {
            formData.append(id, input.files[0]);
        }
    });

    return formData;
}

function renderResultLinks(payload) {
    if (!payload.manifestUrl) {
        resultLinks.hidden = true;
        resultLinks.innerHTML = '';
        return;
    }

    resultLinks.hidden = false;
    resultLinks.innerHTML = `
        <a href="${payload.manifestUrl}" target="_blank" rel="noreferrer">查看任务清单</a>
        ${payload.draftUrl ? `<a href="${payload.draftUrl}" target="_blank" rel="noreferrer">查看方案草稿</a>` : ''}
        ${payload.finalDraftUrl ? `<a href="${payload.finalDraftUrl}" target="_blank" rel="noreferrer">查看正式编写稿</a>` : ''}
        ${payload.finalDocxUrl ? `<a class="download-link" href="${payload.finalDocxUrl}" download>下载 Word 方案</a>` : ''}
        ${payload.pipelineUrl ? `<a href="${payload.pipelineUrl}" target="_blank" rel="noreferrer">查看智能体流水线</a>` : ''}
    `;
}

function rememberCurrentJob(jobId) {
    if (jobId) {
        window.localStorage.setItem(currentJobStorageKey, jobId);
    } else {
        window.localStorage.removeItem(currentJobStorageKey);
    }
}

function setComposeReady(jobId) {
    currentJobId = jobId || null;
    rememberCurrentJob(currentJobId);
    composeBtn.disabled = !currentJobId;
    if (currentJobId) {
        resetComposeProgress('大纲已生成，等待确认');
        composeProgress.hidden = false;
    }
}

function resetComposeProgress(status = '等待确认') {
    composeStatus.textContent = status;
    composeProgress.hidden = true;
    updateProgressMeter({ progress: 0, model: '等待模型', message: status });
    composeSteps.querySelectorAll('li').forEach((item) => {
        item.className = '';
    });
}

function setComposeStep(step, status) {
    composeProgress.hidden = false;
    composeStatus.textContent = status;
    const order = ['confirm', 'compose', 'export', 'done'];
    const activeIndex = order.indexOf(step);

    composeSteps.querySelectorAll('li').forEach((item) => {
        const index = order.indexOf(item.dataset.step);
        item.className = index < activeIndex ? 'done' : '';
        if (index === activeIndex) {
            item.className = 'active';
        }
    });
}

function updateProgressMeter(payload) {
    const progress = Math.max(0, Math.min(100, Number(payload.progress || 0)));
    const model = payload.model || payload.engine || '未选择模型';
    const details = [
        payload.message || '等待任务状态',
        payload.currentSection ? `章节：${payload.currentSection}` : '',
        payload.currentSubtitle ? `小节：${payload.currentSubtitle}` : '',
        payload.modelElapsedSeconds ? `模型调用耗时：${payload.modelElapsedSeconds} 秒` : ''
    ].filter(Boolean);

    if (composeProgressBar) {
        composeProgressBar.style.width = `${progress}%`;
    }
    if (composeProgressPercent) {
        composeProgressPercent.textContent = `${progress}%`;
    }
    if (composeProgressModel) {
        composeProgressModel.textContent = model;
    }
    if (composeLiveDetail) {
        composeLiveDetail.textContent = details.join(' / ');
    }
}

function stopComposePolling() {
    if (composePollTimer) {
        window.clearInterval(composePollTimer);
        composePollTimer = null;
    }
}

function renderOutline(outline = []) {
    if (!outline.length) {
        outlineStatus.textContent = '等待生成';
        outlineOutput.className = 'outline-output empty';
        outlineOutput.textContent = '上传材料并生成后，会在这里展示可确认的大纲。';
        return;
    }

    outlineStatus.textContent = `${outline.length} 个章节`;
    outlineOutput.className = 'outline-output';
    outlineOutput.innerHTML = outline.map((section, index) => {
        const scoringItems = section.relatedScoringItems || [];
        const scoringMarkup = scoringItems.length
            ? `<ul>${scoringItems.map((item) => `<li>${escapeHtml(item)}</li>`).join('')}</ul>`
            : '<p class="outline-muted">暂无关联评分项</p>';

        return `
            <article class="outline-item">
                <div class="outline-index">${index + 1}</div>
                <div class="outline-content">
                    <h4>${escapeHtml(section.title)}</h4>
                    <p>${escapeHtml(section.writingGoal || '待补充写作目标')}</p>
                    ${renderChildren(section.children || [])}
                    <div class="outline-scoring">
                        <strong>关联评分项</strong>
                        ${scoringMarkup}
                    </div>
                </div>
            </article>
        `;
    }).join('');
}

function renderChildren(children) {
    if (!children.length) {
        return '';
    }

    return `
        <div class="outline-children">
            <strong>子标题规划</strong>
            <ol>
                ${children.slice(0, 7).map((child) => `<li>${escapeHtml(child.title)}</li>`).join('')}
            </ol>
        </div>
    `;
}

function escapeHtml(value) {
    return String(value)
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#039;');
}

function collectHeadingRules() {
    const specificRules = [
        {
            title: document.getElementById('chapterTitleA').value.trim(),
            targetLevel: document.getElementById('chapterLevelA').value
        },
        {
            title: document.getElementById('chapterTitleB').value.trim(),
            targetLevel: document.getElementById('chapterLevelB').value
        }
    ].filter((rule) => rule.title && rule.targetLevel);

    return {
        globalAction: document.getElementById('globalHeadingAction').value,
        maxLevel: Number(document.getElementById('maxHeadingLevel').value),
        specificRules: specificRules.map((rule) => ({
            matchTitle: rule.title,
            targetLevel: Number(rule.targetLevel)
        })),
        conflictPolicy: 'specificRulesOverrideGlobalAction'
    };
}

function collectPagePlan() {
    const totalPagesValue = document.getElementById('totalPages').value;
    const totalPages = totalPagesValue ? Number(totalPagesValue) : null;
    const chapterPages = [
        {
            title: document.getElementById('pageChapterTitleA').value.trim(),
            pages: document.getElementById('pageChapterCountA').value
        },
        {
            title: document.getElementById('pageChapterTitleB').value.trim(),
            pages: document.getElementById('pageChapterCountB').value
        },
        {
            title: document.getElementById('pageChapterTitleC').value.trim(),
            pages: document.getElementById('pageChapterCountC').value
        }
    ]
        .filter((rule) => rule.title && rule.pages)
        .map((rule) => ({
            chapterTitle: rule.title,
            pages: Number(rule.pages)
        }));

    const assignedPages = chapterPages.reduce((sum, rule) => sum + rule.pages, 0);

    return {
        totalPages,
        allocationMode: document.getElementById('pageAllocationMode').value,
        chapterPages,
        assignedPages,
        remainingPages: totalPages === null ? null : totalPages - assignedPages,
        generationRule: 'Use totalPages as the target document length. Auto allocate by outline when no manual chapter pages are provided; manual chapter page settings override auto allocation for matched sections.'
    };
}

fileInputs.forEach((id) => {
    const input = document.getElementById(id);
    input.addEventListener('change', () => updateFileLabel(input));
});

form.addEventListener('submit', (event) => {
    event.preventDefault();

    if (!form.reportValidity()) {
        return;
    }

    const summary = collectSummary();
    setComposeReady(null);
    resetComposeProgress();
    renderOutline([]);
    outlineStatus.textContent = '生成中...';
    summaryOutput.textContent = '正在提交生成任务...';

    fetch('/api/proposals', {
        method: 'POST',
        body: buildFormData(summary)
    })
        .then((response) => response.json().then((payload) => ({ ok: response.ok, payload })))
        .then(({ ok, payload }) => {
            if (!ok) {
                throw new Error(payload.error || '后端处理失败');
            }

            renderResultLinks(payload);
            renderOutline(payload.outline || []);
            setComposeReady(payload.jobId);
            rememberCurrentJob(payload.jobId);
            summaryOutput.textContent = JSON.stringify(payload, null, 2);
        })
        .catch((error) => {
            renderResultLinks({});
            setComposeReady(null);
            renderOutline([]);
            summaryOutput.textContent = JSON.stringify({
                mode: 'local-preview',
                warning: `未连接后端，仅显示本地摘要：${error.message}`,
                summary
            }, null, 2);
        });
});

composeBtn.addEventListener('click', () => {
    if (!currentJobId) {
        return;
    }

    composeBtn.disabled = true;
    composeBtn.textContent = '后台编写中...';
    setComposeStep('confirm', '已确认大纲');
    summaryOutput.textContent = '已确认大纲，正在启动后台智能体写作任务...';

    window.setTimeout(() => {
        setComposeStep('compose', '正在启动后台写作任务');
    }, 180);

    fetch(`/api/proposals/${currentJobId}/compose`, {
        method: 'POST'
    })
        .then((response) => response.json().then((payload) => ({ ok: response.ok, payload })))
        .then(({ ok, payload }) => {
            if (!ok) {
                throw new Error(payload.error || '正式编写失败');
            }

            renderResultLinks(payload);
            renderComposeStatus(payload);
            startComposePolling(currentJobId);
        })
        .catch((error) => {
            composeBtn.disabled = false;
            composeStatus.textContent = '编写失败';
            summaryOutput.textContent = JSON.stringify({
                error: error.message,
                jobId: currentJobId
            }, null, 2);
        })
});

function startComposePolling(jobId) {
    stopComposePolling();
    fetchComposeStatus(jobId);
    composePollTimer = window.setInterval(() => {
        fetchComposeStatus(jobId);
    }, 1200);
}

function fetchComposeStatus(jobId) {
    fetch(`/api/proposals/${jobId}/compose/status`)
        .then((response) => response.json().then((payload) => ({ ok: response.ok, payload })))
        .then(({ ok, payload }) => {
            if (!ok) {
                throw new Error(payload.error || '获取编写状态失败');
            }

            renderComposeStatus(payload);
            if (payload.state === 'completed') {
                stopComposePolling();
                renderResultLinks(payload);
                setComposeStep('done', '编写完成，可下载 Word 方案');
                composeBtn.textContent = '确认大纲并开始编写';
                composeBtn.disabled = false;
            }
            if (payload.state === 'failed') {
                stopComposePolling();
                composeStatus.textContent = '编写失败';
                composeBtn.textContent = '确认大纲并开始编写';
                composeBtn.disabled = false;
            }
        })
        .catch((error) => {
            stopComposePolling();
            composeStatus.textContent = '状态获取失败';
            summaryOutput.textContent = JSON.stringify({
                error: error.message,
                jobId
            }, null, 2);
            composeBtn.textContent = '确认大纲并开始编写';
            composeBtn.disabled = false;
        });
}

function renderComposeStatus(payload) {
    const state = payload.state || 'running';
    if (state === 'queued') {
        setComposeStep('confirm', '任务已排队');
    } else if (state === 'running') {
        setComposeStep(payload.stage === 'export' ? 'export' : 'compose', payload.message || '正在编写');
    } else if (state === 'completed') {
        setComposeStep('done', '编写完成，可下载 Word 方案');
    } else if (state === 'failed') {
        composeStatus.textContent = payload.message || '编写失败';
    }
    updateProgressMeter(payload);

    const lines = [
        `状态：${payload.message || state}`,
        `进度：${payload.progress || 0}%`,
        `写作引擎：${payload.engine || 'unknown'}`,
        `当前模型：${payload.model || 'unknown'}`,
        payload.warning ? `提示：${payload.warning}` : '',
        payload.currentSection ? `当前章节：${payload.currentSection}` : '',
        payload.currentSubtitle ? `当前小节：${payload.currentSubtitle}` : '',
        payload.modelElapsedSeconds ? `模型调用耗时：${payload.modelElapsedSeconds} 秒` : '',
        payload.heartbeatAt ? `最近心跳：${payload.heartbeatAt}` : '',
        payload.pageAllocations ? `\n篇幅分配：\n${(payload.pageAllocations || []).map((item) => `${item.title}：${item.pages}页（${item.source === 'manual' ? '手动' : '自动'}）`).join('\n')}` : '',
        payload.finalDocxUrl ? `\nWord 下载：${payload.finalDocxUrl}` : '',
        `\n原始状态：\n${JSON.stringify(payload, null, 2)}`
    ].filter(Boolean);

    summaryOutput.textContent = lines.join('\n');
}

function renderPageAllocations(pageAllocations, payload) {
    if (!pageAllocations.length) {
        return JSON.stringify(payload, null, 2);
    }

    const allocationSummary = pageAllocations
        .map((item) => `${item.title}：${item.pages}页（${item.source === 'manual' ? '手动' : '自动'}）`)
        .join('\n');

    return `正式编写已完成，Word 文件已生成。\n\n篇幅分配：\n${allocationSummary}\n\n原始响应：\n${JSON.stringify(payload, null, 2)}`;
}

resetBtn.addEventListener('click', () => {
    stopComposePolling();
    form.reset();
    fileInputs.forEach((id) => updateFileLabel(document.getElementById(id)));
    renderResultLinks({});
    setComposeReady(null);
    rememberCurrentJob(null);
    resetComposeProgress();
    renderOutline([]);
    summaryOutput.textContent = '等待上传文件...';
});

function restoreCurrentJob() {
    const storedJobId = window.localStorage.getItem(currentJobStorageKey);
    const latestUrl = storedJobId
        ? `/api/proposals/${storedJobId}`
        : '/api/proposals/latest';

    fetch(latestUrl)
        .then((response) => response.json().then((payload) => ({ ok: response.ok, payload })))
        .then(({ ok, payload }) => {
            if (!ok) {
                throw new Error(payload.error || '暂无可恢复任务');
            }

            const jobId = payload.jobId || storedJobId;
            setComposeReady(jobId);
            composeBtn.textContent = '确认大纲并开始编写';
            const composeStatusPayload = payload.composeStatus || payload;
            renderResultLinks(payload.composeStatus ? payload : composeStatusPayload);
            if (payload.outline) {
                renderOutline(payload.outline);
            }
            renderComposeStatus(composeStatusPayload);

            if (['queued', 'running'].includes(composeStatusPayload.state)) {
                composeBtn.disabled = true;
                composeBtn.textContent = '后台编写中...';
                startComposePolling(jobId);
            }

            summaryOutput.textContent = JSON.stringify(payload, null, 2);
        })
        .catch(() => {
            resetComposeProgress();
        });
}

restoreCurrentJob();
