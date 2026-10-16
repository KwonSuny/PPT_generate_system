let previewData = [], activePage = 0, previewFingerprint = '';
      let latestState = {}, viewingVersion = '', dialogueFingerprint = '', formDirty = false;
      const artifactImage = url => /^\/api\/artifacts\/[A-Za-z0-9-]+\/download\?inline=1$/.test(url || '');
      function showWorkbenchError(message) { document.getElementById('chat-status').textContent = message; const alert = document.getElementById('workbench-alert'); alert.textContent = message; alert.hidden = false; }
      function renderPreview(preview) {
        const fingerprint = JSON.stringify(preview || []);
        if (fingerprint === previewFingerprint) return;
        previewFingerprint = fingerprint; previewData = preview || [];
        activePage = Math.max(0, Math.min(activePage, previewData.length - 1));
        drawPageStrip(); drawActiveSlide(); updateEditorMode();
      }
      function selectedPageNumbers(text) {
        if (!text.trim()) return [];
        if (text.length > 2000) return null;
        const values = new Set();
        for (const part of text.replace(/\s*[-–—－]\s*/g, '-').split(/[,，;；\s]+/).filter(Boolean)) {
          if (!/^\d+(?:-\d+)?$/.test(part)) return null;
          const bounds = part.split('-').map(Number).sort((a, b) => a - b);
          const start = bounds[0], end = bounds[bounds.length - 1];
          if (start < 1 || end > previewData.length) return null;
          for (let n = start; n <= end; n++) values.add(n);
        }
        return Array.from(values).sort((a, b) => a - b);
      }
      function drawPageStrip() {
        const strip = document.getElementById('page-strip');
        const fragment = document.createDocumentFragment();
        const parsed = selectedPageNumbers(document.getElementById('modify_page').value || '');
        const selected = new Set((parsed || []).map(String));
        previewData.forEach((slide, index) => {
          const button = document.createElement('div');
          button.className = 'page-thumb' + (index === activePage ? ' active' : '');
          const pageButton = document.createElement('button'); pageButton.type = 'button'; pageButton.className = 'thumb-preview';
          pageButton.setAttribute('aria-label', `预览第 ${index + 1} 页：${slide.title || '未命名'}`);
          pageButton.setAttribute('aria-current', index === activePage ? 'page' : 'false');
          const paper = document.createElement('span'); paper.className = 'thumb-paper';
          if (artifactImage(slide.image_url)) {
            const image = document.createElement('img'); image.src = slide.image_url; image.alt = ''; image.loading = 'lazy'; paper.append(image);
          } else { paper.textContent = String(index + 1); }
          const title = document.createElement('div'); title.className = 'thumb-title'; title.textContent = `${index + 1}. ${slide.title || '未命名'}`;
          pageButton.append(paper, title); button.append(pageButton);
          pageButton.addEventListener('click', () => { activePage = index; drawPageStrip(); drawActiveSlide(); });
          // 页面选择与预览分开，勾选仅用于下一次指定页修改。
          const check = document.createElement('input'); check.type = 'checkbox'; check.className = 'thumb-check';
          check.checked = selected.has(String(index + 1)); check.disabled = Boolean(viewingVersion && viewingVersion !== latestState.version_id);
          check.setAttribute('aria-label', `选择第 ${index + 1} 页用于修改`);
          check.addEventListener('click', event => event.stopPropagation());
          check.addEventListener('change', () => {
            const input = document.getElementById('modify_page');
            const parsed = selectedPageNumbers(input.value);
            if (parsed === null) { showWorkbenchError('当前页码输入无效，请先修正再勾选页面。'); drawPageStrip(); return; }
            const values = new Set(parsed);
            if (check.checked) values.add(index + 1); else values.delete(index + 1);
            input.value = Array.from(values).sort((a, b) => a - b).join(',');
            if ((document.getElementById('mode').value || 'new') === 'new') document.getElementById('mode').value = 'modify_text';
            formDirty = true; updateEditorMode();
          });
          button.append(check); fragment.append(button);
        });
        strip.replaceChildren(fragment);
      }
      function drawActiveSlide() {
        const container = document.getElementById('ppt-preview');
        if (!container) return;
        document.getElementById('page-counter').textContent = previewData.length ? `${activePage + 1} / ${previewData.length}` : '暂无页面';
        document.getElementById('prev-page').disabled = activePage <= 0;
        document.getElementById('next-page').disabled = activePage >= previewData.length - 1;
        document.getElementById('expand-preview').disabled = !previewData[activePage] || !artifactImage(previewData[activePage].image_url);
        if (!previewData.length) {
          const empty = document.createElement('div'); empty.className = 'empty-stage';
          const icon = document.createElement('div'); icon.className = 'empty-icon';
          const title = document.createElement('h2'); title.textContent = '暂无可预览的页面';
          const hint = document.createElement('p'); hint.textContent = '先填写主题和要求，或从左侧打开一份已保存的文稿。';
          empty.append(icon, title, hint); container.replaceChildren(empty); return;
        }
        const slide = previewData[activePage];
        const card = document.createElement('div'); card.className = 'preview-slide';
        const title = document.createElement('h4'); title.textContent = `${activePage + 1}. ${slide.title || '未命名页'}`;
        const body = document.createElement('p'); body.textContent = slide.text || '暂无正文内容';
        card.append(title, body);
        if (artifactImage(slide.image_url)) {
          body.hidden = true;
          const frame = document.createElement('div'); frame.className = 'preview-image';
          const image = document.createElement('img'); image.src = slide.image_url; image.alt = `第 ${activePage + 1} 页成品预览`;
          frame.append(image);
          if (document.getElementById('show-issues').checked) (slide.issues || []).forEach(issue => {
            const box = issue.bbox;
            if (!Array.isArray(box) || box.length !== 4 || !box.every(value => Number.isFinite(value) && value >= 0 && value <= 1) || box[2] <= 0 || box[3] <= 0 || box[0] + box[2] > 1.001 || box[1] + box[3] > 1.001) return;
            const marker = document.createElement('div'); marker.className = 'preview-issue';
            marker.style.left = `${box[0] * 100}%`; marker.style.top = `${box[1] * 100}%`;
            marker.style.width = `${box[2] * 100}%`; marker.style.height = `${box[3] * 100}%`;
            marker.title = issue.message || ''; frame.append(marker);
          });
          card.append(frame);
        } else { title.textContent += ' · 已保存内容摘要（暂无成品截图）'; }
        const review = document.createElement('div'); review.className = 'preview-review';
        review.textContent = slide.review_status === 'passed' ? '视觉模型未发现问题；仍需人工验收。' :
          (slide.review_status === 'needs_review' ? '页面待复核。' : '页面未执行视觉审核。');
        (slide.issues || []).concat(slide.rule_issues || []).forEach(issue => {
          const line = document.createElement('div'); line.textContent = issue.message || ''; review.append(line);
        });
        card.append(review); container.replaceChildren(card);
      }
      function changePage(delta) {
        activePage = Math.max(0, Math.min(activePage + delta, previewData.length - 1)); drawPageStrip(); drawActiveSlide();
      }
      function expandPreview() {
        const slide = previewData[activePage]; if (!slide || !artifactImage(slide.image_url)) return;
        document.getElementById('expanded-image').src = slide.image_url;
        document.getElementById('expanded-title').textContent = `第 ${activePage + 1} 页：${slide.title}`;
        document.getElementById('preview-dialog').showModal();
      }
      function toggleLibrary() { document.getElementById('library-panel').classList.toggle('open'); }
      function switchPanel(panel) {
        ['editor', 'dialogue', 'design'].forEach(name => {
          document.getElementById(name + '-panel').hidden = name !== panel;
          document.getElementById(name + '-tab').setAttribute('aria-selected', String(name === panel));
        });
        if (panel === 'dialogue') loadDialogue();
        if (panel === 'design') refreshDesignCandidates();
      }
      function useCurrentPage() {
        if (!previewData.length) return;
        document.getElementById('modify_page').value = String(activePage + 1);
        if ((document.getElementById('mode').value || 'new') === 'new') document.getElementById('mode').value = 'modify_text';
        formDirty = true; updateEditorMode(); drawPageStrip();
      }
      function updateEditorMode() {
        const mode = document.getElementById('mode').value || 'new';
        document.getElementById('modify-fields').hidden = mode === 'new';
        document.getElementById('page-range-field').hidden = mode !== 'new';
        document.getElementById('page_range').disabled = mode !== 'new';
        document.getElementById('new-mode-hint').hidden = mode !== 'new';
        const pages = document.getElementById('modify_page').value.trim();
        document.getElementById('scope-summary').textContent = mode === 'new' ? '生成新文稿，默认 2–5 页' :
          (pages ? `仅修改第 ${pages} 页` : `将修改全部 ${previewData.length} 页`);
        document.getElementById('modify-hint').textContent = mode === 'modify_text' ?
          '文字修改保留位置、字号和图形。过长时会提示精简或重新排版。' : '选择缩略图或填写页码；留空修改全部页，原成功版本保留。';
        const historical = Boolean(viewingVersion && viewingVersion !== latestState.version_id);
        const stale = Boolean(latestState.version_id && document.getElementById('base-version-id').value && document.getElementById('base-version-id').value !== latestState.version_id);
        document.getElementById('readonly-notice').hidden = !(historical || stale);
        document.getElementById('version-notice-text').textContent = historical ? '正在查看历史版本，修改前请切回当前版本。' : '文稿已在其他操作中更新，请重新打开当前文稿再修改。';
        document.getElementById('generate-button').disabled = ['queued', 'running'].includes(latestState.generation_status) ||
          (mode !== 'new' && (!latestState.version_id || historical || stale));
        document.getElementById('generate-button').textContent = mode === 'new' ? '生成 PPT' : '保存页面修改';
      }
      function reloadCurrentDocument() {
        if (formDirty && !confirm('尚未提交的表单输入将清空，是否重新打开当前文稿？')) return;
        const conversation = document.getElementById('conversation-id').value;
        window.location.href = conversation ? '/?conversation_id=' + encodeURIComponent(conversation) : '/';
      }
      function updateChatMode() {
        const question = document.getElementById('chat-intent').value === 'question';
        document.getElementById('chat-scope').hidden = !question;
        document.getElementById('chat-mode-hint').textContent = question ? '仅回答已保存内容，不生成或修改 PPT。' : '讨论可以形成可应用的生成计划。';
        document.getElementById('chat-input').placeholder = question ? '例如：第 3 页的处理方式有什么区别？' : '说明主题、受众、重点或希望调整的内容。';
      }
      async function chooseVersion(id) {
        try {
          const requested = id && id !== latestState.version_id ? id : '';
          const data = requested ? await apiRequest(`/api/versions/${encodeURIComponent(id)}/preview`) : latestState;
          viewingVersion = requested;
          document.getElementById('version-picker').value = viewingVersion;
          document.getElementById('download-link').textContent = viewingVersion ? '下载此历史版本' : '下载当前 PPT';
          activePage = 0; previewFingerprint = ''; renderPreview(data.preview || []);
          document.getElementById('readonly-notice').hidden = !viewingVersion;
          document.getElementById('download-link').href = data.download_url || '#';
          document.getElementById('download-link').hidden = !data.download_url;
          updateEditorMode();
        } catch (error) { showWorkbenchError(error.message); }
      }
      function applyProposal(plan, version) {
        if (plan.action !== 'new' && (version !== latestState.version_id || viewingVersion)) {
          showWorkbenchError('计划来自历史版本，请切回当前版本后重新讨论修改。'); return;
        }
        document.getElementById('topic').value = plan.topic;
        document.getElementById('requirements').value = plan.requirements;
        document.getElementById('mode').value = plan.action;
        document.getElementById('page_range').value = plan.page_range || '';
        document.getElementById('modify_page').value = (plan.pages || []).join(',');
        formDirty = true; switchPanel('editor'); updateEditorMode(); drawPageStrip(); return true;
      }
      function executeProposal(plan, version) {
        if (!applyProposal(plan, version)) return;
        if (document.getElementById('generate-button').disabled) {
          showWorkbenchError('当前任务尚未完成或版本已更新，请先完成本轮任务再提交计划。'); return;
        }
        document.getElementById('generation-form').requestSubmit();
      }
      async function refreshConversations(selected) {
        const data = await apiRequest('/api/conversations');
        const picker = document.getElementById('conversation-picker');
        const fragment = document.createDocumentFragment();
        const empty = document.createElement('option'); empty.value = ''; empty.textContent = '选择已有文稿'; fragment.append(empty);
        data.conversations.forEach(item => {
          const option = document.createElement('option'); option.value = item.id; option.textContent = item.title; fragment.append(option);
        });
        picker.replaceChildren(fragment); picker.value = selected || '';
      }
      let loadingDialogue = false, pendingDialoguePayload = null;
      const visibleDialogue = new Map();
      let earlierHistory = true;
      async function loadDialogue(beforeId) {
        const conversation = document.getElementById('conversation-id').value;
        if (!conversation || loadingDialogue) return;
        loadingDialogue = true;
        try {
          const data = await apiRequest(`/api/conversations/${encodeURIComponent(conversation)}/dialogue` + (beforeId ? `?before_id=${beforeId}` : ''));
          data.messages.forEach(message => visibleDialogue.set(message.id, message));
          if (beforeId || visibleDialogue.size <= 100) earlierHistory = data.has_older;
          const messages = Array.from(visibleDialogue.values()).sort((a,b) => a.id - b.id);
          const fingerprint = JSON.stringify({...data, messages});
          if (fingerprint === dialogueFingerprint) return;
          dialogueFingerprint = fingerprint;
          const container = document.getElementById('chat-messages');
          const nearBottom = container.scrollHeight - container.scrollTop - container.clientHeight < 80;
          const fragment = document.createDocumentFragment();
          if (earlierHistory && messages.length) {
            const older = document.createElement('button'); older.type = 'button'; older.className = 'quiet'; older.textContent = '加载更早的记录';
            older.addEventListener('click', async () => { older.disabled = true; try { await loadDialogue(messages[0].id); } finally { older.disabled = false; } });
            fragment.append(older);
          }
          messages.forEach(message => {
            const turn = document.createElement('article'); turn.className = 'chat-turn ' + message.role;
            const role = document.createElement('div'); role.className = 'chat-role';
            role.textContent = message.role === 'user' ? (message.kind === 'generation' ? '你 · 生成 / 修改要求' : '你') : 'PPT 助手';
            const content = document.createElement('div'); content.className = 'chat-content'; content.textContent = message.content;
            turn.append(role, content);
            const reply = message.response;
            if (reply) {
              const scope = reply.retrieval || {};
              if (scope.pages && (scope.pages.length < scope.total_pages || scope.summary_only)) {
                const detail = document.createElement('div'); detail.className = 'hint';
                detail.textContent = `回答依据：第 ${scope.pages.join('、')} 页${scope.summary_only ? '的标题与结论' : ''}。`;
                turn.append(detail);
              }
              (reply.citations || []).forEach(citation => {
                const button = document.createElement('button'); button.className = 'citation-button'; button.type = 'button';
                button.textContent = `第 ${citation.page} 页`; button.title = citation.quote;
                button.addEventListener('click', async () => {
                  if (reply.version_id && reply.version_id !== (viewingVersion || latestState.version_id)) await chooseVersion(reply.version_id);
                  activePage = citation.page - 1; drawPageStrip(); drawActiveSlide();
                });
                turn.append(button);
              });
              if (reply.proposal) {
                const box = document.createElement('div'); box.className = 'proposal-box';
                const text = document.createElement('div'); text.textContent = `可应用的计划：${reply.proposal.topic}`;
                const apply = document.createElement('button'); apply.type = 'button'; apply.textContent = '应用到生成 / 修改表单';
                apply.addEventListener('click', () => applyProposal(reply.proposal, reply.version_id));
                const details = document.createElement('details');
                const summary = document.createElement('summary'); summary.textContent = `查看计划要求 · ${reply.proposal.page_range || '保持页数'}`;
                const requirements = document.createElement('div'); requirements.className = 'chat-content'; requirements.textContent = reply.proposal.requirements;
                details.append(summary, requirements);
                const execute = document.createElement('button'); execute.type = 'button'; execute.className = 'primary';
                execute.textContent = reply.proposal.action === 'new' ? '按此计划生成' : '按此计划修改';
                execute.addEventListener('click', () => executeProposal(reply.proposal, reply.version_id));
                box.append(text, details, apply, execute); turn.append(box);
              }
            }
            fragment.append(turn);
          });
          const waiting = data.jobs.filter(job => ['queued', 'running'].includes(job.status));
          const failed = data.jobs.filter(job => ['failed', 'unknown', 'conflict'].includes(job.status));
          failed.forEach(job => {
            const row = document.createElement('div'); row.className = 'chat-job-error';
            row.textContent = job.status === 'unknown' ? '一次对话调用的结果未知，重试前需要确认额度风险。' : (job.error === 'ContextBudgetError' ? '本次资料超出输入预算，请缩小提问页面范围。' : '一次对话未完成，原文与已有文稿保留。');
            const retry = document.createElement('button'); retry.type = 'button'; retry.textContent = '重试这次对话';
            retry.addEventListener('click', () => retryDialogue(job)); row.append(retry); fragment.append(row);
          });
          container.replaceChildren(fragment);
          if (nearBottom || data.messages.length <= 2) container.scrollTop = container.scrollHeight;
          document.getElementById('chat-status').textContent = waiting.length ? '消息已保存，助手正在处理…' : (data.has_older ? '显示最近 100 条记录，更早的历史仍保存在会话中。' : '');
          document.getElementById('chat-send').disabled = waiting.length > 0;
        } catch (error) { showWorkbenchError(error.message); } finally { loadingDialogue = false; }
      }
      async function retryDialogue(job) {
        if (job.status === 'unknown' && !confirm('上次模型调用结果未知，重试可能再次消耗额度，是否继续？')) return;
        try {
          await apiRequest(`/api/jobs/${encodeURIComponent(job.id)}/retry`, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({confirm_unknown:job.status === 'unknown'})});
          dialogueFingerprint = ''; await loadDialogue();
        } catch (error) { showWorkbenchError(error.message); }
      }
      async function sendDialogue(event) {
        event.preventDefault();
        const input = document.getElementById('chat-input'); const text = input.value.trim(); if (!text) return;
        const button = document.getElementById('chat-send'); button.disabled = true;
        const intent = document.getElementById('chat-intent').value || 'discuss';
        const version = viewingVersion || latestState.version_id || '';
        const requestId = globalThis.crypto && globalThis.crypto.randomUUID ? globalThis.crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2);
        try {
          const payload = {message:text, intent, version_id:version, conversation_id:document.getElementById('conversation-id').value,
            pages:intent === 'question' && document.getElementById('chat-scope').value === 'page' ? [activePage + 1] : []};
          const signature = JSON.stringify(payload);
          if (!pendingDialoguePayload || pendingDialoguePayload.signature !== signature) pendingDialoguePayload = {signature, request:{...payload, request_id:requestId}};
          const data = await apiRequest('/api/dialogue', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(pendingDialoguePayload.request)});
          pendingDialoguePayload = null;
          document.getElementById('conversation-id').value = data.conversation_id; input.value = '';
          await refreshConversations(data.conversation_id);
          dialogueFingerprint = ''; await loadDialogue();
        } catch (error) { showWorkbenchError(error.message); button.disabled = false; }
      }

      let refreshingState = false;
      let currentJobId = null;
      let currentJobStatus = 'idle';
      let loadedVersionId;
      async function apiRequest(url, options = {}) {
        const response = await fetch(url, { cache: 'no-store', ...options });
        const data = await response.json();
        if (!response.ok) throw new Error(data.message || '请求失败');
        return data;
      }

      async function chooseConversation(id) {
        const selected = id || document.getElementById('conversation-picker').value;
        if (!selected) return;
        if (formDirty && !confirm('尚未提交的表单输入将清空，是否切换文稿？')) return;
        try {
          await apiRequest(`/api/conversations/${encodeURIComponent(selected)}/select`, { method: 'POST' });
          window.location.href = '/?conversation_id=' + encodeURIComponent(selected);
        } catch (err) { alert(err.message); }
      }

      async function loadArchives() {
        const container = document.getElementById('archive-list');
        container.textContent = '正在读取旧归档…';
        try {
          const data = await apiRequest('/api/archives');
          const fragment = document.createDocumentFragment();
          for (const archive of data.archives) {
            const row = document.createElement('div');
            row.className = 'history-row';
            const text = document.createElement('div');
            text.textContent = `${archive.run_id} · ${archive.title || archive.message || '无标题'}`;
            row.append(text);
            if (archive.status !== 'invalid') {
              const open = document.createElement('button');
              open.type = 'button';
              open.textContent = archive.conversation_id ? '打开已导入会话' : '导入为独立会话';
              open.addEventListener('click', async () => {
                open.disabled = true;
                try {
                  if (archive.conversation_id) {
                    await chooseConversation(archive.conversation_id);
                  } else {
                    await apiRequest(`/api/archives/${encodeURIComponent(archive.run_id)}/import`, {
                      method: 'POST', headers: { 'Content-Type': 'application/json' },
                      body: JSON.stringify({ source_hash: archive.source_hash })
                    });
                    window.location.href = '/';
                  }
                } catch (err) { alert(err.message); } finally { open.disabled = false; }
              });
              row.append(open);
            }
            fragment.append(row);
          }
          if (!data.archives.length) fragment.append(document.createTextNode('没有可扫描的旧会话归档。'));
          container.replaceChildren(fragment);
        } catch (err) { container.textContent = err.message; }
      }

      async function loadVersions(conversationId) {
        const container = document.getElementById('version-list');
        if (!conversationId) { container.replaceChildren(); return; }
        const data = await apiRequest(`/api/conversations/${encodeURIComponent(conversationId)}/versions`);
        const fragment = document.createDocumentFragment();
        const picker = document.getElementById('version-picker');
        const options = document.createDocumentFragment();
        const current = document.createElement('option'); current.value = ''; current.textContent = '当前版本'; options.append(current);
        const labels = { pptx: 'PPT', outline: '内容大纲', blueprint: '排版方案', 'legacy-state': '归档记录', image: '配图', pdf: 'PDF', preview: '页面截图', 'visual-review': '审核报告', 'candidate-pptx': '设计候选' };
        data.versions.forEach(version => {
          const row = document.createElement('div');
          row.className = 'history-row';
          const title = document.createElement('div');
          title.textContent = `${version.created_at} · ${version.file_name}`;
          row.append(title);
          const option = document.createElement('option'); option.value = version.id;
          option.textContent = new Date(version.created_at).toLocaleString('zh-CN', {month:'2-digit', day:'2-digit', hour:'2-digit', minute:'2-digit'});
          options.append(option);
          version.artifacts.forEach(artifact => {
            if (!['pptx', 'outline', 'blueprint'].includes(artifact.role)) return;
            if (!artifact.download_url.startsWith('/api/artifacts/')) return;
            const link = document.createElement('a');
            link.href = artifact.download_url;
            link.className = 'file-button'; link.download = '';
            link.textContent = labels[artifact.role] || artifact.role;
            row.append(link);
          });
          fragment.append(row);
        });
        container.replaceChildren(fragment);
        picker.replaceChildren(options); picker.value = viewingVersion;
      }

      async function retryCurrentJob() {
        if (!currentJobId) return;
        const unknown = currentJobStatus === 'unknown';
        if (unknown && !confirm('上次外部调用的结果未知。重试可能再次消耗模型或生图额度，是否继续？')) return;
        const button = document.getElementById('retry-button');
        button.disabled = true;
        try {
          await apiRequest(`/api/jobs/${encodeURIComponent(currentJobId)}/retry`, {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ confirm_unknown: unknown })
          });
          await refreshState();
        } catch (err) { alert(err.message); } finally { button.disabled = false; }
      }
      async function refreshState() {
        if (refreshingState) return;
        refreshingState = true;
        try {
          const conversation = document.getElementById('conversation-id').value;
          const response = await fetch('/state' + (conversation ? `?conversation_id=${encodeURIComponent(conversation)}` : ''), { cache: 'no-store' });
          if (!response.ok) return;
          const data = await response.json();
          const statusEl = document.getElementById('status-text');
          const lastUpdatedEl = document.getElementById('last-updated');
          const logEl = document.getElementById('progress-log');
          const resultEl = document.getElementById('generation-result');
          latestState = data;
          currentJobId = data.job_id;
          currentJobStatus = data.generation_status;
          document.getElementById('generate-button').disabled = ['queued', 'running'].includes(currentJobStatus);
          document.getElementById('retry-button').hidden = !currentJobId || !['failed', 'unknown', 'conflict'].includes(currentJobStatus);
          const submittedJobId = document.getElementById('generation-form').dataset.jobId;
          if (submittedJobId && submittedJobId === data.job_id && data.generation_status === 'succeeded') {
            document.getElementById('base-version-id').value = data.version_id || '';
          }
          const download = document.getElementById('download-link');
          if (!viewingVersion) { download.hidden = !data.download_url; download.href = data.download_url || '#'; }
          if (loadedVersionId !== data.version_id) {
            await loadVersions(data.conversation_id);
            loadedVersionId = data.version_id;
          }

          if (resultEl) {
            resultEl.dataset.status = data.generation_status || 'idle';
            if (data.artifact_integrity === 'invalid') {
              resultEl.textContent = '已保存文件缺失或被修改，请从备份恢复。';
              resultEl.style.color = '#b91c1c';
            } else if (data.generation_status === 'unknown') {
              resultEl.textContent = '外部调用结果或缓存资产需要确认，本轮未提交新版本。';
              resultEl.style.color = '#92400e';
            } else if (data.generation_status === 'conflict') {
              resultEl.textContent = '基准版本已改变，请刷新后重新提交修改。';
              resultEl.style.color = '#92400e';
            } else if (data.generation_status === 'failed') {
              resultEl.textContent = '本轮生成失败，未创建新的成功快照。' +
                (data.last_file ? '保留上一份成功文件供预览和下载。' : '');
              resultEl.style.color = '#b91c1c';
              if (data.layout_recovery_request) {
                const recovery = data.layout_recovery_request;
                resultEl.textContent = recovery.capacity_conflict ?
                  `第 ${recovery.page} 页在当前稳定布局候选中空间不足。请增加页数或明确可精简的内容；原文已保存。` :
                  `第 ${recovery.page} 页的图示或内容约束仍不满足要求。请检查原始内容；原文已保存。`;
              }
            } else if (data.generation_status === 'succeeded' && data.result_kind === 'offline_demo') {
              resultEl.textContent = '已导出离线演示版，未执行智能体生成与审核。';
              resultEl.style.color = '#92400e';
            } else if (data.generation_status === 'succeeded') {
              const needsReview = data.content_review_status === 'needs_review';
              const visual = data.visual_review || {};
              resultEl.textContent = needsReview ? 'PPT 已导出，内容需要人工复核。' :
                (visual.status === 'passed' ? 'PPT 已导出，视觉模型未发现问题；仍需人工验收。' : 'PPT 已导出，页面未审核或仍需复核。');
              resultEl.style.color = needsReview || visual.status !== 'passed' ? '#92400e' : '#166534';
              const stablePages = (data.layout_fallbacks || []).map(item => item.page).filter(Number.isInteger);
              if (stablePages.length) {
                resultEl.textContent = `PPT 已导出；第 ${stablePages.join('、')} 页自由排版未通过，已使用保留完整内容的稳定布局。仍需人工复核。`;
                resultEl.style.color = '#92400e';
              }
            } else {
              resultEl.textContent = data.generation_status === 'running' ? '正在生成，完成后更新文件。' :
                (data.generation_status === 'queued' ? '任务已保存，等待开始。' : '尚未开始生成。');
              resultEl.style.color = '#334155';
            }
          }

          const statusText = data.status || '等待开始...';
          if (statusEl) {
            statusEl.textContent = statusText;
          }
          if (lastUpdatedEl) {
            lastUpdatedEl.textContent = data.last_updated || '暂无';
          }
          if (logEl) {
            const fragment = document.createDocumentFragment();
            (data.progress_log || []).slice(-60).forEach(item => {
              const entry = document.createElement('li');
              entry.textContent = item;
              fragment.append(entry);
            });
            logEl.replaceChildren(fragment);
          }
          document.getElementById('workspace-title').textContent = data.conversation_title || data.topic || '新文稿';
          document.getElementById('document-subtitle').textContent = data.version_id ? '已保存文稿 · 勾选页面可指定修改范围' : '先完善需求，再生成第一份文稿';
          if (!viewingVersion) renderPreview(data.preview || []);
          updateEditorMode();
          await loadDialogue();
          await refreshDesignCandidates();
        } catch (err) {
          console.error(err);
        } finally {
          refreshingState = false;
        }
      }

      /* 新增：点击“新建对话”按钮触发的 JS 逻辑 */
      async function startNewChat() {
        if (formDirty && !confirm('尚未提交的表单输入将清空，是否新建文稿？')) return;
        try {
          const response = await fetch('/api/reset', { method: 'POST' });
          const data = await response.json();
          if (data.status === 'success') {
            // 重置成功，直接刷新页面回到最初的纯净状态
            window.location.href = '/';
          } else {
            alert('重置失败: ' + (data.message || '未知错误'));
          }
        } catch (err) {
          console.error(err);
          alert('请求失败，请检查后端服务是否正常运行。');
        }
      }

      /* 注销 API Key：清除本机密钥并返回设置页 */
      async function logoutKey() {
        if (!confirm('确定要注销当前 API Key 并返回设置页吗？')) {
          return;
        }
        try {
          const response = await fetch('/api/logout', { method: 'POST' });
          const data = await response.json();
          if (data.status === 'success') {
            window.location.href = '/setup';
          } else {
            alert('注销失败: ' + (data.message || '未知错误'));
          }
        } catch (err) {
          console.error(err);
          alert('请求失败，请检查后端服务是否正常运行。');
        }
      }

      window.addEventListener('DOMContentLoaded', () => {
        updateEditorMode(); updateChatMode(); refreshState();
        document.getElementById('chat-form').addEventListener('submit', sendDialogue);
        document.getElementById('generation-form').addEventListener('input', () => { formDirty = true; });
        document.getElementById('modify_page').addEventListener('input', () => { updateEditorMode(); drawPageStrip(); });
        document.getElementById('generation-form').addEventListener('submit', () => { formDirty = false; document.getElementById('generate-button').disabled = true; });
        document.addEventListener('keydown', event => {
          if (document.getElementById('preview-dialog').open) return;
          if (document.getElementById('design-dialog').open) {
            if (event.key === 'ArrowLeft') { event.preventDefault(); stepDesignPage(-1); }
            if (event.key === 'ArrowRight') { event.preventDefault(); stepDesignPage(1); }
            return;
          }
          if (['INPUT', 'TEXTAREA', 'SELECT'].includes(event.target.tagName)) {
            if (event.ctrlKey && event.key === 'Enter' && event.target.id === 'chat-input') document.getElementById('chat-form').requestSubmit();
            return;
          }
          if (event.key === 'ArrowLeft') changePage(-1);
          if (event.key === 'ArrowRight') changePage(1);
        });
        apiRequest('/api/has_key').then(data => {
          const label = document.getElementById('model-connection');
          label.textContent = data.has_key ? '模型服务已配置' : '未配置模型 · 可离线演示';
          label.className = 'connection' + (data.has_key ? ' connected' : '');
        }).catch(() => { document.getElementById('model-connection').textContent = '无法读取模型配置'; });
        setInterval(refreshState, 2000);
      });
