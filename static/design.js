function designRequestId() { return globalThis.crypto?.randomUUID ? globalThis.crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2); }
let designJobs = [];
let designComparingJob = null;
let designComparingPage = 0;
let designPending = null;
let designAdopting = false;
let designListHash = '';

function designImage(url) {
  return /^\/api\/design\/jobs\/[A-Za-z0-9_-]+\/assets\/[A-Za-z0-9_-]+\?inline=1$/.test(url || '');
}
function designCurrentVersion() { return latestState.version_id || document.getElementById('base-version-id').value; }
function designButtonAllowed() {
  return Boolean(designCurrentVersion() && (!viewingVersion || viewingVersion === latestState.version_id) &&
    document.getElementById('base-version-id').value === designCurrentVersion() && !['queued','running'].includes(currentJobStatus));
}
function designPageSelection() {
  const value = document.getElementById('design-pages').value;
  return value.trim() ? selectedPageNumbers(value) : [activePage + 1];
}
function useDesignSelection() {
  document.getElementById('design-pages').value = document.getElementById('modify_page').value || String(activePage + 1);
}
async function generateDesignCandidates() {
  const pages = designPageSelection();
  if (!designButtonAllowed()) { showWorkbenchError('请重新打开当前文稿，再生成候选。'); return; }
  if (!pages || !pages.length || pages.length > 3) { showWorkbenchError('请选择 1–3 个页面，例如 2 或 1,3。'); return; }
  const data = {conversation_id: document.getElementById('conversation-id').value, version_id: designCurrentVersion(),
    pages, design_request: document.getElementById('design-request').value};
  const signature = JSON.stringify(data);
  if (!designPending || designPending.signature !== signature) designPending = {signature, request_id: designRequestId()};
  document.getElementById('design-generate').disabled = true;
  try {
    await apiRequest('/api/design/candidates', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({...data,request_id:designPending.request_id})});
    designPending = null; await refreshDesignCandidates();
  } catch (error) { showWorkbenchError(error.message); }
  finally { updateDesignControls(); }
}
function updateDesignControls() {
  const active = designJobs.some(job => ['queued','running'].includes(job.status));
  document.getElementById('design-generate').disabled = active || !designButtonAllowed();
  document.getElementById('design-scope').textContent = previewData.length ? `默认当前第 ${activePage + 1} 页；也可选择最多 3 页。` : '先生成或打开已保存文稿。';
}
async function refreshDesignCandidates() {
  const conversation = document.getElementById('conversation-id').value;
  if (!conversation) { designJobs = []; updateDesignControls(); return; }
  try {
    const data = await apiRequest(`/api/conversations/${encodeURIComponent(conversation)}/design`);
    designJobs = data.jobs || [];
    const signature = JSON.stringify([designJobs,designCurrentVersion()]);
    if (signature !== designListHash) {
      designListHash = signature; renderDesignList();
      if (designComparingJob) {
        const refreshed = designJobs.find(job => job.id === designComparingJob.id);
        if (refreshed) designComparingJob = {...refreshed, baseline:designComparingJob.baseline};
        drawDesignComparison();
      }
    }
    updateDesignControls();
  } catch (error) { document.getElementById('design-status').textContent = error.message; }
}
function renderDesignList() {
  const list = document.getElementById('design-list'); const fragment = document.createDocumentFragment();
  designJobs.forEach(job => {
    const row = document.createElement('article'); row.className = 'design-batch';
    const title = document.createElement('strong'); title.textContent = `第 ${job.pages.join('、')} 页 · 两个设计方向`;
    const text = document.createElement('p');
    const ready = job.candidates.filter(item => item.status === 'ready' && item.integrity).length;
    const labels = {queued:'等待开始',running:'正在生成',succeeded:`${ready} 个候选可比较`,unknown:'调用结果未知',failed:'本轮失败',conflict:'版本已改变'};
    text.textContent = (job.status==='succeeded' && !ready ? '未产出可用候选，请查看原因后调整要求或重试' : (labels[job.status] || job.status)) + (job.base_version_id !== designCurrentVersion() ? ' · 来自旧版本，可查看但不能采用' : '');
    row.append(title,text);
    if (job.candidates.length) {
      const compare = document.createElement('button'); compare.type = 'button'; compare.textContent = '比较候选';
      compare.addEventListener('click',() => openDesignComparison(job.id)); row.append(compare);
    }
    if (['failed','unknown','conflict'].includes(job.status)) {
      const retry = document.createElement('button'); retry.type = 'button'; retry.className = 'quiet'; retry.textContent = '重试候选任务';
      retry.disabled = job.status === 'conflict'; retry.addEventListener('click',async () => {
        const unknown = job.status === 'unknown';
        if (unknown && !confirm('上次调用结果未知，重试可能重复消耗额度，是否继续？')) return;
        try { await apiRequest(`/api/jobs/${encodeURIComponent(job.id)}/retry`,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({confirm_unknown:unknown})}); await refreshDesignCandidates(); }
        catch(error) {showWorkbenchError(error.message);}
      }); row.append(retry);
    }
    fragment.append(row);
  });
  if (!designJobs.length) { const p=document.createElement('p');p.className='hint';p.textContent='暂无候选。生成后可与原版并排查看，采用时保存为新版本。';fragment.append(p); }
  list.replaceChildren(fragment);
  const pending = designJobs.find(job => ['queued','running'].includes(job.status));
  document.getElementById('design-status').textContent = pending ? (pending.status === 'queued' ? '候选任务已保存，等待 worker。' : `正在排版和渲染；已计 ${pending.calls_used} 次排版请求。`) : '';
}
async function openDesignComparison(jobId) {
  const job = designJobs.find(item => item.id === jobId); if (!job) return;
  try {
    const base = await apiRequest(`/api/versions/${encodeURIComponent(job.base_version_id)}/preview`);
    designComparingJob = {...job, baseline:base.preview || []}; designComparingPage = 0;
    drawDesignComparison(); document.getElementById('design-dialog-alert').hidden=true; document.getElementById('design-dialog').showModal();
  } catch(error) {showWorkbenchError(error.message);}
}
function expandDesignImage(url,title) {
  if (!artifactImage(url) && !designImage(url)) return;
  document.getElementById('expanded-image').src=url;
  document.getElementById('expanded-title').textContent=title;
  document.getElementById('preview-dialog').showModal();
}
function stepDesignPage(offset) {
  if (!designComparingJob) return;
  designComparingPage = Math.max(0,Math.min(designComparingPage+offset,designComparingJob.pages.length-1)); drawDesignComparison();
}
function drawDesignComparison() {
  const job = designComparingJob; if (!job) return;
  const page = job.pages[designComparingPage];
  document.getElementById('design-compare-page').textContent = `第 ${page} 页 · 固定内容，比较设计`;
  document.getElementById('design-prev').disabled = designComparingPage <= 0;
  document.getElementById('design-next').disabled = designComparingPage >= job.pages.length-1;
  const grid = document.getElementById('design-compare-grid'); const fragment = document.createDocumentFragment();
  const base = job.baseline || [];
  const entries = [{id:'baseline',name:'原版',description:'候选所依据的已保存版本',status:'baseline',image:((job.baseline_assets || []).find(asset=>asset.page===page) || {}).image_url || (base[page-1] || {}).image_url}, ...job.candidates];
  entries.forEach(item => {
    const card=document.createElement('article');card.className='design-card';
    const heading=document.createElement('h3');heading.textContent=item.name;
    const description=document.createElement('p');description.textContent=item.description;card.append(heading,description);
    const url=item.image || (item.assets || []).find(asset=>asset.role==='preview' && asset.page===page)?.image_url;
    if (artifactImage(url) || designImage(url)) {const image=document.createElement('img');image.src=url;image.alt=`${item.name}第${page}页成品截图`;card.append(image);}
    else {const placeholder=document.createElement('div');placeholder.className='design-image-empty';placeholder.textContent=item.status==='baseline'?'原版没有已保存截图':(item.reason || '候选尚未完成渲染');card.append(placeholder);}
    if (artifactImage(url) || designImage(url)) {const zoom=document.createElement('button');zoom.type='button';zoom.textContent='放大此页';zoom.addEventListener('click',()=>expandDesignImage(url,`${item.name} · 第 ${page} 页`));card.append(zoom);}
    if (item.id !== 'baseline') {
      const status=document.createElement('p');status.className='hint';status.textContent=item.status==='ready' ? `规则检查通过 · 使用 ${item.icon_count || 0} 个本地图标 · 视觉待复核` : (item.reason || '尚未完成');card.append(status);
      const ppt=(item.assets || []).find(asset=>asset.role==='pptx');
      if (ppt) {const link=document.createElement('a');link.href=ppt.download_url;link.download='';link.className='file-button';link.textContent='下载完整候选 PPT';card.append(link);}
      const choose=document.createElement('button');choose.type='button';choose.className='primary';choose.textContent='采用此候选，保存新版本';
      choose.disabled=designAdopting || item.status!=='ready' || !item.integrity || job.base_version_id!==designCurrentVersion() || !designButtonAllowed();
      choose.addEventListener('click',()=>adoptDesign(job,item));card.append(choose);
    }
    fragment.append(card);
  }); grid.replaceChildren(fragment);
}
async function adoptDesign(job,candidate) {
  if (!designButtonAllowed() || job.base_version_id!==designCurrentVersion()) {showWorkbenchError('版本已改变，请重新生成候选。');return;}
  const blue=candidate.assets.find(asset=>asset.role==='blueprint'); if (!blue) return;
  const data={conversation_id:document.getElementById('conversation-id').value,version_id:job.base_version_id,
    source_job_id:job.id,candidate_id:candidate.id,candidate_hash:blue.sha256};
  const signature=JSON.stringify(data);
  if (!designPending || designPending.signature!==signature) designPending={signature,request_id:designRequestId()};
  designAdopting=true;drawDesignComparison();
  try {
    const result=await apiRequest('/api/design/adopt',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({...data,request_id:designPending.request_id})});
    designPending=null;
    document.getElementById('generation-form').dataset.jobId=result.job_id;
    document.getElementById('design-dialog').close();
    await refreshState();
  } catch(error) {showWorkbenchError(error.message);const message=document.getElementById('design-dialog-alert');message.textContent=error.message;message.hidden=false;}
  finally {designAdopting=false;drawDesignComparison();}
}
