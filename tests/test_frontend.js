const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

class Element {
  constructor(tag) {
    this.tag = tag;
    this.children = [];
    this.dataset = {};
    this.style = {};
    this.value = '';
    this.classList = { toggle() {} };
    this.attributes = {};
    this.listeners = {};
    this.scrollHeight = this.clientHeight = this.scrollTop = 0;
  }
  set textContent(value) { this.text = String(value); this.children = []; }
  get textContent() { return this.text || this.children.map(child => child.textContent).join(''); }
  set innerHTML(value) { throw new Error('动态内容不能写入 innerHTML'); }
  append(...nodes) {
    for (const node of nodes) {
      if (node.tag === 'fragment') this.children.push(...node.children);
      else this.children.push(node);
    }
  }
  replaceChildren(...nodes) { this.children = []; this.append(...nodes); }
  addEventListener(name, handler) { this.listeners[name] = handler; }
  setAttribute(name, value) { this.attributes[name] = value; }
  requestSubmit() { this.submittedCount = (this.submittedCount || 0) + 1; }
  showModal() { this.open = true; }
  close() { this.open = false; }
}

const elements = new Map();
const get = id => {
  if (!elements.has(id)) elements.set(id, new Element('div'));
  return elements.get(id);
};
get('generation-form').dataset.jobId = 'job-1';
let state = { generation_status: 'queued', job_id: 'job-1', conversation_id: 'conversation-1', version_id: 'version-1',
  progress_log: [], last_file: '', download_url: '', artifact_integrity: 'none', preview: [] };
let stateRequests = 0;
let versionRows = [];
let retryRequests = 0;
let dialoguePosts = [], failDialoguePost = true;
let designData = {jobs:[]}, designPosts = [], adoptPosts = [], failDesignPost = true;
const context = vm.createContext({
  document: { getElementById: get, createElement: tag => new Element(tag),
    createDocumentFragment: () => new Element('fragment'),
    createTextNode: value => { const node = new Element('text'); node.textContent = value; return node; } },
  window: { addEventListener() {}, location: {} },
  console, setInterval() {}, alert() {}, confirm: () => false,
  fetch: async (url, options = {}) => {
    let data;
    if (url.startsWith('/state')) { stateRequests++; data = structuredClone(state); }
    else if (url === '/api/conversations') data = {conversations:[{id:'conversation-1',title:'文稿'}]};
    else if (url.endsWith('/design')) data = structuredClone(designData);
    else if (url === '/api/design/candidates') {
      designPosts.push(JSON.parse(options.body));
      if (failDesignPost) {failDesignPost=false;throw new Error('test design network unknown');}
      data={job_id:'design-job',conversation_id:'conversation-1',status:'queued'};
    }
    else if (url === '/api/design/adopt') {
      adoptPosts.push(JSON.parse(options.body));
      state.job_id='design-adopt-job';state.version_id='version-2';state.generation_status='succeeded';
      data={job_id:'design-adopt-job',status:'queued'};
    }
    else if (url.endsWith('/preview')) data={preview:[{title:'原版',image_url:'/api/artifacts/a-1/download?inline=1'}]};
    else if (url.endsWith('/versions')) data = { versions: versionRows };
    else if (url.endsWith('/dialogue') && url !== '/api/dialogue') data = { messages: [], jobs: [], has_older: false };
    else if (url === '/api/dialogue') {
      const body = JSON.parse(options.body); dialoguePosts.push(body);
      if (failDialoguePost) { failDialoguePost = false; throw new Error('test network outcome unknown'); }
      data = {conversation_id:'conversation-1',job_id:'dialogue-job'};
    }
    else if (url.endsWith('/retry')) { retryRequests++; state.generation_status = 'queued'; data = { status: 'queued' }; }
    else throw new Error('未预期的测试请求');
    return { ok: true, json: async () => data };
  },
});
const script = fs.readFileSync(path.join(__dirname, '..', 'static', 'workbench.js'), 'utf8');
vm.runInContext(script, context);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'static', 'design.js'), 'utf8'), context);

(async () => {
  get('show-issues').checked = true;
  const malicious = '<img src=x onerror="throw new Error(1)">';
  context.renderPreview([{ title: malicious, text: malicious }]);
  assert.equal(get('ppt-preview').children[0].children[0].textContent, '1. ' + malicious + ' · 已保存内容摘要（暂无成品截图）');
  assert.equal(get('ppt-preview').children[0].children[1].textContent, malicious);
  context.renderPreview([{ title: '预览', text: '', image_url: '/api/artifacts/a-1/download?inline=1', review_status: 'needs_review',
    issues: [{ bbox: [0.1, 0.2, 0.3, 0.4], message: malicious }] }]);
  const frame = get('ppt-preview').children[0].children[2];
  assert.equal(frame.children[0].src, '/api/artifacts/a-1/download?inline=1');
  assert.equal(frame.children[1].title, malicious);
  assert.equal(frame.children[1].style.left, '10%');
  context.renderPreview([{ title: '无效图片地址', image_url: 'javascript:alert(1)' }]);
  assert.equal(get('ppt-preview').children[0].children.some(child => child.className === 'preview-image'), false);
  state.progress_log = [malicious];
  await context.refreshState();
  assert.equal(get('progress-log').children[0].textContent, malicious);
  assert.equal(get('generate-button').disabled, true);
  assert.equal(get('generation-result').textContent, '任务已保存，等待开始。');

  state.generation_status = 'unknown';
  await context.refreshState();
  assert.equal(get('retry-button').hidden, false);
  await context.retryCurrentJob();
  assert.equal(retryRequests, 0);

  state.generation_status = 'failed';
  state.last_file = 'previous.pptx';
  await context.refreshState();
  assert.match(get('generation-result').textContent, /保留上一份成功文件/);
  await context.retryCurrentJob();
  assert.equal(retryRequests, 1);

  state.generation_status = 'succeeded';
  state.result_kind = 'offline_demo';
  state.download_url = '/api/artifacts/a/download';
  await context.refreshState();
  assert.match(get('generation-result').textContent, /离线演示版/);
  assert.equal(get('download-link').hidden, false);
  assert.equal(get('base-version-id').value, 'version-1');
  state.job_id = 'other-client-job';
  state.version_id = 'version-2';
  await context.refreshState();
  assert.equal(get('base-version-id').value, 'version-1');

  const before = stateRequests;
  await Promise.all([context.refreshState(), context.refreshState()]);
  assert.equal(stateRequests - before, 1);
  context.renderPreview([{title:'第一页',text:'A'},{title:'第二页',text:'B'},{title:'第三页',text:'C'}]);
  context.changePage(1);
  assert.match(get('ppt-preview').children[0].children[0].textContent, /2\. 第二页/);
  context.useCurrentPage();
  assert.equal(get('modify_page').value, '2');
  assert.equal(get('mode').value, 'modify_text');
  assert.deepEqual(Array.from(context.selectedPageNumbers('1 - 3')), [1,2,3]);
  assert.equal(context.selectedPageNumbers('0'), null);
  context.applyProposal({action:'modify_text',topic:'不应应用',requirements:'不应应用',pages:[1]}, 'old-version');
  assert.match(get('chat-status').textContent, /历史版本/);
  assert.notEqual(get('topic').value, '不应应用');
  context.applyProposal({action:'new',topic:'新主题',requirements:'用户明确不生图',page_range:'4',pages:[]}, null);
  assert.equal(get('topic').value, '新主题');
  assert.equal(get('requirements').value, '用户明确不生图');
  assert.equal(get('page_range').value, '4');
  context.executeProposal({action:'new',topic:'新文稿',requirements:'四页不生图',page_range:'4',pages:[]}, null);
  assert.equal(get('generation-form').submittedCount, 1);
  context.renderPreview([]);
  assert.match(get('ppt-preview').textContent, /暂无可预览/);
  get('chat-input').value = '不生成文件，只提问';
  get('chat-intent').value = 'question';
  const event = {preventDefault(){}};
  await context.sendDialogue(event);
  await context.sendDialogue(event);
  assert.equal(dialoguePosts.length, 2);
  assert.equal(dialoguePosts[0].request_id, dialoguePosts[1].request_id);
  assert.equal(dialoguePosts[0].intent, 'question');
  assert.equal(get('chat-input').value, '');
  state.generation_status='succeeded';state.version_id='version-1';state.job_id='job-1';
  state.result_kind='agent_generated';state.layout_fallbacks=[{page:3}];
  await context.refreshState();
  assert.match(get('generation-result').textContent,/第 3 页.*稳定布局/);
  state.generation_status='failed';state.layout_recovery_request={page:3,capacity_conflict:true};
  await context.refreshState();
  assert.match(get('generation-result').textContent,/请增加页数.*原文已保存/);
  state.generation_status='succeeded';state.layout_recovery_request=null;
  state.layout_fallbacks=[];
  state.preview=[{title:'咖啡'}]; get('base-version-id').value='version-1';get('conversation-id').value='conversation-1';
  await context.refreshState();get('design-pages').value='1';get('design-request').value='使用本地图标';
  await context.generateDesignCandidates();await context.generateDesignCandidates();
  assert.equal(designPosts.length,2);assert.equal(designPosts[0].request_id,designPosts[1].request_id);
  designData={jobs:[{id:'design-job',status:'succeeded',pages:[1],base_version_id:'version-1',candidates:[{
    id:'clear',name:malicious,description:malicious,status:'ready',integrity:true,icon_count:2,
    assets:[{id:'b',role:'blueprint',sha256:'fixed-hash'},{id:'p',role:'preview',page:1,image_url:'/api/design/jobs/design-job/assets/p?inline=1'}]
  }]}]};
  await context.refreshDesignCandidates();await context.openDesignComparison('design-job');
  assert.equal(get('design-dialog').open,true);
  assert.equal(get('design-compare-grid').children[1].children[0].textContent,malicious);
  assert.equal(get('design-compare-grid').children[1].children[2].src,'/api/design/jobs/design-job/assets/p?inline=1');
  designData.jobs[0].created_at='changed';await context.refreshDesignCandidates();
  assert.equal(get('design-compare-grid').children[0].children[2].src,'/api/artifacts/a-1/download?inline=1');
  await context.adoptDesign(designData.jobs[0],designData.jobs[0].candidates[0]);
  assert.equal(adoptPosts[0].candidate_hash,'fixed-hash');assert.equal(get('design-dialog').open,false);
  assert.equal(get('base-version-id').value,'version-2');
  assert.equal(context.designImage('javascript:alert(1)'),false);
  assert.equal(context.designImage('https://evil.test/api/design/jobs/a/assets/b?inline=1'),false);
  assert.equal(get('design-compare-grid').children[1].children.at(-1).disabled,true);
  versionRows=[{id:'version-1',created_at:'2026-10-15T01:00:00Z',file_name:'RAG.pptx',artifacts:
    ['pptx','outline','blueprint','preview','pdf'].map((role,i)=>({role,download_url:`/api/artifacts/asset-${i}/download`}))}];
  await context.loadVersions('conversation-1');
  const files=get('version-list').children[0].children.slice(1);
  assert.equal(files.length,3);assert.ok(files.every(link=>link.className==='file-button' && link.download===''));
  designData.jobs=[{id:'empty',status:'succeeded',pages:[1],base_version_id:'version-1',candidates:[]}];
  await context.refreshDesignCandidates();assert.match(get('design-list').textContent,/未产出可用候选/);
  console.log('前端检查通过：文件按钮及截图过滤、空候选提示、文本注入、幂等提交与轮询。');
})().catch(error => { console.error(error); process.exitCode = 1; });
