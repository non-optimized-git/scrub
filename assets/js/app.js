// ══ STATE ══════════════════════════════════════════
let workingData = [];
let headers = [];
let fileName = '';
let ops = [];
let opCounter = 0;
const msI = {};
let undoStack = [];

const OP_NAMES = {
  dedup:   '删除相邻重复行',
  replace: '文字替换',
  delrow:  '删除含关键词的整行',
  delcol:  '删除列',
  addcol:  '在某列后添加空白列',
};
const TYPE_ORDER = ['dedup','replace','delrow','delcol','addcol'];
const OP_ICONS = {
  dedup: '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="4" y="6" width="11" height="9" rx="1.5"></rect><rect x="9" y="9" width="11" height="9" rx="1.5"></rect></svg>',
  replace: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M7 7h10"></path><path d="M13 4l4 3-4 3"></path><path d="M17 17H7"></path><path d="M11 14l-4 3 4 3"></path></svg>',
  delrow: '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="4" y="5" width="16" height="14" rx="1.5"></rect><path d="M4 10h16"></path><path d="M10 14h4"></path></svg>',
  delcol: '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="4" y="5" width="16" height="14" rx="1.5"></rect><path d="M10 5v14"></path><path d="M14 12h4"></path></svg>',
  addcol: '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="4" y="5" width="16" height="14" rx="1.5"></rect><path d="M10 5v14"></path><path d="M16 12h4"></path><path d="M18 10v4"></path></svg>',
};

// ══ UPLOAD ═════════════════════════════════════════
const fileInput = document.getElementById('file-input');

function initDropZone() {
  const dz = document.getElementById('drop-zone');
  if(!dz) return;
  dz.addEventListener('click', ()=>fileInput.click());
  dz.addEventListener('dragover', e=>{e.preventDefault();dz.classList.add('drag');});
  dz.addEventListener('dragleave', ()=>dz.classList.remove('drag'));
  dz.addEventListener('drop', e=>{e.preventDefault();dz.classList.remove('drag');if(e.dataTransfer.files[0])handleFile(e.dataTransfer.files[0]);});
}
initDropZone();
fileInput.addEventListener('change', e=>{if(e.target.files[0])handleFile(e.target.files[0]);e.target.value='';});

function handleFile(file) {
  fileName = file.name.replace(/\.(xlsx|xls|csv)$/i,'');
  const reader = new FileReader();
  reader.onload = ev => {
    const wb = XLSX.read(ev.target.result, {type:'array'});
    const ws = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(ws, {header:1,defval:''});
    workingData = data.map(r=>[...r]);
    refreshHeaders();
    showFileLoaded(file);
    renderStats();
    document.getElementById('stats-wrap').style.display = 'block';
    document.getElementById('add-op-wrap').style.display = 'block';
    document.getElementById('ops-card').style.display = 'block';
    document.getElementById('btn-export').disabled = false;
    ops.forEach(op=>rebuildSelectors(op.id));
    renderTypeGroups();
    toast('文件导入成功');
  };
  reader.readAsArrayBuffer(file);
}

function showFileLoaded(file) {
  const area = document.getElementById('upload-area');
  const rows = workingData.length - 1;
  area.innerHTML = `
    <div class="file-loaded" id="file-loaded-bar"
      ondragover="event.preventDefault();this.classList.add('drag')"
      ondragleave="this.classList.remove('drag')"
      ondrop="event.preventDefault();this.classList.remove('drag');if(event.dataTransfer.files[0])handleFile(event.dataTransfer.files[0])">
      <span style="color:var(--muted);font-size:13px">📄</span>
      <span class="fname">${esc(file.name)}</span>
      <span class="fmeta">${rows} 行 · ${headers.length} 列</span>
      <button class="file-del" onclick="clearFile()" title="移除文件">✕</button>
    </div>`;
}

function clearFile() {
  workingData=[]; headers=[]; fileName='';
  // restore drop zone
  document.getElementById('upload-area').innerHTML = `
    <div class="drop-zone" id="drop-zone">
      <div class="drop-icon">⊞</div>
      <p><strong>点击选择</strong>或拖拽文件到此处</p>
      <p style="margin-top:3px">支持 .xlsx · .xls · .csv</p>
    </div>`;
  initDropZone();
  document.getElementById('stats-wrap').style.display='none';
  document.getElementById('add-op-wrap').style.display='none';
  document.getElementById('ops-card').style.display='none';
  document.getElementById('btn-export').disabled=true;
  renderStats();
}

function refreshHeaders(){
  headers=(workingData[0]||[]).map((h,i)=>h!==''?String(h):`列${i+1}`);
}
function renderStats(){
  const rows=workingData.length-1;
  const el=document.getElementById('stats-row'); if(!el) return;
  el.innerHTML=`
    <div class="stat"><div class="val">${rows<0?0:rows}</div><div class="lbl">数据行</div></div>
    <div class="stat"><div class="val">${headers.length}</div><div class="lbl">列数</div></div>
    <div class="stat"><div class="val">${rows>0?(rows*headers.length).toLocaleString():0}</div><div class="lbl">单元格</div></div>`;
}

// ══ MULTI-SELECT ═══════════════════════════════════
function msCreate(msId,options,placeholder,containerId,onChange){
  if(!msI[msId]) msI[msId]={selected:new Set()};
  Object.assign(msI[msId],{options,placeholder:placeholder||'请选择列…',containerId,onChange:onChange||null});
  const valid=new Set(options.map(o=>o.value));
  for(const v of [...msI[msId].selected]) if(!valid.has(v)) msI[msId].selected.delete(v);
}
function msRender(msId){
  const inst=msI[msId]; if(!inst) return;
  const el=document.getElementById(inst.containerId); if(!el) return;
  const sel=[...inst.selected];
  const trigTxt=sel.length===0
    ?`<span style="color:var(--light)">${inst.placeholder}</span>`
    :`<span style="color:var(--accent)">${sel.map(v=>msLbl(msId,v)).join('、')}</span>`;
  el.innerHTML=`
    <div class="ms-wrap" id="msw-${msId}">
      <div class="ms-trigger" onclick="msTrigger('${msId}')">${trigTxt}</div>
      <span class="ms-arrow">▼</span>
      <div class="ms-drop">
        <input class="ms-search" placeholder="搜索列名…" oninput="msFilter('${msId}',this.value)" onclick="event.stopPropagation()">
        <div id="mso-${msId}">${msRenderOpts(msId,inst.options)}</div>
      </div>
    </div>`;
  // tags go into op-info-row (extract opId from msId: dc-{id})
  const opId=msId.startsWith('dc-')?msId.slice(3):null;
  if(opId){
    const infoEl=document.getElementById(`op-info-${opId}`);
    if(infoEl){
      infoEl.innerHTML=sel.length
        ?`<div class="ms-tags">${sel.map(v=>`<span class="ms-tag">${msLbl(msId,v)}<button class="ms-tag-x" onclick="msRm('${msId}','${v}')">✕</button></span>`).join('')}</div>`
        :'';
    }
  }
}
function msRenderOpts(msId,opts){
  if(!opts.length) return '<div class="ms-empty">暂无数据</div>';
  const inst=msI[msId];
  return opts.map(o=>`<div class="ms-opt${inst.selected.has(o.value)?' sel':''}" onclick="msPick('${msId}','${o.value}',event)"><div class="ms-chk">${inst.selected.has(o.value)?'✓':''}</div><span class="ms-opt-lbl">${esc(o.label)}</span></div>`).join('');
}
function msLbl(msId,v){return msI[msId]?.options.find(o=>o.value===v)?.label||v;}
window.msTrigger=function(msId){
  const wrap=document.getElementById(`msw-${msId}`); if(!wrap) return;
  const was=wrap.classList.contains('open'); msCloseAll();
  if(!was){wrap.classList.add('open');const s=wrap.querySelector('.ms-search');if(s){s.value='';msFilter(msId,'');setTimeout(()=>s.focus(),30);}}
};
function msCloseAll(){document.querySelectorAll('.ms-wrap.open').forEach(el=>el.classList.remove('open'));}
document.addEventListener('click',e=>{if(!e.target.closest('.ms-wrap'))msCloseAll();});
window.msFilter=function(msId,q){
  const inst=msI[msId]; if(!inst) return;
  const el=document.getElementById(`mso-${msId}`); if(!el) return;
  const f=q?inst.options.filter(o=>o.label.toLowerCase().includes(q.toLowerCase())):inst.options;
  el.innerHTML=f.length?msRenderOpts(msId,f):'<div class="ms-empty">无匹配</div>';
};
window.msPick=function(msId,value,event){
  event.stopPropagation();
  const inst=msI[msId]; if(!inst) return;
  inst.selected.has(value)?inst.selected.delete(value):inst.selected.add(value);
  msRender(msId);
  const wrap=document.getElementById(`msw-${msId}`);
  if(wrap){wrap.classList.add('open');const s=wrap.querySelector('.ms-search');if(s)msFilter(msId,s.value);}
  if(inst.onChange) inst.onChange();
};
window.msRm=function(msId,value){
  const inst=msI[msId]; if(!inst) return;
  inst.selected.delete(value); msRender(msId);
  if(inst.onChange) inst.onChange();
};
function msGet(msId){return[...(msI[msId]?.selected||[])];}
function msSet(msId,values){const inst=msI[msId];if(!inst)return;inst.selected.clear();values.forEach(v=>inst.selected.add(v));}
function colOpts(){return headers.map((h,i)=>({value:String(i),label:h}));}

// ══ UNDO ══════════════════════════════════════════
function pushUndo(desc){
  undoStack.push({data:workingData.map(r=>[...r]),desc});
  renderUndoState();
}
function undoLast(){
  if(!undoStack.length) return;
  const snap=undoStack.pop();
  workingData=snap.data; afterChange(); renderUndoState();
  // reset op statuses
  ops.forEach(op=>{setStatus(op.id,'待分析','');setInfo(op.id,'');});
  toast(`已撤回：${snap.desc}`);
}
function renderUndoState(){
  const btn=document.getElementById('btn-undo');
  const lbl=document.getElementById('undo-label');
  if(btn) btn.disabled=undoStack.length===0;
  if(lbl) lbl.textContent=undoStack.length>0?`可撤回 ${undoStack.length} 步`:'';
}

// ══ RENDER TYPE GROUPS ════════════════════════════
function renderTypeGroups(){
  const container=document.getElementById('type-groups'); if(!container) return;
  const grouped={};
  TYPE_ORDER.forEach(t=>grouped[t]=[]);
  ops.forEach(op=>{if(grouped[op.type])grouped[op.type].push(op);});

  let html='';
  TYPE_ORDER.forEach(type=>{
    const group=grouped[type]; if(!group.length) return;
    html+=`<div class="type-group" id="tg-${type}">
      <div class="type-group-hd">
        <span class="type-name">${OP_NAMES[type]}</span>
        <button class="add-same" onclick="addOpOfType('${type}')">＋ 再添一条</button>
      </div>
      <div class="op-cards" id="tg-cards-${type}">
        ${group.map((op,idx)=>opCardHTML(op,idx+1)).join('')}
      </div>
    </div>`;
  });
  container.innerHTML=html;

  // init multi-selects
  ops.filter(op=>op.type==='delcol').forEach(op=>{
    msCreate(`dc-${op.id}`,colOpts(),'请选择要删除的列…',`ms-c-${op.id}`,()=>saveDelColParams(op.id));
    if(op.params?.colHeaders?.length) restoreDelColSel(op.id,op.params.colHeaders);
    else msRender(`dc-${op.id}`);
  });

  // restore delrow saved col
  ops.filter(op=>op.type==='delrow').forEach(op=>{
    const sel=document.getElementById(`delrow-col-${op.id}`); if(!sel) return;
    if(op.params?.colHeader==='__any__') sel.value='__any__';
    else if(op.params?.colHeader){const i=headers.indexOf(op.params.colHeader);if(i>=0)sel.value=String(i);}
  });
  // restore addcol
  ops.filter(op=>op.type==='addcol').forEach(op=>{
    const sel=document.getElementById(`addcol-after-${op.id}`); if(!sel||!op.params?.afterColHeader) return;
    const i=headers.indexOf(op.params.afterColHeader); if(i>=0) sel.value=String(i);
  });

  document.getElementById('ops-empty').style.display=ops.length===0?'block':'none';
}

function opCardHTML(op, seqInGroup){
  return `<div class="op-card" id="op-${op.id}">
    <div class="op-row">
      <div class="op-icon" title="${esc(OP_NAMES[op.type]||'操作')}">${OP_ICONS[op.type]||''}</div>
      <div class="op-seq">${seqInGroup}</div>
      <div class="op-controls">${controlsHTML(op)}</div>
      <span class="op-stat" id="op-stat-${op.id}">${op.statusText||'待分析'}</span>
      <button class="op-x" onclick="removeOp(${op.id})">✕</button>
    </div>
    <div class="op-info-row" id="op-info-${op.id}"></div>
  </div>`;
}

function controlsHTML(op){
  const id=op.id;
  if(op.type==='dedup'){
    return `<span style="font-size:12px;color:var(--light)">第4列起相邻行完全相同则删除</span>`;
  }
  if(op.type==='replace'){
    return `<div class="frow">
      <input type="text" id="rep-from-${id}" placeholder="查找" value="${esc(op.params?.from||'')}" style="flex:1;min-width:80px">
      <span style="color:var(--light);font-size:12px;flex-shrink:0">→</span>
      <input type="text" id="rep-to-${id}" placeholder="替换为（空则删除）" value="${esc(op.params?.to||'')}" style="flex:1.4;min-width:80px">
    </div>`;
  }
  if(op.type==='delrow'){
    const colOH=headers.map((h,i)=>`<option value="${i}">${esc(h)}</option>`).join('')+'<option value="__any__">任意列</option>';
    return `<div class="frow">
      <select class="sel-single" id="delrow-col-${id}" style="flex:1;min-width:0">${colOH}</select>
      <span style="color:var(--light);font-size:12px;flex-shrink:0">含</span>
      <input type="text" id="delrow-kw-${id}" placeholder="关键词" value="${esc(op.params?.kw||'')}" style="flex:1.5;min-width:60px">
      <span style="color:var(--light);font-size:12px;flex-shrink:0">则删行</span>
    </div>`;
  }
  if(op.type==='delcol'){
    return `<div id="ms-c-${id}" style="flex:1;min-width:0"></div>`;
  }
  if(op.type==='addcol'){
    const colOH=headers.map((h,i)=>`<option value="${i}">${esc(h)}</option>`).join('');
    return `<div class="frow">
      <select class="sel-single" id="addcol-after-${id}" style="flex:1;min-width:0">${colOH}</select>
      <span style="color:var(--light);font-size:12px;flex-shrink:0">后插入</span>
      <input type="number" id="addcol-count-${id}" value="${op.params?.count||1}" min="1" max="50">
      <span style="color:var(--light);font-size:12px;flex-shrink:0">列</span>
    </div>`;
  }
  return '';
}

// ══ OP MANAGEMENT ═════════════════════════════════
function addOp(){
  const type=document.getElementById('op-type-sel').value;
  if(!type){toast('请先选择操作类型');return;}
  addOpOfType(type);
  document.getElementById('op-type-sel').value='';
}

function addOpOfType(type){
  const id=++opCounter;
  // prepend to front of this type group
  const newOp={id,type,statusText:'待分析',statusClass:'',params:{}};
  // find index of first op of this type and insert before it
  const firstIdx=ops.findIndex(o=>o.type===type);
  if(firstIdx===-1) ops.push(newOp);
  else ops.splice(firstIdx,0,newOp);

  renderTypeGroups();

  // animate the new card
  setTimeout(()=>{
    const el=document.getElementById(`op-${id}`);
    if(el){el.classList.add('new-anim');el.scrollIntoView({behavior:'smooth',block:'nearest'});}
  },30);
}

function removeOp(id){
  ops=ops.filter(o=>o.id!==id);
  renderTypeGroups();
}

function clearOps(){
  ops=[];opCounter=0;renderTypeGroups();toast('操作已清空');
}

function getOp(id){return ops.find(o=>o.id===id);}

function setStatus(id,txt,cls){
  const op=getOp(id);if(!op)return;
  op.statusText=txt;op.statusClass=cls;
  const el=document.getElementById(`op-stat-${id}`);
  if(el){el.textContent=txt;el.className=`op-stat${cls?' '+cls:''}`;}
}

function setInfo(id,html){
  const el=document.getElementById(`op-info-${id}`);if(el)el.innerHTML=html;
}

function saveDelColParams(id){
  const op=getOp(id); if(!op) return;
  const names=msGet(`dc-${id}`).map(v=>headers[parseInt(v)]||v);
  op.params.colHeaders=[...names];
}

function restoreDelColSel(id,colHeaders){
  const msId=`dc-${id}`;
  const vals=colHeaders.map(name=>{const i=headers.indexOf(name);return i>=0?String(i):null;}).filter(v=>v!==null);
  msSet(msId,vals);msRender(msId);
}

function rebuildSelectors(id){
  const op=getOp(id);if(!op)return;
  if(op.type==='delrow'){
    const sel=document.getElementById(`delrow-col-${id}`);if(!sel)return;
    const cur=sel.value;
    sel.innerHTML=headers.map((h,i)=>`<option value="${i}">${esc(h)}</option>`).join('')+'<option value="__any__">任意列</option>';
    sel.value=cur;
  }
  if(op.type==='addcol'){
    const sel=document.getElementById(`addcol-after-${id}`);if(!sel)return;
    const cur=sel.value;
    sel.innerHTML=headers.map((h,i)=>`<option value="${i}">${esc(h)}</option>`).join('');
    sel.value=cur;
  }
  if(op.type==='delcol'){
    const msId=`dc-${id}`;
    const prev=msGet(msId);
    msCreate(msId,colOpts(),'请选择要删除的列…',`ms-c-${id}`,()=>saveDelColParams(id));
    msSet(msId,prev);msRender(msId);
  }
}

// ══ GLOBAL RUN ALL ════════════════════════════════
async function runAll(mode){
  if(!workingData.length){toast('请先导入文件');return;}
  if(!ops.length){toast('暂无操作');return;}
  for(const op of ops){
    if(mode==='analyze'){
      await analyzeOp(op.id);
    } else {
      const canExec=await analyzeOp(op.id);
      if(canExec) await execOp(op.id);
    }
    await new Promise(r=>setTimeout(r,60));
  }
  toast(mode==='analyze'?'全部分析完成':'全部执行完成');
}

async function analyzeOp(id){
  const op=getOp(id);if(!op)return false;
  if(op.type==='dedup')  return analyzeDedup(id);
  if(op.type==='replace')return analyzeReplace(id);
  if(op.type==='delrow') return analyzeDelRow(id);
  if(op.type==='delcol') return analyzeDelCol(id);
  if(op.type==='addcol') return analyzeAddCol(id);
  return false;
}
async function execOp(id){
  const op=getOp(id);if(!op)return;
  if(op.type==='dedup')  execDedup(id);
  else if(op.type==='replace') execReplace(id);
  else if(op.type==='delrow')  execDelRow(id);
  else if(op.type==='delcol')  execDelCol(id);
  else if(op.type==='addcol')  execAddCol(id);
}

// ══ DEDUP ══════════════════════════════════════════
function dedupKey(row){return JSON.stringify(row.slice(3).map(String));}

function analyzeDedup(id){
  if(headers.length<=3){
    setInfo(id,'<div class="info warn">列数不足4列，无法使用此规则</div>');
    setStatus(id,'列数不足','warn');return false;
  }
  const data=workingData.slice(1);
  const dupRows=[];
  for(let i=1;i<data.length;i++){
    if(dedupKey(data[i-1])===dedupKey(data[i])) dupRows.push(data[i]);
  }
  if(!dupRows.length){
    setInfo(id,'<div class="info ok">✓ 无相邻重复行</div>');
    setStatus(id,'无重复','ok');return false;
  }
  let html=`<div class="info warn">发现 <strong>${dupRows.length}</strong> 条相邻重复行（第4列起与上一行完全相同），执行后将删除：</div>`;
  html+='<div class="tbl-wrap"><table><thead><tr>'+headers.map(h=>`<th>${esc(h)}</th>`).join('')+'</tr></thead><tbody>';
  dupRows.forEach(row=>{html+='<tr class="dup-row">'+row.map(c=>`<td title="${esc(String(c))}">${c!==''?esc(String(c)):'—'}</td>`).join('')+'</tr>';});
  html+='</tbody></table></div>';
  setInfo(id,html);setStatus(id,`${dupRows.length} 条重复`,'warn');
  return true;
}
function execDedup(id){
  pushUndo('删除相邻重复行');
  const newData=[workingData[0]];
  const data=workingData.slice(1);
  for(let i=0;i<data.length;i++){
    if(i===0||dedupKey(data[i-1])!==dedupKey(data[i])) newData.push(data[i]);
  }
  const removed=workingData.length-newData.length;
  workingData=newData;afterChange();
  setInfo(id,`<div class="info ok">✓ 已删除 <strong>${removed}</strong> 条</div>`);
  setStatus(id,`已删 ${removed} 条`,'ok');
  toast(`已删除 ${removed} 条相邻重复行`);
}

// ══ REPLACE ════════════════════════════════════════
function analyzeReplace(id){
  const from=document.getElementById(`rep-from-${id}`)?.value||'';
  if(!from){setInfo(id,'<div class="info warn">请填写查找内容</div>');return false;}
  let count=0;
  workingData.forEach(r=>r.forEach(c=>{if(String(c).includes(from))count++;}));
  const to=document.getElementById(`rep-to-${id}`)?.value||'';
  const op=getOp(id);if(op){op.params.from=from;op.params.to=to;}
  if(!count){
    setInfo(id,`<div class="info">未找到「${esc(from)}」</div>`);
    setStatus(id,'无匹配','');return false;
  }
  setInfo(id,`<div class="info warn">找到 <strong>${count}</strong> 处，将替换为「${esc(to)||'（空）'}」</div>`);
  setStatus(id,`${count} 处匹配`,'warn');
  return true;
}
function execReplace(id){
  const from=document.getElementById(`rep-from-${id}`)?.value||'';
  const to=document.getElementById(`rep-to-${id}`)?.value||'';
  pushUndo(`替换「${from}」`);
  let count=0;
  workingData=workingData.map(r=>r.map(c=>{
    const s=String(c);if(s.includes(from)){count++;return s.replaceAll(from,to);}return c;
  }));
  afterChange();
  setInfo(id,`<div class="info ok">✓ 已替换 <strong>${count}</strong> 处</div>`);
  setStatus(id,`已替换 ${count}`,'ok');
  toast(`已替换 ${count} 处`);
}

// ══ DEL ROW ════════════════════════════════════════
function analyzeDelRow(id){
  const colVal=document.getElementById(`delrow-col-${id}`)?.value||'__any__';
  const kw=document.getElementById(`delrow-kw-${id}`)?.value||'';
  if(!kw){setInfo(id,'<div class="info warn">请填写关键词</div>');return false;}
  let count=0;
  workingData.slice(1).forEach(r=>{
    const hit=colVal==='__any__'?r.some(c=>String(c).includes(kw)):String(r[parseInt(colVal)]||'').includes(kw);
    if(hit)count++;
  });
  const colName=colVal==='__any__'?'任意列':(headers[parseInt(colVal)]||colVal);
  const op=getOp(id);if(op){op.params.colVal=colVal;op.params.colHeader=colVal!=='__any__'?(headers[parseInt(colVal)]||''):'__any__';op.params.kw=kw;}
  if(!count){
    setInfo(id,`<div class="info">在「${esc(colName)}」中未找到「${esc(kw)}」</div>`);
    setStatus(id,'无匹配','');return false;
  }
  setInfo(id,`<div class="info warn">「${esc(colName)}」中找到 <strong>${count}</strong> 行含「${esc(kw)}」，将整行删除</div>`);
  setStatus(id,`${count} 行匹配`,'warn');
  return true;
}
function execDelRow(id){
  const colVal=document.getElementById(`delrow-col-${id}`)?.value||'__any__';
  const kw=document.getElementById(`delrow-kw-${id}`)?.value||'';
  pushUndo(`删除含「${kw}」的行`);
  let removed=0;
  const newData=[workingData[0]];
  workingData.slice(1).forEach(r=>{
    const hit=colVal==='__any__'?r.some(c=>String(c).includes(kw)):String(r[parseInt(colVal)]||'').includes(kw);
    if(hit)removed++;else newData.push(r);
  });
  workingData=newData;afterChange();
  setInfo(id,`<div class="info ok">✓ 已删除 <strong>${removed}</strong> 行</div>`);
  setStatus(id,`已删 ${removed} 行`,'ok');
  toast(`已删除 ${removed} 行`);
}

// ══ DEL COL ════════════════════════════════════════
function analyzeDelCol(id){
  const sel=msGet(`dc-${id}`);
  saveDelColParams(id);
  if(!sel.length){setInfo(id,'');setStatus(id,'未选列','');return false;}
  const names=sel.map(v=>headers[parseInt(v)]||v);
  setInfo(id,`<div class="info warn">将删除 <strong>${sel.length}</strong> 列：${names.map(esc).join('、')}</div>`);
  setStatus(id,`${sel.length} 列`,'warn');
  return true;
}
function execDelCol(id){
  const sel=msGet(`dc-${id}`).map(v=>parseInt(v)).sort((a,b)=>b-a);
  if(!sel.length)return;
  pushUndo(`删除 ${sel.length} 列`);
  workingData=workingData.map(r=>{const row=[...r];sel.forEach(i=>row.splice(i,1));return row;});
  afterChange();
  setInfo(id,`<div class="info ok">✓ 已删除 <strong>${sel.length}</strong> 列</div>`);
  setStatus(id,`已删 ${sel.length} 列`,'ok');
  msI[`dc-${id}`]?.selected.clear();msRender(`dc-${id}`);
  toast(`已删除 ${sel.length} 列`);
}

// ══ ADD COL ════════════════════════════════════════
function analyzeAddCol(id){
  if(!workingData.length)return false;
  const afterIdx=parseInt(document.getElementById(`addcol-after-${id}`)?.value||0);
  const count=Math.max(1,parseInt(document.getElementById(`addcol-count-${id}`)?.value||1));
  const op=getOp(id);if(op){op.params.afterColHeader=headers[afterIdx]||'';op.params.count=count;}
  setInfo(id,`<div class="info warn">将在「${esc(headers[afterIdx]||'')}」列后添加 <strong>${count}</strong> 个空白列</div>`);
  setStatus(id,`+${count} 列`,'warn');
  return true;
}
function execAddCol(id){
  const afterIdx=parseInt(document.getElementById(`addcol-after-${id}`)?.value||0);
  const count=Math.max(1,parseInt(document.getElementById(`addcol-count-${id}`)?.value||1));
  pushUndo(`添加 ${count} 个空白列`);
  workingData=workingData.map(r=>{const row=[...r];for(let i=0;i<count;i++)row.splice(afterIdx+1+i,0,'');return row;});
  afterChange();
  setInfo(id,`<div class="info ok">✓ 已添加 <strong>${count}</strong> 个空白列</div>`);
  setStatus(id,`已添 ${count} 列`,'ok');
  toast(`已添加 ${count} 个空白列`);
}

// ══ AFTER CHANGE ══════════════════════════════════
function afterChange(){
  refreshHeaders();renderStats();
  ops.forEach(op=>rebuildSelectors(op.id));
}

// ══ CONFIG ════════════════════════════════════════
function gatherConfig(){
  return{
    version:4,
    operations:ops.map(op=>{
      const c={type:op.type,params:{...op.params}};
      if(op.type==='replace'){
        c.params.from=document.getElementById(`rep-from-${op.id}`)?.value||'';
        c.params.to=document.getElementById(`rep-to-${op.id}`)?.value||'';
      }
      if(op.type==='delrow'){
        const colEl=document.getElementById(`delrow-col-${op.id}`);
        const colIdx=colEl?.value||'__any__';
        c.params.colVal=colIdx;
        c.params.colHeader=colIdx!=='__any__'?(headers[parseInt(colIdx)]||''):'__any__';
        c.params.kw=document.getElementById(`delrow-kw-${op.id}`)?.value||'';
      }
      if(op.type==='delcol'){
        const sel=msGet(`dc-${op.id}`);
        c.params.colHeaders=sel.map(v=>headers[parseInt(v)]||v);
      }
      if(op.type==='addcol'){
        const idx=parseInt(document.getElementById(`addcol-after-${op.id}`)?.value||0);
        c.params.afterColHeader=headers[idx]||'';
        c.params.count=parseInt(document.getElementById(`addcol-count-${op.id}`)?.value||1);
      }
      return c;
    })
  };
}

function exportConfig(){
  const cfg=gatherConfig();
  const blob=new Blob([JSON.stringify(cfg,null,2)],{type:'application/json'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');a.href=url;a.download='scrub-config.json';a.click();
  URL.revokeObjectURL(url);toast('配置已导出');
}

function importConfig(e){
  const file=e.target.files[0];if(!file)return;
  const reader=new FileReader();
  reader.onload=ev=>{
    try{
      const cfg=JSON.parse(ev.target.result);
      if(!cfg.version||!cfg.operations){toast('配置格式不兼容');return;}
      clearOps();
      cfg.operations.forEach(item=>{
        const id=++opCounter;
        ops.push({id,type:item.type,statusText:'待分析',statusClass:'',params:item.params||{}});
      });
      renderTypeGroups();
      document.getElementById('cfg-status').textContent='配置已导入 ✓';
      toast('配置导入成功');
    }catch(err){toast('配置文件格式错误');}
  };
  reader.readAsText(file);e.target.value='';
}

// ══ EXPORT FILE ═══════════════════════════════════
function exportFile(){
  if(!workingData.length){toast('没有数据可导出');return;}
  const ws=XLSX.utils.aoa_to_sheet(workingData);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Sheet1');
  XLSX.writeFile(wb,`${fileName}_处理后.xlsx`);
  toast('文件已下载');
}

// ══ UTILS ═════════════════════════════════════════
function esc(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
function toast(msg){
  const el=document.getElementById('toast');
  el.textContent=msg;el.classList.add('show');
  clearTimeout(el._t);el._t=setTimeout(()=>el.classList.remove('show'),2500);
}
