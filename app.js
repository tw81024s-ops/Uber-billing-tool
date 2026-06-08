/* =====================================================================
   UBER 帳單整理工具
   所有處理在瀏覽器內完成，資料不會上傳到任何伺服器
===================================================================== */

(function(){
'use strict';

// ---- 狀態 ----
const state = {
  main:   { rows:null, headers:null, mapping:null, merged:null },  // 原始帳單
  huogu:  { rows:null, headers:null, mapping:null },                // 貨故請賠
  xukou:  { rows:null, headers:null, mapping:null, adjustments:null, dupMap:null },                // 虛扣／重複收款（多分頁）
  adjustments: [],   // 右塊 [{order, amount, note}]
  manual: [],        // 待人工確認 [{order, ...}]
  negSummary: [],    // B 欄負數單號對應狀態 [{order, mainAmount, status, ...}]
};

// ---- 工具 ----
const $ = sel => document.querySelector(sel);
const $$ = sel => document.querySelectorAll(sel);

function toast(msg, type){
  const t = $('#toast');
  t.textContent = msg;
  t.className = 'toast show' + (type==='err'?' err':'');
  clearTimeout(t._timer);
  t._timer = setTimeout(()=>{ t.className='toast'; }, 2400);
}

// 讓 UI 有機會更新（不然同步重活會讓進度條卡住）
function tick(){
  return new Promise(r => requestAnimationFrame(()=>setTimeout(r, 0)));
}

// 進度條控制：scope = 'main' | 'huogu' | 'xukou'；pct=null 隱藏
function setProgress(scope, pct, label){
  const el = document.getElementById('prog_' + scope);
  if (!el) return;
  if (pct == null){
    el.classList.add('hidden');
    el.classList.remove('done');
    return;
  }
  el.classList.remove('hidden');
  const pctNum = Math.max(0, Math.min(1, pct));
  el.querySelector('.bar-fill').style.width = (pctNum*100).toFixed(0) + '%';
  el.querySelector('.bar-label').textContent = label || `處理中… ${(pctNum*100).toFixed(0)}%`;
  if (pctNum >= 1) el.classList.add('done');
  else el.classList.remove('done');
}

function fmt(n){
  if (n===null || n===undefined || n==='') return '';
  if (typeof n !== 'number') return n;
  return n.toLocaleString('zh-TW', {maximumFractionDigits:0});
}

function normalizeHeader(s){
  return String(s||'').trim().replace(/\s+/g,'').toLowerCase();
}

// ---- 智慧判讀欄位 ----
// 規則：以「欄位名稱關鍵字」為主，輔以「資料樣本特徵」
const FIELD_RULES = {
  orderId: {
    label: '商家訂單 ID',
    keywords: ['商家訂單id','商家訂單','商家訂單編號','external_order_id','externalorderid','order_id','virtual_order_id','虛扣單號','查貨號碼'],
    // 12-13 位數字串
    detectByData: vals => {
      const sample = vals.slice(0, 50).filter(v => v != null && v !== '');
      if (sample.length < 3) return 0;
      const hit = sample.filter(v => /^\d{10,15}$/.test(String(v).trim())).length;
      return hit / sample.length;
    }
  },
  amount: {
    label: '帳單金額',
    keywords: ['帳單金額','金額','amount','delivery_fee','deliveryfee','含稅','扣款金額(含稅)','應退金額'],
    detectByData: vals => {
      const sample = vals.slice(0, 50).filter(v => v != null && v !== '');
      if (sample.length < 3) return 0;
      const hit = sample.filter(v => typeof v === 'number' || /^-?\d+(\.\d+)?$/.test(String(v).trim())).length;
      return hit / sample.length;
    }
  },
  uberId: {
    label: 'Uber 訂單 ID',
    keywords: ['uber訂單id','uber訂單','uber_order_id','uberorderid'],
    detectByData: vals => {
      const sample = vals.slice(0, 50).filter(v => v != null && v !== '');
      if (sample.length < 3) return 0;
      const hit = sample.filter(v => /^#?[0-9a-f]{4,6}$/i.test(String(v).trim())).length;
      return hit / sample.length;
    }
  },
  amountIncTax: {
    label: '扣款金額（含稅）',
    keywords: ['扣款金額(含稅)','扣款金額（含稅）','含稅','含稅金額'],
  },
  shouldRefund: {
    label: '應退金額',
    keywords: ['應退金額','應退','should_refund'],
  },
};

function detectColumns(headers, data, fields){
  // fields: [{key, required}]
  const result = {};
  const used = new Set();
  for (const f of fields){
    const rule = FIELD_RULES[f.key];
    let bestIdx = -1, bestScore = 0;
    for (let i=0; i<headers.length; i++){
      if (used.has(i)) continue;
      const h = normalizeHeader(headers[i]);
      let score = 0;
      // 1) 名稱命中
      for (const kw of (rule.keywords||[])){
        if (h === normalizeHeader(kw)) { score = Math.max(score, 1.0); break; }
        if (h.includes(normalizeHeader(kw))) score = Math.max(score, 0.85);
      }
      // 2) 資料樣本判斷
      if (rule.detectByData){
        const colVals = data.map(r => r[i]);
        const dataScore = rule.detectByData(colVals);
        score = Math.max(score, dataScore * 0.7);  // 純資料判斷上限 0.7
      }
      if (score > bestScore){ bestScore = score; bestIdx = i; }
    }
    if (bestIdx >= 0 && bestScore >= 0.4){
      result[f.key] = { index: bestIdx, score: bestScore, guessed: true };
      used.add(bestIdx);
    } else {
      result[f.key] = { index: -1, score: 0, guessed: false };
    }
  }
  return result;
}

// ---- 讀取 Excel/CSV ----
function readWorkbook(file){
  return new Promise((resolve, reject)=>{
    const reader = new FileReader();
    reader.onload = e => {
      try{
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, {type:'array', cellDates:false, raw:true});
        resolve(wb);
      }catch(err){ reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function sheetToRows(ws){
  const arr = XLSX.utils.sheet_to_json(ws, {header:1, defval:null, raw:true, blankrows:false});
  if (!arr.length) return { headers:[], rows:[] };
  // 找表頭：通常是第一列非空白的
  let headerRow = 0;
  for (let i=0; i<Math.min(arr.length,3); i++){
    if (arr[i] && arr[i].some(v => v!=null && v!=='')) { headerRow = i; break; }
  }
  const headers = arr[headerRow].map(h => h==null ? '' : String(h));
  const rows = arr.slice(headerRow+1).filter(r => r && r.some(v => v!=null && v!==''));
  return { headers, rows };
}

// ---- 渲染欄位對應 ----
function renderMapping(containerEl, headers, fields, mapping, onChange){
  containerEl.innerHTML = '';
  for (const f of fields){
    const rule = FIELD_RULES[f.key];
    const cur = mapping[f.key];
    const row = document.createElement('div');
    row.className = 'map-row' + (cur && cur.guessed ? ' guessed' : '');
    row.innerHTML = `
      <label>${rule.label}${f.required?' <span style="color:#FF1F8F">*</span>':''}</label>
      <select data-key="${f.key}">
        <option value="-1">— 不對應 —</option>
        ${headers.map((h,i)=>`<option value="${i}" ${cur && cur.index===i?'selected':''}>${(h||'(空白欄)').replace(/</g,'&lt;')}</option>`).join('')}
      </select>
    `;
    containerEl.appendChild(row);
    row.querySelector('select').addEventListener('change', e=>{
      mapping[f.key] = { index: parseInt(e.target.value, 10), score: 1, guessed: false };
      row.classList.remove('guessed');
      if (onChange) onChange();
    });
  }
}

// =============================================================
// 處理流程
// =============================================================

// 1) 主檔處理：讀取＋判讀
async function loadMain(file){
  $('#hintMain').textContent = '處理中…';
  setProgress('main', 0.05, '讀取檔案中…');
  await tick();
  try{
    const wb = await readWorkbook(file);
    setProgress('main', 0.45, '解析 Excel 內容…');
    await tick();
    const ws = wb.Sheets[wb.SheetNames[0]];
    const { headers, rows } = sheetToRows(ws);
    if (!headers.length || !rows.length){
      throw new Error('檔案無有效資料');
    }
    setProgress('main', 0.85, `智慧判讀欄位中…（${rows.length.toLocaleString()} 筆）`);
    await tick();
    state.main.headers = headers;
    state.main.rows = rows;
    state.main.mapping = detectColumns(headers, rows, [
      { key:'orderId', required:true },
      { key:'amount',  required:true },
    ]);
    // 欄位改動時自動重新合併
    renderMapping($('#mapGridMain'), headers, [
      { key:'orderId', required:true },
      { key:'amount',  required:true },
    ], state.main.mapping, ()=>{ previewMerge(); });
    $('#mapMain').classList.remove('hidden');
    $('#dropMain').classList.add('is-loaded');
    $('#hintMain').textContent = `已讀取：${file.name} ｜ ${rows.length.toLocaleString()} 筆`;
    setProgress('main', 1, '讀取完成 ✓');
    await tick();
    setTimeout(()=>setProgress('main', null), 900);
    toast('已讀取原始帳單');
    // 自動跑合併（不用使用者按按鈕）
    await previewMerge();
  }catch(err){
    console.error(err);
    setProgress('main', null);
    $('#hintMain').textContent = '讀取失敗：' + (err.message || err);
    toast('讀取失敗', 'err');
  }
}

// 2) 合併＋預覽
async function previewMerge(){
  const m = state.main.mapping;
  if (!m.orderId || m.orderId.index < 0 || !m.amount || m.amount.index < 0){
    toast('請先選擇商家訂單 ID 與帳單金額欄位', 'err');
    return false;
  }
  const idCol = m.orderId.index, amtCol = m.amount.index;
  const merged = new Map();
  let srcTotal = 0;
  const total = state.main.rows.length;
  const CHUNK = 8000;

  setProgress('main', 0, `合併中… 0 / ${total.toLocaleString()}`);
  await tick();

  for (let i=0; i<total; i+=CHUNK){
    const end = Math.min(i+CHUNK, total);
    for (let j=i; j<end; j++){
      const r = state.main.rows[j];
      const id = r[idCol];
      if (id == null || id === '') continue;
      const key = String(id).trim();
      const amt = Number(r[amtCol]) || 0;
      srcTotal += amt;
      if (merged.has(key)){
        const g = merged.get(key);
        g.amount += amt;
        g.count += 1;
      } else {
        merged.set(key, { orderId:key, amount:amt, count:1 });
      }
    }
    setProgress('main', end/total, `合併中… ${end.toLocaleString()} / ${total.toLocaleString()}`);
    await tick();
  }

  const mergedArr = Array.from(merged.values());
  const mergedTotal = mergedArr.reduce((s,g)=>s+g.amount, 0);
  state.main.merged = mergedArr;

  // 摘要
  $('#srcRows').textContent = total.toLocaleString();
  $('#mergedRows').textContent = mergedArr.length.toLocaleString();
  $('#dedupedCount').textContent = (total - mergedArr.length).toLocaleString();
  $('#srcTotal').textContent = fmt(Math.round(srcTotal));
  $('#mergedTotal').textContent = fmt(Math.round(mergedTotal));
  const consistent = Math.abs(srcTotal - mergedTotal) < 0.5;
  $('#reconBadge').className = 'badge ' + (consistent?'ok':'bad');
  $('#reconBadge').textContent = consistent ? '總額一致 ✓' : '總額不一致 ✗';
  $('#resultMain').classList.remove('hidden');

  setProgress('main', 1, `合併完成 ✓ ${mergedArr.length.toLocaleString()} 筆`);
  await tick();
  setTimeout(()=>setProgress('main', null), 1000);

  // 自動跑後續比對 & 預覽
  runAdjustments();
  renderPreview();
  return true;
}

// 3) 貨故請賠
async function loadHuogu(file){
  $('#hintHuogu').textContent = '處理中…';
  setProgress('huogu', 0.1, '讀取檔案中…');
  await tick();
  try{
    const wb = await readWorkbook(file);
    setProgress('huogu', 0.5, '解析 Excel 內容…');
    await tick();
    const ws = wb.Sheets[wb.SheetNames[0]];
    const { headers, rows } = sheetToRows(ws);
    setProgress('huogu', 0.85, '智慧判讀欄位中…');
    await tick();
    state.huogu.headers = headers;
    state.huogu.rows = rows;
    state.huogu.mapping = detectColumns(headers, rows, [
      { key:'orderId', required:true },
      { key:'amountIncTax', required:true },
    ]);
    renderMapping($('#mapGridHuogu'), headers, [
      { key:'orderId', required:true },
      { key:'amountIncTax', required:true },
    ], state.huogu.mapping, ()=>{ if (state.main.merged) { runAdjustments(); renderPreview(); }});
    $('#mapHuogu').classList.remove('hidden');
    $('#dropHuogu').classList.add('is-loaded');
    $('#hintHuogu').textContent = `已讀取：${file.name} ｜ ${rows.length.toLocaleString()} 筆`;
    if (state.main.merged) { runAdjustments(); renderPreview(); }
    setProgress('huogu', 1, '讀取完成 ✓');
    await tick();
    setTimeout(()=>setProgress('huogu', null), 900);
    toast('已讀取貨故請賠檔');
  }catch(err){
    console.error(err);
    setProgress('huogu', null);
    $('#hintHuogu').textContent = '讀取失敗：' + (err.message || err);
    toast('讀取失敗', 'err');
  }
}

// 4) 虛扣／重複收款（重複收款 Summary 分頁）
async function loadXukou(file){
  $('#hintXukou').textContent = '處理中…';
  setProgress('xukou', 0.1, '讀取檔案中…');
  await tick();
  try{
    const wb = await readWorkbook(file);
    setProgress('xukou', 0.55, '解析 Excel 內容…');
    await tick();
    // 找「重複收款 Summary」分頁
    let summaryRows = null;
    for (const name of wb.SheetNames){
      if (name.replace(/\s+/g,'').includes('重複收款Summary') || name.includes('Summary')){
        const arr = XLSX.utils.sheet_to_json(wb.Sheets[name], {header:1, defval:null, raw:true, blankrows:false});
        summaryRows = arr;
        break;
      }
    }
    // 找「虛扣單號總整理」分頁 → 建立 虛扣單號 -> {原單號, 調整目的, 金額}
    const xukouMap = new Map();
    for (const name of wb.SheetNames){
      if (name.includes('虛扣單號總整理')){
        const arr = XLSX.utils.sheet_to_json(wb.Sheets[name], {header:1, defval:null, raw:true, blankrows:false});
        let lastPurpose = null, lastSrcOrder = null;
        for (let i=1; i<arr.length; i++){
          const r = arr[i];
          if (!r || r[1]==null || r[1]==='') continue;
          const srcOrder = r[0];
          const xukouId = String(r[1]).trim();
          let purpose = r[2];
          const amount = r[3];
          // 續行：本列沒有調整目的時，繼承上一筆（例如「重複收款單號調整」一組多筆）
          if (!purpose && lastPurpose){
            purpose = lastPurpose;
          } else if (purpose){
            lastPurpose = purpose;
            lastSrcOrder = srcOrder;
          }
          xukouMap.set(xukouId, {
            srcOrder: srcOrder != null && srcOrder !== '' ? String(srcOrder).trim() : null,
            purpose: purpose || null,
            amount: amount != null && amount !== '' ? Number(amount) : null
          });
        }
        break;
      }
    }
    state.xukou.adjustments = xukouMap;
    // 找「重複收款 - N 張重複宅單」分頁 → 建立 單號 -> {times 收款次數, refund 應退金額}
    // 表頭：external_order_id / delivery_fee(=應退金額) / duplicated_check(=收款次數)
    const dupMap = new Map();
    for (const name of wb.SheetNames){
      const m = name.replace(/\s+/g,'').match(/重複收款[-－](\d+)張重複宅單/);
      if (!m) continue;
      const arr = XLSX.utils.sheet_to_json(wb.Sheets[name], {header:1, defval:null, raw:true, blankrows:false});
      for (let i=1; i<arr.length; i++){
        const r = arr[i];
        if (!r || r[0]==null || r[0]==='') continue;
        const id = String(r[0]).trim();
        if (!/^\d{6,}$/.test(id)) continue;            // 僅收訂單號
        const refund = Number(r[1]) || 0;               // delivery_fee 欄 = 應退金額
        const times  = Number(r[2]) || parseInt(m[1],10); // duplicated_check = 收款次數
        dupMap.set(id, { times, refund });
      }
    }
    state.xukou.dupMap = dupMap;
    setProgress('xukou', 0.9, '判讀分頁中…');
    await tick();
    state.xukou.summary = summaryRows;
    $('#dropXukou').classList.add('is-loaded');
    const sheetCount = wb.SheetNames.length;
    $('#hintXukou').textContent = `已讀取：${file.name} ｜ ${sheetCount} 個分頁`;
    if (state.main.merged) { runAdjustments(); renderPreview(); }
    setProgress('xukou', 1, '讀取完成 ✓');
    await tick();
    setTimeout(()=>setProgress('xukou', null), 900);
    toast('已讀取虛扣／重複收款檔');
  }catch(err){
    console.error(err);
    setProgress('xukou', null);
    $('#hintXukou').textContent = '讀取失敗：' + (err.message || err);
    toast('讀取失敗', 'err');
  }
}

// 5) 跑比對 → 產生 adjustments 與 manual
function runAdjustments(){
  state.adjustments = [];
  state.manual = [];
  if (!state.main.merged) return;
  const mainMap = new Map(state.main.merged.map(g => [g.orderId, g.amount]));

  // ---- 貨故請賠 ----
  if (state.huogu.rows && state.huogu.mapping &&
      state.huogu.mapping.orderId.index >= 0 &&
      state.huogu.mapping.amountIncTax.index >= 0){
    const idCol = state.huogu.mapping.orderId.index;
    const amtCol = state.huogu.mapping.amountIncTax.index;
    // 同查貨號碼加總含稅
    const huoguSum = new Map();
    const huoguCnt = new Map();
    for (const r of state.huogu.rows){
      const id = r[idCol];
      if (id == null || id === '') continue;
      const key = String(id).trim();
      const amt = Number(r[amtCol]) || 0;
      huoguSum.set(key, (huoguSum.get(key)||0) + amt);
      huoguCnt.set(key, (huoguCnt.get(key)||0) + 1);
    }
    // 為每個查貨號碼產生調整列
    for (const [id, sum] of huoguSum.entries()){
      const claim = Math.round(sum);
      // 貨故請賠列（轉負數）
      state.adjustments.push({ order:id, amount:-claim, note:'貨故請賠' });
      // 運費回溯：主檔金額絕對值 − 貨故含稅加總
      if (mainMap.has(id)){
        const mainAmt = Math.round(Math.abs(mainMap.get(id)));
        const diff = mainAmt - claim;
        if (diff === 0){
          // 剛好對上，不產生運費回溯列
        } else if ([41, 45, 47, 82, 90, 123].includes(diff)){
          // 41/45 是運費；82=41*2、90=45*2、123=41*3 等情況
          state.adjustments.push({ order:id, amount:-diff, note:'貨故運費回溯' });
        } else {
          state.manual.push({
            order:id, mainAmount:mainMap.get(id), claim:claim, diff:diff,
            type:'貨故差額異常',
            note:'差額非運費金額（41/45），需人工確認'
          });
        }
      } else {
        state.manual.push({
          order:id, mainAmount:null, claim:claim, diff:null,
          type:'主檔查無',
          note:'貨故請賠檔有此單號，但主檔合併後查無，需人工確認'
        });
      }
    }
  }

  // ---- 重複收款 Summary ----
  if (state.xukou.summary){
    const rows = state.xukou.summary;
    // 規格：第一欄 label，第二欄起為值
    // 結構像：虛扣單號 300490107094 / null 300490463114 / 應退金額 89462
    let orders = [];
    let shouldRefund = null;
    for (const r of rows){
      if (!r) continue;
      const label = String(r[0]||'').trim();
      const val = r[1];
      if (val == null || val === '') continue;
      if (label.includes('虛扣') || label === ''){
        // 虛扣單號或續行
        if (/^\d{10,15}$/.test(String(val).trim())) orders.push(String(val).trim());
      }
      if (label.includes('應退')){
        shouldRefund = Number(val) || 0;
      }
    }
    if (orders.length > 0 && shouldRefund != null){
      // 列入待人工確認，讓 user 自行拆分
      state.manual.push({
        order: orders.join(' / '),
        mainAmount: null,
        claim: shouldRefund,
        diff: null,
        type:'重複收款應退',
        note:`應退 ${fmt(shouldRefund)} 元，涉及單號 ${orders.length} 個，拆分方式由人工決定`
      });
    }
  }

  // ---- 合併後負金額但未對到任何已知來源（貨故請賠檔／虛扣單號總整理）→ 標記為待查 ----
  // user 需向 Uber 確認實際單號，再判斷是貨故或其他原因
  const xukouIds = (state.xukou.adjustments && state.xukou.adjustments.size)
    ? new Set(state.xukou.adjustments.keys()) : new Set();
  const huoguIds = new Set();
  const huoguSumMap = new Map();   // 查貨號碼 -> 含稅加總
  if (state.huogu.rows && state.huogu.mapping && state.huogu.mapping.orderId.index >= 0){
    const c = state.huogu.mapping.orderId.index;
    const ac = state.huogu.mapping.amountIncTax ? state.huogu.mapping.amountIncTax.index : -1;
    for (const r of state.huogu.rows){
      if (r[c] != null && r[c] !== ''){
        const k = String(r[c]).trim();
        huoguIds.add(k);
        if (ac >= 0) huoguSumMap.set(k, (huoguSumMap.get(k)||0) + (Number(r[ac])||0));
      }
    }
  }
  // 建立「B 欄負數單號對應表」：每個合併後為負的訂單，判斷它對應到什麼
  state.negSummary = [];
  for (const g of state.main.merged){
    if (g.amount >= 0) continue;
    let status, srcOrder='', purpose='', detailAmount='';
    if (xukouIds.has(g.orderId)){
      const x = state.xukou.adjustments.get(g.orderId);
      status = '虛扣調整';
      srcOrder = x.srcOrder || '';
      purpose = x.purpose || '';
      detailAmount = x.amount != null ? x.amount : '';
    } else if (huoguIds.has(g.orderId)){
      status = '貨故請賠';
      purpose = '貨故請賠';
      detailAmount = huoguSumMap.has(g.orderId) ? -Math.round(huoguSumMap.get(g.orderId)) : '';
    } else {
      status = '需確認';
      purpose = '需確認';
    }
    state.negSummary.push({
      order: g.orderId, mainAmount: Math.round(g.amount),
      status, srcOrder, purpose, detailAmount
    });
    // 對不到任何來源的負金額 → 進待人工確認
    if (status === '需確認'){
      state.manual.push({
        order: g.orderId,
        mainAmount: g.amount,
        claim: null, diff: null,
        type: '負金額待查',
        note: '合併後為負金額，未對到貨故請賠或虛扣單號，需向 Uber 確認實際單號（可能是貨故或其他原因）'
      });
    }
  }
}

// 6) 預覽渲染
function renderPreview(){
  $('#previewBox').classList.remove('hidden');

  // 處理結果摘要
  const summary = $('#processSummary');
  if (summary){
    const chips = [];
    if (state.main.merged){
      chips.push(`<span class="sum-chip"><span>合併帳單</span><b>${state.main.merged.length.toLocaleString()}</b><span>筆</span></span>`);
    }
    if (state.adjustments.length){
      chips.push(`<span class="sum-chip ok"><span>貨故調整</span><b>${state.adjustments.length}</b><span>列</span></span>`);
    } else if (state.huogu.rows){
      chips.push(`<span class="sum-chip empty"><span>貨故調整</span><b>0</b><span>列</span></span>`);
    }
    if (state.xukou.adjustments && state.xukou.adjustments.size){
      const mainSet = state.main.merged ? new Set(state.main.merged.map(g=>g.orderId)) : new Set();
      let matched = 0;
      for (const k of state.xukou.adjustments.keys()) if (mainSet.has(k)) matched++;
      const cls = matched > 0 ? 'ok' : 'warn';
      chips.push(`<span class="sum-chip ${cls}"><span>虛扣對應</span><b>${matched}</b><span>/ ${state.xukou.adjustments.size} 筆已入帳</span></span>`);
    } else if (state.xukou.summary){
      chips.push(`<span class="sum-chip empty"><span>虛扣對應</span><b>0</b><span>筆</span></span>`);
    }
    if (state.manual.length){
      chips.push(`<span class="sum-chip warn"><span>待人工確認</span><b>${state.manual.length}</b><span>筆</span></span>`);
    }
    summary.innerHTML = chips.join('');
  }

  // 左塊
  const leftPane = $('#paneLeft');
  if (state.main.merged && state.main.merged.length){
    const sample = state.main.merged.slice(0, 200);
    leftPane.innerHTML = `
      <table>
        <thead><tr><th>商家訂單 ID</th><th>加總帳單金額</th><th style="text-align:right">合併筆數</th></tr></thead>
        <tbody>
          ${sample.map(g => `
            <tr>
              <td>${g.orderId}</td>
              <td class="num ${g.amount<0?'neg':''}">${fmt(Math.round(g.amount))}</td>
              <td class="num">${g.count}</td>
            </tr>`).join('')}
        </tbody>
      </table>
      ${state.main.merged.length > 200 ? `<div class="empty" style="padding:14px">… 僅顯示前 200 筆，匯出 Excel 包含全部 ${state.main.merged.length.toLocaleString()} 筆</div>` : ''}
    `;
  } else {
    leftPane.innerHTML = '<div class="empty">尚無資料</div>';
  }

  // 右塊
  const rightPane = $('#paneRight');
  if (state.adjustments.length){
    rightPane.innerHTML = `
      <table>
        <thead><tr><th>單號</th><th>金額</th><th>備註</th></tr></thead>
        <tbody>
          ${state.adjustments.map(a => `
            <tr>
              <td>${a.order}</td>
              <td class="num ${a.amount<0?'neg':''}">${fmt(a.amount)}</td>
              <td>${a.note}</td>
            </tr>`).join('')}
        </tbody>
      </table>
    `;
  } else {
    rightPane.innerHTML = '<div class="empty">尚無調整明細（請上傳貨故請賠檔以產生）</div>';
  }

  // 待人工
  const manualPane = $('#paneManual');
  if (state.manual.length){
    manualPane.innerHTML = `
      <table>
        <thead><tr><th>單號</th><th>類型</th><th>主檔金額</th><th>請賠/應退</th><th>差額</th><th>說明</th></tr></thead>
        <tbody>
          ${state.manual.map(m => `
            <tr>
              <td>${m.order}</td>
              <td>${m.type}</td>
              <td class="num">${m.mainAmount==null?'—':fmt(Math.round(m.mainAmount))}</td>
              <td class="num">${m.claim==null?'—':fmt(Math.round(m.claim))}</td>
              <td class="num ${m.diff<0?'neg':''}">${m.diff==null?'—':fmt(m.diff)}</td>
              <td>${m.note}</td>
            </tr>`).join('')}
        </tbody>
      </table>
    `;
  } else {
    manualPane.innerHTML = '<div class="empty">沒有待人工確認的項目</div>';
  }

  // 虛扣對應：把主檔合併後的單號 跟 虛扣單號總整理 對照
  const xukouPane = $('#paneXukou');
  const xukouMap = state.xukou.adjustments;
  if (xukouMap && xukouMap.size && state.main.merged){
    const mainSet = new Map(state.main.merged.map(g => [g.orderId, g.amount]));
    const matched = [], unmatched = [];
    for (const [xid, info] of xukouMap.entries()){
      if (mainSet.has(xid)){
        matched.push({ xukouId:xid, mainAmount:Math.round(mainSet.get(xid)), ...info });
      } else {
        unmatched.push({ xukouId:xid, ...info });
      }
    }
    const renderRow = m => `
      <tr>
        <td>${m.xukouId}</td>
        <td>${m.srcOrder||'—'}</td>
        <td>${m.purpose||'—'}</td>
        <td class="num ${m.amount<0?'neg':''}">${m.amount==null?'—':fmt(m.amount)}</td>
        <td class="num ${m.mainAmount!=null && m.mainAmount<0?'neg':''}">${m.mainAmount==null?'未入帳':fmt(m.mainAmount)}</td>
      </tr>`;
    xukouPane.innerHTML = `
      <div style="padding:10px 14px; background:var(--pink-50); font-size:12.5px; color:var(--ink-700); border-bottom:1px solid var(--pink-100);">
        虛扣單號總整理 ${xukouMap.size} 筆 ｜ 已入帳 <b style="color:var(--ok)">${matched.length}</b> 筆 ｜ 未入帳 <b style="color:var(--warn)">${unmatched.length}</b> 筆
      </div>
      <table>
        <thead><tr><th>虛扣單號</th><th>原單號</th><th>調整目的</th><th>表內金額</th><th>主檔金額</th></tr></thead>
        <tbody>
          ${matched.map(renderRow).join('')}
          ${unmatched.map(renderRow).join('')}
        </tbody>
      </table>
    `;
  } else if (xukouMap && xukouMap.size){
    xukouPane.innerHTML = '<div class="empty">請先上傳原始帳單以進行對應</div>';
  } else {
    xukouPane.innerHTML = '<div class="empty">未上傳虛扣／重複收款檔（可略過）</div>';
  }
}

// 7) 匯出 Excel
function exportExcel(){
  if (!state.main.merged){
    toast('請先上傳並合併原始帳單', 'err');
    return;
  }
  const wb = XLSX.utils.book_new();
  const xukouMap = state.xukou.adjustments || new Map();
  const dupMap = state.xukou.dupMap || new Map();   // 重複收款：單號 -> {times, refund}
  // 負數單號對應表（orderId -> 對應狀態），用於同列標註
  const negMap = new Map((state.negSummary||[]).map(n => [n.order, n]));

  // -- Final 分頁（9 欄）--
  // 單號｜原單號｜原帳單合併金額｜備註｜已通知調整金額｜重複收款次數｜待人工修改｜修改後實際金額｜貨故請賠
  const aoa = [['單號','原單號','原帳單合併金額','備註','已通知調整金額','重複收款次數','待人工修改','修改後實際金額','貨故請賠']];
  // 貨故請賠含稅對照（單號 -> 請賠含稅金額，正數）
  const huoguClaim = new Map();
  for (const a of (state.adjustments || [])){
    if (a.note === '貨故請賠') huoguClaim.set(a.order, Math.round(-a.amount)); // a.amount 為負，轉正
  }
  // 待人工確認單號集合（G 欄標記用）
  const manualSet = new Set((state.manual || []).flatMap(m => String(m.order || '').split('/').map(s => s.trim())));
  let matchCount = 0;
  for (const g of state.main.merged){
    const id = g.orderId;
    const C = Math.round(g.amount);
    const dup = dupMap.get(id);
    const neg = negMap.get(id);
    let B='', D='', E='', F='', G='', I='';
    if (huoguClaim.has(id)){
      // 貨故請賠：E＝-請賠含稅，I 欄標記
      E = -huoguClaim.get(id);
      I = '貨故請賠';
      matchCount++;
    } else if (neg && neg.status === '虛扣調整'){
      B = neg.srcOrder || '';
      D = neg.purpose || '';
      E = (neg.detailAmount !== '' && neg.detailAmount != null) ? neg.detailAmount : '';
      matchCount++;
    } else if (dup){
      // 重複收款：E＝多收金額（正）
      E = dup.refund;
      matchCount++;
    } else if (xukouMap.has(id)){
      // 正金額虛扣對應（如月異常明細）
      const x = xukouMap.get(id);
      B = x.srcOrder || '';
      D = x.purpose || '';
      E = x.amount != null ? x.amount : '';
      matchCount++;
    } else if (neg && neg.status === '需確認'){
      D = '需確認';
    }
    if (dup) F = dup.times;                 // 重複收款次數（混合情況也標）
    if (manualSet.has(id)) G = '待人工修改'; // 待人工標記（樞紐用）
    const Eval = (E === '' || E == null) ? 0 : E;
    const H = Math.abs(C - Eval);            // 修改後實際金額＝|C-E|，不要負數
    aoa.push([id, B, C, D, E, F, G, H, I]);
  }
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [
    {wch:16}, {wch:16}, {wch:16}, {wch:22}, {wch:16}, {wch:12}, {wch:12}, {wch:16}, {wch:12},
  ];
  // 單號 / 原單號 設文字格式（避免變科學記號）
  for (let r=2; r<=aoa.length; r++){
    ['A','B'].forEach(col => {
      const cell = ws[col+r];
      if (cell && cell.v !== '' && cell.v != null) { cell.t = 's'; cell.v = String(cell.v); }
    });
  }
  ws['!autofilter'] = { ref:`A1:I${aoa.length}` };  // 開篩選，方便樞紐/篩 G、I 欄
  XLSX.utils.book_append_sheet(wb, ws, 'Final');

  // -- 虛扣對應 分頁（獨立呈現所有虛扣單號的對應狀態） --
  if (xukouMap.size){
    const mainSet = new Map(state.main.merged.map(g => [g.orderId, g.amount]));
    const xukouAoa = [['虛扣單號','原單號','調整目的','表內金額','主檔金額','入帳狀態']];
    // 先排已入帳，再排未入帳
    const matched = [], unmatched = [];
    for (const [xid, info] of xukouMap.entries()){
      const row = {
        xukouId: xid,
        srcOrder: info.srcOrder || '',
        purpose: info.purpose || '',
        amount: info.amount != null ? info.amount : '',
        mainAmount: mainSet.has(xid) ? Math.round(mainSet.get(xid)) : null,
      };
      if (mainSet.has(xid)) matched.push(row);
      else unmatched.push(row);
    }
    for (const r of matched){
      xukouAoa.push([r.xukouId, r.srcOrder, r.purpose, r.amount, r.mainAmount, '已入帳']);
    }
    for (const r of unmatched){
      xukouAoa.push([r.xukouId, r.srcOrder, r.purpose, r.amount, '未入帳', '未入帳']);
    }
    const ws3 = XLSX.utils.aoa_to_sheet(xukouAoa);
    ws3['!cols'] = [{wch:16},{wch:16},{wch:26},{wch:14},{wch:14},{wch:12}];
    for (let r=2; r<=xukouAoa.length; r++){
      ['A','B'].forEach(col => {
        const c = ws3[col+r];
        if (c && c.v !== '' && c.v != null) { c.t = 's'; c.v = String(c.v); }
      });
    }
    ws3['!autofilter'] = { ref: `A1:F${xukouAoa.length}` };
    XLSX.utils.book_append_sheet(wb, ws3, '虛扣對應');
  }

  // -- 負數單號彙整 分頁（所有 B 欄負數的對應狀態） --
  if (state.negSummary && state.negSummary.length){
    const negAoa = [['商家訂單 ID','合併後金額','對應狀態','原單號','調整目的','對應金額']];
    // 先排「需確認」，再貨故請賠，再虛扣調整（需確認排最前面，最該注意）
    const order = { '需確認':0, '貨故請賠':1, '虛扣調整':2 };
    const sorted = [...state.negSummary].sort((a,b)=>(order[a.status]??9)-(order[b.status]??9));
    for (const n of sorted){
      negAoa.push([n.order, n.mainAmount, n.status, n.srcOrder||'', n.purpose||'', n.detailAmount!==''?n.detailAmount:'']);
    }
    const wsN = XLSX.utils.aoa_to_sheet(negAoa);
    wsN['!cols'] = [{wch:16},{wch:14},{wch:12},{wch:16},{wch:24},{wch:14}];
    wsN['!autofilter'] = { ref: `A1:F${negAoa.length}` };
    for (let r=2; r<=negAoa.length; r++){
      ['A','D'].forEach(col => {
        const c = wsN[col+r];
        if (c && c.v !== '' && c.v != null) { c.t = 's'; c.v = String(c.v); }
      });
    }
    XLSX.utils.book_append_sheet(wb, wsN, '負數單號彙整');
  }

  // -- 待人工確認 分頁 --
  const manualAoa = [['單號','類型','主檔金額','請賠/應退','差額','說明']];
  for (const m of state.manual){
    manualAoa.push([
      String(m.order||''),
      m.type, m.mainAmount, m.claim, m.diff, m.note
    ]);
  }
  const ws2 = XLSX.utils.aoa_to_sheet(manualAoa);
  ws2['!cols'] = [{wch:24},{wch:16},{wch:14},{wch:14},{wch:10},{wch:36}];
  for (let r=2; r<=manualAoa.length; r++){
    const cell = ws2['A'+r];
    if (cell) { cell.t = 's'; cell.v = String(cell.v); }
  }
  XLSX.utils.book_append_sheet(wb, ws2, '待人工確認');

  // 檔名
  const today = new Date();
  const ymd = `${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}${String(today.getDate()).padStart(2,'0')}`;
  XLSX.writeFile(wb, `UBER帳單整理_${ymd}_v1.5.xlsx`);
  const parts = [`合併 ${state.main.merged.length.toLocaleString()} 筆`];
  if (state.adjustments.length) parts.push(`貨故 ${state.adjustments.length} 列`);
  if (matchCount > 0) parts.push(`虛扣對應 ${matchCount} 筆`);
  if (state.manual.length) parts.push(`待人工 ${state.manual.length} 筆`);
  toast(`已下載 Excel ｜ ${parts.join('，')}`);
}

// =============================================================
// 事件綁定
// =============================================================

function bindDrop(dropId, inputId, handler){
  const drop = document.getElementById(dropId);
  const input = document.getElementById(inputId);
  drop.addEventListener('click', e => {
    if (e.target.tagName === 'BUTTON') return;
    input.click();
  });
  drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('is-drag'); });
  drop.addEventListener('dragleave', e => { drop.classList.remove('is-drag'); });
  drop.addEventListener('drop', e => {
    e.preventDefault(); drop.classList.remove('is-drag');
    if (e.dataTransfer.files.length) handler(e.dataTransfer.files[0]);
  });
  input.addEventListener('change', e => {
    if (input.files.length) handler(input.files[0]);
  });
}

document.addEventListener('DOMContentLoaded', ()=>{
  // 版本標示：由 app.js 設定，可確認 app.js 是否為新版
  const APP_VERSION = 'v1.5 · 06/05 ✓';
  const verChip = document.getElementById('verChip');
  if (verChip) verChip.textContent = APP_VERSION;
  console.log('[UBER 帳單整理工具] app.js 版本：' + APP_VERSION + '（含負金額待查、虛扣對應分頁）');

  // pickers
  $$('[data-pick]').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      document.getElementById(btn.dataset.pick).click();
    });
  });

  bindDrop('dropMain',  'fileMain',  loadMain);
  bindDrop('dropHuogu', 'fileHuogu', loadHuogu);
  bindDrop('dropXukou', 'fileXukou', loadXukou);

  $('#btnPreview').addEventListener('click', previewMerge);
  $('#btnExport').addEventListener('click', exportExcel);
  $('#btnReset').addEventListener('click', ()=>{
    if (!confirm('確定要重新開始？所有上傳的資料都會清除。')) return;
    location.reload();
  });

  // tabs
  $$('.tab').forEach(t => {
    t.addEventListener('click', ()=>{
      $$('.tab').forEach(x => x.classList.remove('active'));
      $$('.tab-pane').forEach(x => x.classList.remove('active'));
      t.classList.add('active');
      const map = {left:'paneLeft', right:'paneRight', xukou:'paneXukou', manual:'paneManual'};
      $('#'+map[t.dataset.tab]).classList.add('active');
    });
  });
});

})();
