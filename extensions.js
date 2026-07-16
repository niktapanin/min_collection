// ============================================================
// EXTENSIONS.JS v4 — модуль расширений каталога минералов
// Подключение: <script src="extensions.js?v=4"></script> перед </body>
//
// v4 — принципиально новая карточка и структура системы:
//  • карточка пересобрана в секции: Систематика / Происхождение / Химия /
//    Изоморфный ряд / Минералы-спутники / Хранение / Учёт
//  • убраны дубли: символ IMA, английская сингония, английский генезис
//  • «Развитие» и «Раскладка» — отдельные разделы (кнопки под сводкой),
//    а не вкладки среди графиков аналитики
//  • спутники: суффиксы видов больше не теряются; статус учитывает
//    «семейство» (monazite засчитывается, если есть monazite-(Ce))
// ============================================================
(function(){
'use strict';

console.log('%c[ext] extensions.js v4 загружен', 'color:#c8a96e;font-weight:bold');

const EXT = {
  ref: new Map(),          // ключ вида -> справочная запись
  owned: new Set(),        // основные виды коллекции
  sat: new Set(),          // виды-спутники составных образцов
  slugMap: new Map(),
  refReady: false,
  hashGuard: false
};
window.EXT = EXT;

// ---------- нормализация имён ----------
// Суффикс «-(Ce)» не обрубается, а приводится к «-ce»: Agardite-(Ce)/(Y)/(La) — разные виды.
function baseName(n){
  n = (n||'').toString().trim().toLowerCase();
  try { n = n.normalize('NFD').replace(/[\u0300-\u036f]/g,''); } catch(e){}
  n = n.replace(/-\(([a-z0-9+]+)\)$/, '-$1');
  return n.replace('baryte','barite');
}
function keyRoot(n){ return baseName(n).replace(/-[a-z0-9+]+$/,''); }
function ruCap(s){ return s ? s[0].toUpperCase()+s.slice(1) : s; }

const SYNGONY_RU = { 'triclinic':'триклинная','monoclinic':'моноклинная','orthorhombic':'ромбическая',
  'tetragonal':'тетрагональная','trigonal':'тригональная','hexagonal':'гексагональная',
  'isometric':'кубическая','cubic':'кубическая','amorphous':'аморфный','icosahedral':'икосаэдрическая' };

const FLAG_META = {
  light_decompose: {icon:'☀', label:'Светочувствителен', note:'Разлагается на свету — вглубь полки, от прямого света', color:'#b8862c'},
  light_fade:      {icon:'◐', label:'Выцветает', note:'Окраска может выцветать — не под прямым светом', color:'#9a8f7e'},
  hygroscopic:     {icon:'💧', label:'Гигроскопичен', note:'Закрытый бокс с силикагелем', color:'#4a7fb8'},
  dehydration:     {icon:'△', label:'Дегидратация', note:'Не пересушивать и не нагревать; стабильная влажность', color:'#7a9e6a'},
  sulfide_decay:   {icon:'⚠', label:'Сульфидный распад', note:'Следить: при белёсом налёте или запахе серы — изолировать в сухой бокс', color:'#c06a2c'},
  radioactive:     {icon:'☢', label:'Радиоактивен', note:'Дистанция от рабочих мест, закрытый бокс', color:'#b83a3a'},
  handle_care:     {icon:'✋', label:'Осторожно', note:'Осторожность при контакте (волокна/токсичность)', color:'#8a5a9e'}
};

// ---------- статус вида относительно коллекции ----------
// 'owned' | 'sat' | 'family' (нет точного вида, но есть представитель семейства) | 'none'
function statusOf(name){
  if (EXT.owned.has(name)) return 'owned';
  if (EXT.sat.has(name)) return 'sat';
  const pref = name + '-';
  for (const k of EXT.owned) if (k.startsWith(pref)) return 'family';
  for (const k of EXT.sat)   if (k.startsWith(pref)) return 'family';
  return 'none';
}
function specimenByBase(b){
  let hit = state.data.find(d => baseName(d.ima_name) === b);
  if (!hit) hit = state.data.find(d => baseName(d.ima_name).startsWith(b + '-'));
  return hit;
}
function activeFlags(x){
  const r = EXT.ref.get(baseName(x.ima_name));
  if (!r) return [];
  const over = new Set(Array.isArray(x.flag_override) ? x.flag_override : []);
  return (r.cons_flags||[]).filter(f => !over.has(f));
}
async function patchMineral(id, body){
  return fetch(`${SB_URL}/rest/v1/minerals?id=eq.${encodeURIComponent(id)}`, {
    method:'PATCH', headers:{...SB_HEADERS,'Content-Type':'application/json','Prefer':'return=minimal'},
    body: JSON.stringify(body)
  });
}

// ---------- стили v4 ----------
function injectCSS(){
  if (el('extStyles')) return;
  const st = document.createElement('style');
  st.id = 'extStyles';
  st.textContent = `
  .ext-sec { margin-bottom: 18px; }
  .ext-sec-title {
    font-family:'Syne',sans-serif; font-size:0.72rem; letter-spacing:0.12em;
    text-transform:uppercase; color:#9a8f7e; margin-bottom:8px;
    display:flex; align-items:center; gap:10px;
  }
  .ext-sec-title::after { content:''; flex:1; height:1px; background:#ece7de; }
  .ext-kv { display:flex; gap:10px; padding:4px 0; font-size:0.9rem; align-items:baseline; }
  .ext-kv .k { color:#9a8f7e; font-size:0.78rem; width:118px; flex-shrink:0; }
  .ext-kv .v { color:var(--stone); }
  .ext-chip { font-size:0.78rem; padding:5px 11px; border-radius:8px; border:1px solid #ddd8d0;
    background:#f5f2ed; color:#8a8070; display:inline-flex; align-items:center; gap:5px; }
  .ext-chip.owned { background:#e6f2ea; border-color:#bcd8c4; color:#2c5c3c; cursor:pointer; }
  .ext-chip.owned:hover { background:#d8ecdf; }
  .ext-chip.family { background:#eef4ec; border-color:#ccdcc8; color:#4a6c50; cursor:pointer; }
  .ext-chip.sat { background:#fdf6e3; border-color:#e8d9a8; color:#8a6d1a; }
  .ext-chip.plain { background:transparent; border-style:dashed; opacity:0.65; }
  .ext-series { display:flex; gap:8px; overflow-x:auto; padding:4px 2px 8px; -webkit-overflow-scrolling:touch; }
  .ext-series-item { min-width:92px; flex-shrink:0; border-radius:12px; border:1.5px solid #e2ddd5;
    background:#fff; padding:10px 10px 8px; text-align:center; }
  .ext-series-item .subst { font-family:ui-monospace,monospace; font-size:1rem; font-weight:600; color:var(--stone); }
  .ext-series-item .nm { font-size:0.68rem; color:#8a8070; margin-top:3px; line-height:1.2; word-break:break-word; }
  .ext-series-item .st { font-size:0.6rem; margin-top:5px; padding:2px 7px; border-radius:9px; display:inline-block; }
  .ext-series-item.cur { border-color:var(--amber); background:#fdf9f0; box-shadow:0 2px 8px rgba(200,169,110,0.18); }
  .ext-series-item.owned { border-color:#bcd8c4; cursor:pointer; }
  .ext-series-item.owned:hover { background:#f2f9f4; }
  .ext-nav { display:grid; grid-template-columns:1fr 1fr; gap:12px; margin-bottom:20px; }
  .ext-nav-card { border:1px solid #e2ddd5; border-radius:16px; background:linear-gradient(135deg,#fff, #faf7f1);
    padding:16px 18px; cursor:pointer; transition:all .15s; display:flex; flex-direction:column; gap:4px; }
  .ext-nav-card:hover { transform:translateY(-1px); box-shadow:0 6px 22px rgba(0,0,0,0.08); border-color:#d4c9b4; }
  .ext-nav-card .t { font-family:'Syne',sans-serif; font-size:0.95rem; color:var(--stone); display:flex; align-items:center; gap:8px; }
  .ext-nav-card .d { font-size:0.74rem; color:#9a8f7e; line-height:1.35; }
  @media (max-width:640px){ .ext-nav { grid-template-columns:1fr; gap:8px; } .ext-nav-card{ padding:13px 15px; } }
  .ext-view-bg { position:fixed; inset:0; z-index:60; background:rgba(26,23,20,0.6); backdrop-filter:blur(4px);
    display:none; align-items:center; justify-content:center; padding:16px; }
  .ext-view { background:#f5f2ed; width:100%; max-width:860px; max-height:92dvh; border-radius:20px;
    overflow-y:auto; -webkit-overflow-scrolling:touch; }
  .ext-view-head { position:sticky; top:0; z-index:2; background:var(--stone); padding:16px 22px;
    display:flex; align-items:center; justify-content:space-between; border-radius:20px 20px 0 0; }
  .ext-view-head .t { font-family:'Syne',sans-serif; color:#fff; font-size:1.1rem; }
  .ext-view-head .sub { color:rgba(232,228,218,0.55); font-size:0.72rem; margin-top:2px; }
  @media (max-width:640px){ .ext-view-bg{padding:0; align-items:flex-end;} .ext-view{border-radius:18px 18px 0 0; max-height:96dvh;} }
  `;
  document.head.appendChild(st);
}

// ---------- загрузка справочника (статический файл) ----------
async function loadReference(){
  try {
    const res = await fetch('species_reference.json', { cache: 'no-cache' });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const data = await res.json();
    Object.entries(data).forEach(([name, r]) => EXT.ref.set(name, { name, ...r }));
    EXT.refReady = true;
    console.log('[ext] справочник загружен:', EXT.ref.size, 'видов');
  } catch(e){
    console.error('[ext] species_reference.json не загрузился:', e.message);
  }
}
function computeOwnedSets(){
  EXT.owned.clear(); EXT.sat.clear();
  const abbrMap = new Map();
  state.data.forEach(d => {
    const n = (d.ima_name||'').trim();
    if (n && n!=='0') {
      EXT.owned.add(baseName(n));
      const a = (d.abbr||'').trim();
      if (a && a!=='0' && !abbrMap.has(a)) abbrMap.set(a, baseName(n));
    }
  });
  Object.entries({'Nph':'nepheline','Kln':'kaolinite','Heu':'heulandite-ca','Cpy':'chalcopyrite',
    'Cc':'chalcocite','Ilt':'illite','Lö':'lollingite','Gp-S':'gypsum','Ol-Fo':'forsterite',
    'Fo?':'forsterite','Cr-Di':'diopside','Ms-Fu':'muscovite','Mn-Cal':'calcite','Ckt':'cookeite','Nc':'nacrite'
  }).forEach(([a,b]) => abbrMap.set(a,b));
  state.data.forEach(d => {
    const groups = (d.id||'').match(/\(([^)]+)\)/g) || [];
    groups.slice(1).forEach(g => {
      g.slice(1,-1).split(/[,;]\s*/).forEach(a => {
        const b = abbrMap.get(a.trim());
        if (b && !EXT.owned.has(b)) EXT.sat.add(b);
      });
    });
  });
}
async function loadExtraColumns(){
  let offset = 0, limit = 1000;
  while (true) {
    const res = await fetch(`${SB_URL}/rest/v1/minerals?select=id,shelf,shelf_pos,acquired_at,flag_override&order=id&limit=${limit}&offset=${offset}`, { headers: SB_HEADERS });
    if (!res.ok) return;
    const rows = await res.json();
    rows.forEach(r => {
      const d = state.data.find(x => x.id === r.id);
      if (d) Object.assign(d, { shelf: r.shelf, shelf_pos: r.shelf_pos, acquired_at: r.acquired_at, flag_override: r.flag_override || [] });
    });
    if (rows.length < limit) break;
    offset += limit;
  }
}

// ---------- РОУТИНГ ----------
function slugOf(id){
  const prefix = id.split(' (')[0].trim().replace(/\s+/g,'_');
  const groups = (id.match(/\(([^)]+)\)/g)||[]).map(g => g.slice(1,-1).split(/[,;]\s*/).map(s=>s.trim()).join('.'));
  return groups.length ? `${prefix}/${groups.join('+')}` : prefix;
}
function buildSlugMap(){
  EXT.slugMap.clear();
  state.data.forEach(d => EXT.slugMap.set(slugOf(d.id), d));
}
function setHash(h){
  EXT.hashGuard = true;
  location.hash = h;
  setTimeout(()=>{ EXT.hashGuard = false; }, 60);
}
function clearHash(){
  if (/^#\/(mineral|development|layout)\//.test(location.hash) || /^#\/(development|layout)$/.test(location.hash) || location.hash.startsWith('#/mineral/')) {
    EXT.hashGuard = true;
    history.replaceState(null,'', location.pathname + location.search + '#/');
    setTimeout(()=>{ EXT.hashGuard = false; }, 60);
  }
}
function handleRoute(){
  if (EXT.hashGuard) return;
  const h = decodeURIComponent(location.hash||'');
  let m;
  if ((m = h.match(/^#\/mineral\/(.+)$/))) {
    const obj = EXT.slugMap.get(m[1]);
    if (obj) openDetailModal(obj);
  } else if ((m = h.match(/^#\/shelf\/(.+)$/))) {
    openShelfView(m[1]);
  } else if (h === '#/development') {
    openExtView('coverage');
  } else if (h === '#/layout') {
    openExtView('order');
  }
}

// ============================================================
// НОВАЯ КАРТОЧКА — секции вместо плоского списка
// ============================================================
function secHtml(title, inner){
  return inner ? `<div class="ext-sec"><div class="ext-sec-title">${title}</div>${inner}</div>` : '';
}
function kv(k, v){ return v ? `<div class="ext-kv"><span class="k">${k}</span><span class="v">${v}</span></div>` : ''; }

function chip(name){
  const r = EXT.ref.get(name);
  const label = r ? r.name_display : ruCap(name);
  const st = statusOf(name);
  if (st === 'owned' || st === 'family')
    return `<span class="ext-chip ${st} ext-open" data-b="${name}" title="${st==='owned'?'есть в коллекции':'есть представитель семейства'} — открыть">${escapeHtml(label)}</span>`;
  if (st === 'sat')
    return `<span class="ext-chip sat" title="есть как вид-спутник составного образца">${escapeHtml(label)}</span>`;
  if (!r) return `<span class="ext-chip plain">${escapeHtml(label)}</span>`;
  return `<span class="ext-chip" title="нет в коллекции">${escapeHtml(label)}</span>`;
}

function seriesHtml(x, r){
  const an = (r && r.analogs) || [];
  if (!an.length) return '';
  const cur = `<div class="ext-series-item cur">
      <div class="subst">●</div>
      <div class="nm">${escapeHtml(r.name_display||x.ima_name)}</div>
      <div class="st" style="background:#f3e8cf;color:#8a6d1a">этот вид</div>
    </div>`;
  const items = an.slice(0, 14).map(([other, myEl, otherEl]) => {
    const rr = EXT.ref.get(other);
    const st = statusOf(other);
    const cls = (st==='owned'||st==='family') ? 'owned' : '';
    const badge = st==='owned' ? ['есть','#e6f2ea','#2c5c3c'] : st==='family' ? ['семейство','#eef4ec','#4a6c50']
                : st==='sat' ? ['спутник','#fdf6e3','#8a6d1a'] : ['нет','#f0ede8','#9a8f7e'];
    return `<div class="ext-series-item ${cls}" ${cls?`data-b="${other}"`:''} title="${rr&&rr.formula_html?rr.formula_html.replace(/<[^>]+>/g,''):''}">
      <div class="subst">${myEl}→${otherEl}</div>
      <div class="nm">${escapeHtml(rr ? rr.name_display : ruCap(other))}</div>
      <div class="st" style="background:${badge[1]};color:${badge[2]}">${badge[0]}</div>
    </div>`;
  }).join('');
  const more = an.length > 14 ? `<div class="ext-series-item" style="opacity:0.6"><div class="subst">…</div><div class="nm">+${an.length-14}</div></div>` : '';
  return `<div class="ext-series">${cur}${items}${more}</div>
    <div style="font-size:0.68rem;color:#b8b0a0;margin-top:2px">Та же формула, один элемент замещён. Зелёная рамка — есть в коллекции, тап открывает образец.</div>`;
}

function groupMembersHtml(x, r){
  if (!r || !r.struct_group) return '';
  const b0 = baseName(x.ima_name);
  const members = [];
  const seen = new Set();
  state.data.forEach(d => {
    const b = baseName(d.ima_name);
    if (!b || b === b0 || seen.has(b)) return;
    const rr = EXT.ref.get(b);
    if (rr && rr.struct_group === r.struct_group) { members.push(b); seen.add(b); }
  });
  if (!members.length) return '';
  return `<div class="ext-kv" style="align-items:flex-start"><span class="k">В коллекции из группы</span>
    <span class="v" style="display:flex;flex-wrap:wrap;gap:5px">${members.slice(0,12).map(chip).join('')}${members.length>12?`<span class="ext-chip plain">+${members.length-12}</span>`:''}</span></div>`;
}

function buildCard(x){
  const b = baseName(x.ima_name);
  const r = (x.ima_name && x.ima_name!=='0') ? EXT.ref.get(b) : null;
  const mBody = el('mBody');

  const syng = x.syngony || (r && SYNGONY_RU[r.crystal_system]) || '';
  const cost = (x.cost != null && !isNaN(x.cost)) ? (state.costUnlocked ? fmtMoney(x.cost)+' ₽' : '••••••') : '—';
  const updated = x.updated_at ? new Date(x.updated_at).toLocaleString('ru-RU',{dateStyle:'medium',timeStyle:'short'}) : '';

  // --- Систематика
  const sysHtml = kv('Класс', x.class ? escapeHtml(x.class) : '')
    + kv('Структурная группа', r && r.struct_group ? escapeHtml(r.struct_group) : '')
    + kv('Сингония', syng ? escapeHtml(syng) : '');

  // --- Происхождение
  const originHtml = kv('Месторождение', x.locality ? `${x.is_type_locality?'<span class="type-star" title="Типовое месторождение">★</span> ':''}${escapeHtml(x.locality)}` : '')
    + kv('Страна', x.country ? `${countryFlag(x.country)} ${escapeHtml(x.country)}` : '')
    + kv('Год открытия', x.discovery_year || (r && r.year_published) || '')
    + (r && r.etymology ? `<div style="font-size:0.82rem;color:#6a6050;font-style:italic;line-height:1.5;margin-top:6px">${escapeHtml(r.etymology)}</div>` : '');

  // --- Химия
  const els = normalizeElements(x.elements || []);
  const chemHtml = kv('Формула', (x.formula_html || x.formula_text) ? `<span class="formula">${x.formula_html || escapeHtml(x.formula_text)}</span>` : '')
    + (els.length ? `<div class="ext-kv" style="align-items:flex-start"><span class="k">Элементы</span><span class="v" style="display:flex;flex-wrap:wrap;gap:5px">${els.map(e=>`<span class="ext-chip">${e}</span>`).join('')}</span></div>` : '');

  // --- Связи: изоморфный ряд + группа
  const series = r ? seriesHtml(x, r) : '';
  const groupM = r ? groupMembersHtml(x, r) : '';
  const linksHtml = (series || groupM) ? series + groupM : '';

  // --- Минералы-спутники
  const satHtml = (r && r.assoc_parsed && r.assoc_parsed.length)
    ? `<div style="display:flex;flex-wrap:wrap;gap:5px">${r.assoc_parsed.map(chip).join('')}</div>
       <div style="font-size:0.68rem;color:#b8b0a0;margin-top:6px">С кем этот вид встречается в природе. Зелёные — есть в коллекции.</div>`
    : '';

  // --- Хранение
  const over = new Set(Array.isArray(x.flag_override) ? x.flag_override : []);
  const allFlags = r ? (r.cons_flags||[]) : [];
  let careHtml = '';
  allFlags.forEach(f => {
    const m = FLAG_META[f]; if (!m) return;
    const off = over.has(f);
    careHtml += `<div style="display:flex;align-items:flex-start;gap:8px;padding:8px 10px;border-radius:10px;background:${off?'#f5f2ed':'#fdf9f2'};border:1px solid ${off?'#e2ddd5':m.color+'40'};${off?'opacity:0.55':''};margin-bottom:6px">
      <span style="color:${m.color};flex-shrink:0">${m.icon}</span>
      <span style="flex:1;font-size:0.8rem;color:var(--stone)"><b>${m.label}.</b> ${m.note}${off?' <i style="color:#9a8f7e">(флаг снят)</i>':''}</span>
      ${state.costUnlocked ? `<button class="ext-flag-toggle" data-flag="${f}" style="font-size:0.7rem;color:#8a8070;background:none;border:none;cursor:pointer;text-decoration:underline;flex-shrink:0">${off?'вернуть':'снять'}</button>` : ''}
    </div>`;
  });

  // --- Учёт
  let acctHtml = kv('Стоимость', `<span style="font-weight:600">${cost}</span>`)
    + (x.shelf ? kv('Полка', `<a href="#/shelf/${encodeURIComponent(x.shelf)}" style="text-decoration:underline">${escapeHtml(x.shelf)}${x.shelf_pos?` · поз. ${x.shelf_pos}`:''}</a>`) : '');
  if (state.costUnlocked) {
    acctHtml += `<div class="ext-kv"><span class="k">Дата поступления</span><span class="v" style="display:flex;gap:8px;align-items:center">
      <input type="date" id="extAcq" value="${x.acquired_at||''}" style="border:1.5px solid #d6d0c6;border-radius:8px;padding:4px 8px;font-size:0.82rem"/>
      <button id="extAcqSave" style="font-size:0.72rem;color:#8a8070;background:none;border:none;cursor:pointer;text-decoration:underline">сохранить</button></span></div>`;
    acctHtml += `<div class="ext-kv"><span class="k">Типовое м-ние</span><span class="v">
      <button id="extTlToggle" style="display:inline-flex;align-items:center;gap:8px;background:none;border:none;cursor:pointer;padding:0">
        <span style="width:34px;height:19px;border-radius:10px;background:${x.is_type_locality?'var(--amber)':'#d1cdc7'};position:relative;display:inline-block;transition:background .2s">
          <span style="width:15px;height:15px;border-radius:50%;background:#fff;position:absolute;top:2px;left:${x.is_type_locality?'17px':'2px'};transition:left .2s;box-shadow:0 1px 3px rgba(0,0,0,0.2)"></span>
        </span>
        <span style="font-size:0.82rem;color:${x.is_type_locality?'var(--stone)':'#9a8f7e'}">${x.is_type_locality?'★ да':'нет'}</span>
      </button></span></div>`;
  } else if (x.acquired_at) {
    acctHtml += kv('Дата поступления', new Date(x.acquired_at).toLocaleDateString('ru-RU'));
  }

  mBody.innerHTML = `
    <div id="extPhoto"></div>
    ${secHtml('Систематика', sysHtml)}
    ${secHtml('Происхождение', originHtml)}
    ${secHtml('Химия', chemHtml)}
    ${secHtml('Изоморфный ряд', linksHtml)}
    ${secHtml('Минералы-спутники', satHtml)}
    ${secHtml('Хранение', careHtml)}
    ${secHtml('Учёт', acctHtml)}
    ${updated ? `<div style="font-size:0.68rem;color:#c8bfaf;text-align:right;margin-top:4px">изменено ${updated}</div>` : ''}
    ${(!r && x.ima_name && x.ima_name!=='0') ? `<div style="font-size:0.72rem;color:#b8a06a;margin-top:8px">Справочных данных для «${escapeHtml(x.ima_name)}» нет (вид вне списка IMA/справочника).</div>` : ''}
  `;

  // фото — родная функция каталога
  if (x.id) renderPhotoSection(x, el('extPhoto'));

  // обработчики
  mBody.querySelectorAll('.ext-open[data-b], .ext-series-item[data-b]').forEach(c => {
    c.addEventListener('click', () => {
      const obj = specimenByBase(c.dataset.b);
      if (obj) { closeDetailModal(); setTimeout(()=>openDetailModal(obj), 80); }
    });
  });
  mBody.querySelectorAll('.ext-flag-toggle').forEach(btn => {
    btn.onclick = async () => {
      const f = btn.dataset.flag;
      let ov = Array.isArray(x.flag_override) ? [...x.flag_override] : [];
      ov = ov.includes(f) ? ov.filter(z=>z!==f) : [...ov, f];
      x.flag_override = ov;
      const rec = state.data.find(d=>d.id===x.id); if (rec) rec.flag_override = ov;
      await patchMineral(x.id, {flag_override: ov});
      buildCard(x); decorateCards(true);
    };
  });
  const acq = el('extAcqSave');
  if (acq) acq.onclick = async () => {
    const v = el('extAcq').value || null;
    x.acquired_at = v;
    const rec = state.data.find(d=>d.id===x.id); if (rec) rec.acquired_at = v;
    await patchMineral(x.id, {acquired_at: v});
    acq.textContent = '✓'; setTimeout(()=>acq.textContent='сохранить', 1200);
  };
  const tl = el('extTlToggle');
  if (tl) tl.onclick = async () => {
    const nv = !x.is_type_locality;
    x.is_type_locality = nv;
    const rec = state.data.find(d=>d.id===x.id); if (rec) rec.is_type_locality = nv;
    await patchMineral(x.id, {is_type_locality: nv});
    buildCard(x);
    if (typeof rerenderCardInGrid === 'function') rerenderCardInGrid(x);
  };
}

// ============================================================
// РАЗДЕЛЫ «РАЗВИТИЕ» И «РАСКЛАДКА» — полноэкранные, вход с главной
// ============================================================
function renderCoverage(list, body){
  if (!EXT.refReady) { body.innerHTML = '<div class="text-sm text-stone-400 text-center py-6">Справочник ещё загружается…</div>'; return; }
  const frame = new Map();
  EXT.ref.forEach(r => {
    if (!r.canonical_class) return;
    if (!frame.has(r.canonical_class)) frame.set(r.canonical_class, []);
    const st = statusOf(r.name);
    frame.get(r.canonical_class).push({name:r.name, disp:r.name_display, st});
  });
  const rows = [...frame.entries()].map(([cls, arr]) => {
    const have = arr.filter(a=>a.st==='owned'||a.st==='family').length;
    return {cls, arr, have, total: arr.length, pct: Math.round(100*have/arr.length)};
  }).sort((a,b)=>a.pct-b.pct);

  let html = `<div class="text-xs text-stone-500 mb-3">Эталонный каркас: классические представители каждого класса. Клик по строке раскрывает дыры.</div>`;
  rows.forEach((r,i) => {
    const missing = r.arr.filter(a=>a.st==='none');
    html += `<div class="bar-row ext-cov-row">
      <div class="flex items-center justify-between gap-3 mb-1">
        <div class="text-sm font-medium text-stone-800">${r.cls}</div>
        <div class="mono text-xs text-stone-500">${r.have}/${r.total} · ${r.pct}%</div>
      </div>
      <div class="h-2 rounded-full bg-stone-100 overflow-hidden"><div class="bar-fill" style="width:${Math.max(4,r.pct)}%;background:${r.pct<50?'#c06a2c':r.pct<80?'#c8a96e':'#4fad7f'};opacity:0.8"></div></div>
      <div class="ext-cov-detail hidden" style="margin-top:6px;font-size:0.78rem;color:#6a6050">
        ${missing.length ? 'Не хватает: ' + missing.map(m=>`<a href="https://www.mindat.org/search.php?search=${encodeURIComponent(m.disp||m.name)}" target="_blank" rel="noopener" style="text-decoration:underline dotted">${m.disp||m.name}</a>`).join(', ') : 'Каркас закрыт ✓'}
      </div></div>`;
  });

  const linkCnt = new Map(), via = new Map();
  EXT.owned.forEach(b => {
    const r = EXT.ref.get(b); if (!r) return;
    (r.assoc_parsed||[]).forEach(a => {
      if (statusOf(a)!=='none' || !EXT.ref.has(a)) return;
      linkCnt.set(a,(linkCnt.get(a)||0)+1);
      if (!via.has(a)) via.set(a,new Set());
      if (via.get(a).size<5) via.get(a).add(EXT.ref.get(b)?.name_display||b);
    });
  });
  const wl = [...linkCnt.entries()]
    .map(([n,c]) => ({n, c, common: EXT.ref.get(n)?.assoc_count||0, disp: EXT.ref.get(n)?.name_display||n}))
    .sort((a,b)=> (b.c*2+Math.min(b.common,50)*0.2) - (a.c*2+Math.min(a.common,50)*0.2)).slice(0,20);
  html += `<div class="text-sm mt-6 mb-2 font-semibold" style="font-family:'Syne',sans-serif;color:var(--stone)">Отсутствующие виды-спутники</div>`;
  wl.forEach(w => {
    html += `<div style="display:flex;align-items:baseline;gap:8px;padding:5px 8px;font-size:0.82rem;border-bottom:1px solid #ece7de">
      <a href="https://www.mindat.org/search.php?search=${encodeURIComponent(w.disp)}" target="_blank" rel="noopener" style="font-weight:500;color:var(--stone);text-decoration:underline dotted">${w.disp}</a>
      <span class="mono" style="font-size:0.7rem;color:#8a8070">${w.c} связ.</span>
      <span style="font-size:0.7rem;color:#b8b0a0;flex:1;text-align:right" class="truncate">спутник: ${[...(via.get(w.n)||[])].slice(0,3).join(', ')}</span></div>`;
  });

  const anCnt = new Map(), anVia = new Map();
  EXT.owned.forEach(b => {
    const r = EXT.ref.get(b); if (!r) return;
    (r.analogs||[]).forEach(([other,e1,e2]) => {
      if (statusOf(other)!=='none') return;
      anCnt.set(other,(anCnt.get(other)||0)+1);
      if (!anVia.has(other)) anVia.set(other, `${r.name_display||b} (${e1}→${e2})`);
    });
  });
  const an = [...anCnt.entries()].sort((a,b)=>b[1]-a[1]).slice(0,20);
  html += `<div class="text-sm mt-6 mb-2 font-semibold" style="font-family:'Syne',sans-serif;color:var(--stone)">Отсутствующие химические аналоги</div>`;
  an.forEach(([n,c]) => {
    const rr = EXT.ref.get(n);
    html += `<div style="display:flex;align-items:baseline;gap:8px;padding:5px 8px;font-size:0.82rem;border-bottom:1px solid #ece7de">
      <a href="https://www.mindat.org/search.php?search=${encodeURIComponent(rr?.name_display||n)}" target="_blank" rel="noopener" style="font-weight:500;color:var(--stone);text-decoration:underline dotted">${rr?.name_display||n}</a>
      ${rr?.formula_html ? `<span class="formula" style="color:#8a8070;font-size:0.72rem">${rr.formula_html}</span>` : ''}
      <span class="mono" style="font-size:0.7rem;color:#8a8070">${c} вид.</span>
      <span style="font-size:0.7rem;color:#b8b0a0;flex:1;text-align:right" class="truncate">${anVia.get(n)||''}</span></div>`;
  });

  body.innerHTML = html;
  body.querySelectorAll('.ext-cov-row').forEach(row =>
    row.addEventListener('click', () => row.querySelector('.ext-cov-detail')?.classList.toggle('hidden')));
}

function assocLinked(b1,b2){
  if (!b1 || !b2) return false;
  const r1 = EXT.ref.get(b1), r2 = EXT.ref.get(b2);
  if (r1 && (r1.assoc_parsed||[]).includes(b2)) return true;
  if (r2 && (r2.assoc_parsed||[]).includes(b1)) return true;
  if (r1 && (r1.analogs||[]).some(a=>a[0]===b2)) return true;
  return false;
}
function greedyChain(items){
  if (items.length<=2) return items;
  const bOf = d => (d.ima_name && d.ima_name!=='0') ? baseName(d.ima_name) : '';
  const score = (a,b2) => { const x=bOf(a), y=bOf(b2); if(!x||!y) return 0; if(x===y) return 3; return assocLinked(x,y)?2:0; };
  const deg = d => items.reduce((s,o)=> s + (o!==d && score(d,o)>0 ? 1:0), 0);
  let remaining=[...items].sort((a,b)=> deg(b)-deg(a) || (a.name_ru||'').localeCompare(b.name_ru||'','ru'));
  const chain=[remaining.shift()];
  while (remaining.length){
    const last=chain[chain.length-1];
    let best=remaining[0], bs=-1;
    remaining.forEach(d=>{ const s=score(last,d); if (s>bs) { bs=s; best=d; } });
    chain.push(best); remaining=remaining.filter(d=>d!==best);
  }
  return chain;
}
function renderOrder(list, body){
  if (!EXT.refReady) { body.innerHTML = '<div class="text-sm text-stone-400 text-center py-6">Справочник ещё загружается…</div>'; return; }
  const byClass = new Map();
  list.forEach(d => { const c=d.class||'без класса'; if(!byClass.has(c)) byClass.set(c,[]); byClass.get(c).push(d); });
  const classes=[...byClass.keys()].sort((a,b)=>(classSortKey(a)-classSortKey(b)));
  let html = `<div class="text-xs text-stone-500 mb-3">Порядок: класс → структурная группа → цепочка природных соседств (спутники + изоморфизм). Зелёное тире между образцами = природное соседство. <button id="extOrderCsv" style="text-decoration:underline;color:var(--stone);background:none;border:none;cursor:pointer;font-size:inherit">Скачать CSV</button></div>`;
  classes.forEach(cls => {
    const items=byClass.get(cls);
    const groups=new Map();
    items.forEach(d=>{
      const b=(d.ima_name&&d.ima_name!=='0')?baseName(d.ima_name):'';
      const g=(b && EXT.ref.get(b)?.struct_group) || '— вне групп —';
      if(!groups.has(g)) groups.set(g,[]);
      groups.get(g).push(d);
    });
    const gnames=[...groups.keys()].filter(g=>g!=='— вне групп —').sort((a,b)=>groups.get(b).length-groups.get(a).length);
    if (groups.has('— вне групп —')) gnames.push('— вне групп —');
    html += `<div style="margin-bottom:16px"><div style="font-family:'Syne',sans-serif;font-size:0.92rem;color:var(--stone);margin-bottom:6px;border-bottom:2px solid ${cardAccent(cls)}40;padding-bottom:3px">${cls} <span class="mono" style="font-size:0.7rem;color:#9a8f7e">${items.length}</span></div>`;
    gnames.forEach(g=>{
      const chain=greedyChain(groups.get(g));
      html += `<div style="margin-bottom:7px"><div style="font-size:0.7rem;color:#8a8070;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:3px">${g} · ${chain.length}</div><div style="display:flex;flex-wrap:wrap;gap:4px;align-items:center">`;
      chain.forEach((d,i)=>{
        const prev = i>0 ? chain[i-1] : null;
        const linked = prev && assocLinked(baseName(d.ima_name), baseName(prev.ima_name));
        const flags = activeFlags(d).map(f=>FLAG_META[f]?FLAG_META[f].icon:'').join('');
        html += `${i>0?`<span style="color:${linked?'#4fad7f':'#ddd8d0'};font-size:0.7rem">${linked?'—':'·'}</span>`:''}<span class="ext-chip ext-order-chip" data-id="${escapeHtml(d.id)}" style="background:${cardAccent(d.class)}12;border-color:${cardAccent(d.class)}35;color:var(--stone);cursor:pointer" title="${escapeHtml(d.locality||'')}">${d.is_type_locality?'★ ':''}${escapeHtml(d.name_ru||d.id)}${flags?' '+flags:''}</span>`;
      });
      html += `</div></div>`;
    });
    html += `</div>`;
  });
  body.innerHTML = html;
  body.querySelectorAll('.ext-order-chip').forEach(c => c.addEventListener('click', () => {
    const d = state.data.find(z=>z.id===c.dataset.id); if (d) openDetailModal(d);
  }));
  const csvBtn = el('extOrderCsv');
  if (csvBtn) csvBtn.onclick = () => {
    let csv = 'класс;группа;позиция;ID;название;типовое;хранение\n';
    classes.forEach(cls => {
      const items=byClass.get(cls); const groups=new Map();
      items.forEach(d=>{ const b=(d.ima_name&&d.ima_name!=='0')?baseName(d.ima_name):''; const g=(b&&EXT.ref.get(b)?.struct_group)||'— вне групп —'; if(!groups.has(g)) groups.set(g,[]); groups.get(g).push(d); });
      const gn=[...groups.keys()].filter(g=>g!=='— вне групп —').sort((a,b)=>groups.get(b).length-groups.get(a).length);
      if (groups.has('— вне групп —')) gn.push('— вне групп —');
      let pos=1;
      gn.forEach(g=>greedyChain(groups.get(g)).forEach(d=>{
        const notes = activeFlags(d).map(f=>FLAG_META[f]?.label||f).join(', ');
        csv += `${cls};${g};${pos++};${d.id};${(d.name_ru||'').replace(/;/g,',')};${d.is_type_locality?'★':''};${notes}\n`;
      }));
    });
    const a=document.createElement('a');
    a.href=URL.createObjectURL(new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'}));
    a.download='shelf_order.csv'; a.click();
  };
}

const VIEW_META = {
  coverage: { title:'Развитие коллекции', sub:'покрытие классов · вишлист спутников · дыры в изоморфных рядах', route:'#/development', render: renderCoverage },
  order:    { title:'Раскладка витрины',  sub:'класс → структурная группа → цепочка природных соседств', route:'#/layout', render: renderOrder }
};
function openExtView(kind){
  const meta = VIEW_META[kind]; if (!meta) return;
  let bg = el('extViewBg');
  if (!bg) {
    bg = document.createElement('div');
    bg.id = 'extViewBg'; bg.className = 'ext-view-bg';
    bg.innerHTML = `<div class="ext-view">
      <div class="ext-view-head">
        <div><div class="t" id="extViewTitle"></div><div class="sub" id="extViewSub"></div></div>
        <button id="extViewClose" style="color:rgba(255,255,255,0.6);background:none;border:none;cursor:pointer;font-size:1.1rem">✕</button>
      </div>
      <div id="extViewBody" style="padding:18px 22px"></div>
    </div>`;
    document.body.appendChild(bg);
    const close = () => { bg.style.display='none'; document.body.style.overflow=''; clearHash(); };
    bg.querySelector('#extViewClose').onclick = close;
    bg.addEventListener('click', e => { if (e.target === bg) close(); });
    document.addEventListener('keydown', e => { if (e.key==='Escape' && bg.style.display==='flex') close(); });
  }
  el('extViewTitle').textContent = meta.title;
  el('extViewSub').textContent = meta.sub;
  bg.style.display = 'flex';
  document.body.style.overflow = 'hidden';
  meta.render(state.filtered && state.filtered.length ? state.filtered : state.data, el('extViewBody'));
  if (location.hash !== meta.route) setHash(meta.route);
}

function injectNav(){
  if (el('extNav')) return;
  const vizPanel = el('vizPanel'); if (!vizPanel) return;
  const nav = document.createElement('div');
  nav.id = 'extNav'; nav.className = 'ext-nav';
  nav.innerHTML = `
    <div class="ext-nav-card" data-view="coverage">
      <div class="t">📈 Развитие коллекции</div>
      <div class="d">Покрытие классов эталонными видами, отсутствующие спутники и химические аналоги — что приобретать дальше</div>
    </div>
    <div class="ext-nav-card" data-view="order">
      <div class="t">🗄 Раскладка витрины</div>
      <div class="d">Теоретический порядок: класс → структурная группа → цепочка природных соседств, с флагами хранения</div>
    </div>`;
  vizPanel.parentElement.insertBefore(nav, vizPanel);
  nav.querySelectorAll('.ext-nav-card').forEach(c => c.addEventListener('click', () => openExtView(c.dataset.view)));
}

// ============================================================
// ПОЛКА, ФИЛЬТРЫ, БЕЙДЖИ, ЗДОРОВЬЕ ДАННЫХ, ПЕРЕХВАТЫ, PWA, INIT
// ============================================================
function openShelfView(code){
  let modal = el('shelfModal');
  if (!modal) {
    modal = document.createElement('div');
    modal.id = 'shelfModal';
    modal.className = 'fixed inset-0 hidden items-center justify-center detail-modal-bg p-4 z-50';
    modal.innerHTML = `
      <div class="bg-white w-full max-w-2xl rounded-2xl border border-slate-200 shadow-xl" style="max-height:88dvh;overflow-y:auto">
        <div class="p-4 flex items-center justify-between sticky top-0" style="background:var(--stone);border-radius:16px 16px 0 0;z-index:2">
          <div class="font-semibold text-white" style="font-family:'Syne',sans-serif" id="shelfTitle">Полка</div>
          <button id="closeShelf" class="text-white/60 hover:text-white">✕</button>
        </div>
        <div id="shelfBody" class="p-4"></div>
      </div>`;
    document.body.appendChild(modal);
    const close = () => { modal.classList.add('hidden'); modal.classList.remove('flex'); clearHash(); };
    modal.querySelector('#closeShelf').onclick = close;
    modal.addEventListener('click', e => { if (e.target === modal) close(); });
  }
  const items = state.data.filter(d => (d.shelf||'') === code).sort((a,b)=>(a.shelf_pos||0)-(b.shelf_pos||0));
  el('shelfTitle').textContent = `Полка ${code} · ${items.length} образцов`;
  el('shelfBody').innerHTML = items.length ? '' : '<div class="text-sm text-stone-400 py-6 text-center">На этой полке пока нет привязанных образцов.</div>';
  items.forEach((d,i) => {
    const row = document.createElement('div');
    row.style.cssText = 'display:flex;gap:12px;align-items:center;padding:8px;border-bottom:1px solid #f0ede8;cursor:pointer;border-radius:8px';
    row.onmouseover = () => row.style.background='#f5f2ed';
    row.onmouseout = () => row.style.background='';
    const flags = activeFlags(d).map(f => FLAG_META[f] ? `<span title="${FLAG_META[f].note}" style="color:${FLAG_META[f].color}">${FLAG_META[f].icon}</span>` : '').join(' ');
    row.innerHTML = `
      <span class="mono text-xs text-stone-400" style="width:22px;text-align:right">${i+1}</span>
      ${d.photo_url ? `<img src="${d.photo_url}" style="width:44px;height:44px;object-fit:cover;border-radius:8px" loading="lazy"/>` : `<span style="width:44px;height:44px;border-radius:8px;background:#f0ede8;display:inline-flex;align-items:center;justify-content:center;color:${cardAccent(d.class)};font-family:'Syne',sans-serif">${(d.abbr||'?').slice(0,3)}</span>`}
      <span style="flex:1;min-width:0">
        <span class="text-sm font-medium text-stone-800">${d.is_type_locality?'<span style="color:var(--amber)">★</span> ':''}${escapeHtml(d.name_ru||'')}</span> ${flags}
        <div class="text-xs text-stone-400 truncate">${escapeHtml(d.locality||'')}</div>
      </span>
      <span class="mono text-stone-300" style="font-size:0.6rem">${escapeHtml(d.id)}</span>`;
    row.onclick = () => { modal.classList.add('hidden'); modal.classList.remove('flex'); openDetailModal(d); };
    el('shelfBody').appendChild(row);
  });
  modal.classList.remove('hidden'); modal.classList.add('flex');
}

function decorateCards(force){
  document.querySelectorAll('.mineral-card[data-id]').forEach(card => {
    if (!force && card.querySelector('.ext-badges')) return;
    const old = card.querySelector('.ext-badges'); if (old) old.remove();
    const d = state.data.find(z => z.id === card.dataset.id);
    if (!d) return;
    const flags = activeFlags(d);
    if (!flags.length) return;
    const span = document.createElement('span');
    span.className = 'ext-badges';
    span.style.cssText = 'position:absolute;top:8px;right:8px;display:flex;gap:3px;font-size:0.72rem;z-index:1;pointer-events:none';
    span.innerHTML = flags.map(f => FLAG_META[f] ? `<span title="${FLAG_META[f].label}" style="color:${FLAG_META[f].color};background:#fff;border-radius:50%;width:18px;height:18px;display:inline-flex;align-items:center;justify-content:center;box-shadow:0 1px 3px rgba(0,0,0,0.12)">${FLAG_META[f].icon}</span>` : '').join('');
    card.style.position = 'relative';
    card.appendChild(span);
  });
}

function injectControls(){
  ['sort_d','sort_m'].forEach(id=>{
    const sel=el(id); if(!sel || sel.querySelector('option[value="tl_first"]')) return;
    const o=document.createElement('option'); o.value='tl_first'; o.textContent='★ Типовые сначала';
    sel.insertBefore(o, sel.children[1]);
  });
  const mk = () => {
    const s=document.createElement('select');
    s.className='select-control px-3 py-1.5 text-xs min-w-[110px]';
    s.innerHTML = `<option value="">Хранение: все</option><option value="any">С особыми условиями</option>` +
      Object.entries(FLAG_META).map(([k,m])=>`<option value="${k}">${m.icon} ${m.label}</option>`).join('');
    return s;
  };
  if (!el('fCons_d')) {
    const d=mk(); d.id='fCons_d';
    document.querySelector('.control-row .hidden.sm\\:flex')?.appendChild(d);
    d.addEventListener('change', ()=>{ const m=el('fCons_m'); if(m) m.value=d.value; applyFilters(); });
  }
  if (!el('fCons_m')) {
    const m=mk(); m.id='fCons_m'; m.className='select-control w-full mt-1 px-3 py-2 text-sm';
    const wrap=document.createElement('div');
    wrap.innerHTML='<label class="text-xs text-stone-500 uppercase tracking-wide">Условия хранения</label>';
    wrap.appendChild(m);
    const sheet = document.querySelector('#sheet .space-y-3');
    if (sheet) sheet.insertBefore(wrap, sheet.lastElementChild);
    m.addEventListener('change', ()=>{ const d=el('fCons_d'); if(d) d.value=m.value; });
  }
}

const _applyFilters = window.applyFilters;
window.applyFilters = function(){
  _applyFilters();
  const consSel = el('fCons_d') || el('fCons_m');
  const cons = consSel ? consSel.value : '';
  const sortSel = el('sort_m')?.value || el('sort_d')?.value || 'name';
  let changed=false;
  if (cons) {
    state.filtered = state.filtered.filter(x => {
      const f = activeFlags(x);
      return cons==='any' ? f.length>0 : f.includes(cons);
    });
    changed=true;
  }
  if (sortSel==='tl_first') {
    state.filtered.sort((a,b)=> (b.is_type_locality?1:0)-(a.is_type_locality?1:0) || (a.name_ru||'').localeCompare(b.name_ru||'','ru'));
    changed=true;
  }
  if (changed) {
    state.page=0;
    updateSummary(state.filtered);
    renderViz(state.filtered);
    renderCards(true);
  }
};

function openHealth(){
  const issues = [];
  const add=(t,arr)=>{ if(arr.length) issues.push({t,arr}); };
  add('Без класса', state.data.filter(d=>!d.class));
  add('Без IMA-имени', state.data.filter(d=>!d.ima_name || d.ima_name==='0'));
  add('IMA-имя — группа, а не вид (уточнить)', state.data.filter(d=>['apatite','tourmaline','hornblende','biotite','lepidolite','wolframite','columbite','garnet','serpentine','chlorite','olivine'].includes(baseName(d.ima_name))));
  add('Без месторождения', state.data.filter(d=>!d.locality));
  add('Без стоимости (авто-черновики)', state.data.filter(d=>d.cost==null||isNaN(d.cost)));
  add('Вид не найден в справочнике', state.data.filter(d=>d.ima_name && d.ima_name!=='0' && !EXT.ref.has(baseName(d.ima_name))));
  let modal=el('healthModal');
  if(!modal){
    modal=document.createElement('div'); modal.id='healthModal';
    modal.className='fixed inset-0 hidden items-center justify-center detail-modal-bg p-4 z-50';
    modal.innerHTML=`<div class="bg-white w-full max-w-2xl rounded-2xl border border-slate-200 shadow-xl" style="max-height:88dvh;overflow-y:auto">
      <div class="p-4 flex items-center justify-between sticky top-0" style="background:var(--stone);border-radius:16px 16px 0 0;z-index:2">
        <div class="font-semibold text-white" style="font-family:'Syne',sans-serif">Здоровье данных</div>
        <button id="closeHealth" class="text-white/60 hover:text-white">✕</button></div>
      <div id="healthBody" class="p-4"></div></div>`;
    document.body.appendChild(modal);
    modal.querySelector('#closeHealth').onclick=()=>{modal.classList.add('hidden');modal.classList.remove('flex');};
    modal.addEventListener('click',e=>{if(e.target===modal){modal.classList.add('hidden');modal.classList.remove('flex');}});
  }
  el('healthBody').innerHTML = issues.map((g,i)=>`
    <details style="margin-bottom:10px" ${i===0?'open':''}>
      <summary style="cursor:pointer;font-size:0.88rem;font-weight:500;color:var(--stone)">${g.t} <span class="mono" style="color:#b8362c">${g.arr.length}</span></summary>
      <div style="padding:6px 0 0 14px;font-size:0.78rem;color:#6a6050">
        ${g.arr.slice(0,60).map(d=>`<div class="ext-health-item" data-id="${escapeHtml(d.id)}" style="cursor:pointer;padding:2px 0;text-decoration:underline dotted">${escapeHtml(d.name_ru||d.id)} <span class="mono" style="color:#b8b0a0;font-size:0.65rem">${escapeHtml(d.id)}</span></div>`).join('')}
        ${g.arr.length>60?`<div style="color:#9a8f7e">+${g.arr.length-60} ещё…</div>`:''}
      </div>
    </details>`).join('') || '<div class="text-sm text-stone-500">Проблем не найдено ✓</div>';
  el('healthBody').querySelectorAll('.ext-health-item').forEach(it=>it.onclick=()=>{
    const d=state.data.find(z=>z.id===it.dataset.id);
    if(d){ modal.classList.add('hidden'); modal.classList.remove('flex'); openDetailModal(d); }
  });
  modal.classList.remove('hidden'); modal.classList.add('flex');
}
function injectHealth(){
  const sheet=document.querySelector('#settingsSheet .settings-row')?.parentElement;
  if (sheet && !el('healthRowM')) {
    const row=document.createElement('div');
    row.className='settings-row'; row.id='healthRowM';
    row.innerHTML=`<div><div class="settings-label">Здоровье данных</div><div class="settings-desc">Неполные записи, черновики, нераспознанные виды</div></div>
      <button id="openHealthM" class="settings-btn">Открыть</button>`;
    sheet.appendChild(row);
    row.querySelector('#openHealthM').onclick=()=>{ closeSettings(); openHealth(); };
  }
  const hdr=document.querySelector('.header-quick-buttons');
  if (hdr && !el('openHealthD')) {
    const b=document.createElement('button'); b.id='openHealthD';
    b.className='px-3 py-1.5 rounded-lg text-xs font-medium transition-all';
    b.style.cssText='background:rgba(200,169,110,0.15);color:var(--amber);border:1px solid rgba(200,169,110,0.3)';
    b.textContent='Данные'; b.onclick=openHealth; hdr.appendChild(b);
  }
}

// ---------- перехваты ----------
const _openDetailModal = window.openDetailModal;
window.openDetailModal = function(x){
  _openDetailModal(x);
  try { buildCard(x); } catch(e){ console.error('[ext] карточка:', e); }
  setHash('#/mineral/' + slugOf(x.id));
};
const _closeDetailModal = window.closeDetailModal;
window.closeDetailModal = function(){ _closeDetailModal(); clearHash(); };
const _renderCards = window.renderCards;
window.renderCards = function(reset){ _renderCards(reset); try{ decorateCards(); }catch(e){} };

// ---------- PWA ----------
function initPWA(){
  if (!document.querySelector('link[rel="manifest"]')) {
    const l=document.createElement('link'); l.rel='manifest'; l.href='manifest.json';
    document.head.appendChild(l);
  }
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(()=>{});
}

// ---------- init ----------
async function extInit(){
  injectCSS();
  computeOwnedSets();
  buildSlugMap();
  injectControls();
  injectNav();
  injectHealth();
  initPWA();
  loadExtraColumns().then(()=>decorateCards(true));
  await loadReference();
  decorateCards(true);
  window.addEventListener('hashchange', handleRoute);
  handleRoute();
  console.log('[ext] готово: справочник', EXT.ref.size, 'видов; основных', EXT.owned.size, '; спутников', EXT.sat.size);
}
let ticks = 0;
const waiter = setInterval(()=>{
  ticks++;
  let ready = false;
  try { ready = typeof state !== 'undefined' && state && Array.isArray(state.data) && state.data.length > 0; }
  catch(e){ clearInterval(waiter); console.error('[ext] нет доступа к данным каталога:', e); return; }
  if (ready) {
    clearInterval(waiter);
    console.log('[ext] каталог готов, образцов:', state.data.length);
    extInit().catch(e => console.error('[ext] ошибка инициализации:', e));
  } else if (ticks > 150) {
    clearInterval(waiter);
    console.error('[ext] каталог не отдал данные за 30 секунд — расширения не запущены.');
  }
}, 200);
})();
