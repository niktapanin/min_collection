// ============================================================
// EXTENSIONS.JS — модуль расширений каталога минералов
// Подключается ПОСЛЕ основного скрипта: <script src="extensions.js"></script>
// Не требует правок основного кода: перехватывает глобальные функции.
// Функции: справочник видов (генезис/парагенезис/этимология/группы),
// химические аналоги, флаги хранения, hash-роутинг (#/mineral/…, #/shelf/…),
// вкладки «Развитие» и «Раскладка», панель здоровья данных, PWA.
// ============================================================
(function(){
'use strict';

const EXT = {
  ref: new Map(),          // base name -> reference row
  owned: new Set(),        // базовые имена основных видов
  sat: new Set(),          // базовые имена видов-спутников
  abbrMap: new Map(),      // аббревиатура -> базовое имя
  slugMap: new Map(),      // slug -> specimen
  extraLoaded: false,      // shelf/acquired_at/flag_override подтянуты
  refReady: false,
  hashGuard: false,
  CACHE_KEY: 'ext_ref_cache_v3',
  CACHE_TTL: 7*24*3600*1000
};
window.EXT = EXT;

// ---------- нормализация имён (зеркало серверного пайплайна) ----------
function baseName(n){
  n = (n||'').toString().trim().toLowerCase();
  try { n = n.normalize('NFD').replace(/[\u0300-\u036f]/g,''); } catch(e){}
  n = n.replace(/-\([a-z0-9+]+\)$/,'');
  return n.replace('baryte','barite');
}
const MANUAL_ABBR = {'Nph':'nepheline','Kln':'kaolinite','Heu':'heulandite','Cpy':'chalcopyrite',
  'Cc':'chalcocite','Ilt':'illite','Lö':'lollingite','Gp-S':'gypsum','Ol-Fo':'forsterite',
  'Fo?':'forsterite','Cr-Di':'diopside','Ms-Fu':'muscovite','Mn-Cal':'calcite','Ckt':'cookeite','Nc':'nacrite'};

const FLAG_META = {
  light_decompose: {icon:'☀', label:'Светочувствителен', note:'Разлагается на свету — вглубь полки, от прямого света', color:'#b8862c'},
  light_fade:      {icon:'◐', label:'Выцветает', note:'Окраска может выцветать — не под прямым светом', color:'#9a8f7e'},
  hygroscopic:     {icon:'💧', label:'Гигроскопичен', note:'Закрытый бокс с силикагелем', color:'#4a7fb8'},
  dehydration:     {icon:'△', label:'Дегидратация', note:'Не пересушивать и не нагревать; стабильная влажность', color:'#7a9e6a'},
  sulfide_decay:   {icon:'⚠', label:'Сульфидный распад', note:'Следить: при белёсом налёте/запахе серы — изолировать в сухой бокс', color:'#c06a2c'},
  radioactive:     {icon:'☢', label:'Радиоактивен', note:'Дистанция от рабочих мест, закрытый бокс', color:'#b83a3a'},
  handle_care:     {icon:'✋', label:'Осторожно', note:'Осторожность при контакте (волокна/токсичность)', color:'#8a5a9e'}
};

function activeFlags(x){
  const b = baseName(x.ima_name);
  const r = EXT.ref.get(b);
  if (!r) return [];
  const over = new Set(Array.isArray(x.flag_override) ? x.flag_override : []);
  return (r.cons_flags||[]).filter(f => !over.has(f));
}
function specimenByBase(b){
  return state.data.find(d => baseName(d.ima_name) === b);
}

// ---------- загрузка справочника ----------
async function fetchRefChunk(names, cols){
  const out = [];
  for (let i=0; i<names.length; i+=90){
    const chunk = names.slice(i,i+90).map(n=>'"'+n.replace(/"/g,'')+'"').join(',');
    const url = `${SB_URL}/rest/v1/species_reference?name=in.(${encodeURIComponent(chunk)})&select=${cols}`;
    try {
      const res = await fetch(url, { headers: SB_HEADERS });
      if (res.ok) out.push(...await res.json());
    } catch(e){ console.warn('ref chunk failed', e); }
  }
  return out;
}
async function fetchCanonical(){
  try {
    const res = await fetch(`${SB_URL}/rest/v1/species_reference?canonical_class=neq.&select=name,name_display,formula_html,canonical_class,assoc_count,struct_group`, { headers: SB_HEADERS });
    return res.ok ? await res.json() : [];
  } catch(e){ return []; }
}

function computeOwnedSets(){
  EXT.owned.clear(); EXT.sat.clear(); EXT.abbrMap.clear();
  state.data.forEach(d => {
    const n = (d.ima_name||'').trim();
    if (n && n!=='0') {
      EXT.owned.add(baseName(n));
      const a = (d.abbr||'').trim();
      if (a && a!=='0' && !EXT.abbrMap.has(a)) EXT.abbrMap.set(a, baseName(n));
    }
  });
  Object.entries(MANUAL_ABBR).forEach(([a,b]) => EXT.abbrMap.set(a,b));
  state.data.forEach(d => {
    const groups = (d.id||'').match(/\(([^)]+)\)/g) || [];
    groups.slice(1).forEach(g => {
      g.slice(1,-1).split(/[,;]\s*/).forEach(a => {
        const b = EXT.abbrMap.get(a.trim());
        if (b && !EXT.owned.has(b)) EXT.sat.add(b);
      });
    });
  });
}

async function loadReference(){
  // кэш
  try {
    const c = JSON.parse(localStorage.getItem(EXT.CACHE_KEY) || 'null');
    if (c && Date.now()-c.t < EXT.CACHE_TTL && Array.isArray(c.rows)) {
      c.rows.forEach(r => EXT.ref.set(r.name, r));
      EXT.refReady = true;
      return;
    }
  } catch(e){}

  const full = 'name,name_display,formula_html,elements,crystal_system,year_published,ima_symbol,struct_group,occurrence,association,assoc_parsed,etymology,cons_flags,canonical_class,assoc_count,analogs';
  const lite = 'name,name_display,formula_html,struct_group,assoc_count,canonical_class,cons_flags';

  const ownedRows = await fetchRefChunk([...EXT.owned, ...EXT.sat], full);
  ownedRows.forEach(r => EXT.ref.set(r.name, r));

  // фаза 2: имена, нужные для «Развития» и аналогов
  const need = new Set();
  ownedRows.forEach(r => {
    (r.assoc_parsed||[]).forEach(n => { if (!EXT.owned.has(n) && !EXT.sat.has(n)) need.add(n); });
    (r.analogs||[]).forEach(a => { if (!EXT.ref.has(a[0])) need.add(a[0]); });
  });
  const canon = await fetchCanonical();
  canon.forEach(r => { if (!EXT.ref.has(r.name)) EXT.ref.set(r.name, r); need.delete(r.name); });
  const extraRows = await fetchRefChunk([...need], lite);
  extraRows.forEach(r => { if (!EXT.ref.has(r.name)) EXT.ref.set(r.name, r); });

  EXT.refReady = true;
  try {
    localStorage.setItem(EXT.CACHE_KEY, JSON.stringify({t: Date.now(), rows: [...EXT.ref.values()]}));
  } catch(e){ console.warn('ref cache too big for localStorage — работаем без кэша'); }
}

// подтянуть новые колонки minerals (их нет в выборке основного скрипта)
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
  EXT.extraLoaded = true;
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
function setHashFor(x){
  EXT.hashGuard = true;
  location.hash = '#/mineral/' + slugOf(x.id);
  setTimeout(()=>{ EXT.hashGuard = false; }, 50);
}
function clearHash(){
  if (location.hash.startsWith('#/mineral/')) {
    EXT.hashGuard = true;
    history.replaceState(null,'', location.pathname + location.search + '#/');
    setTimeout(()=>{ EXT.hashGuard = false; }, 50);
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
  } else if ((m = h.match(/^#\/viz\/([a-z]+)$/))) {
    const tab = document.querySelector(`.viz-tab[data-viz="${m[1]}"]`);
    if (tab) tab.click();
  }
}

// ---------- ПРОСМОТР ПОЛКИ (#/shelf/A-03) ----------
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
    modal.querySelector('#closeShelf').onclick = () => { modal.classList.add('hidden'); modal.classList.remove('flex'); };
    modal.addEventListener('click', e => { if (e.target === modal) { modal.classList.add('hidden'); modal.classList.remove('flex'); } });
  }
  const items = state.data.filter(d => (d.shelf||'') === code).sort((a,b)=>(a.shelf_pos||0)-(b.shelf_pos||0));
  el('shelfTitle').textContent = `Полка ${code} · ${items.length} образцов`;
  el('shelfBody').innerHTML = items.length ? '' : '<div class="text-sm text-stone-400 py-6 text-center">На этой полке пока нет привязанных образцов.<br>Раскладка загружается скриптом из shelf_order.csv.</div>';
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

// ---------- ОБОГАЩЕНИЕ КАРТОЧКИ (детальная модалка) ----------
function chipHtml(b){
  const r = EXT.ref.get(b);
  const label = r ? r.name_display : b;
  if (EXT.owned.has(b))
    return `<span class="el-chip ext-chip" data-b="${b}" style="background:#e6f2ea;border-color:#bcd8c4;color:#2c5c3c" title="есть в коллекции — открыть">${escapeHtml(label)}</span>`;
  if (EXT.sat.has(b))
    return `<span class="el-chip" style="background:#fdf6e3;border-color:#e8d9a8;color:#8a6d1a;cursor:default" title="есть как вид-спутник">${escapeHtml(label)}</span>`;
  return `<span class="el-chip" style="opacity:0.55;cursor:default" title="нет в коллекции">${escapeHtml(label)}</span>`;
}

function enrichModal(x){
  const b = baseName(x.ima_name);
  const r = EXT.ref.get(b);
  const host = document.createElement('div');
  host.id = 'extRefSection';
  const mBody = el('mBody');
  const old = el('extRefSection'); if (old) old.remove();

  let html = '';

  // флаги хранения
  const over = new Set(Array.isArray(x.flag_override) ? x.flag_override : []);
  const allFlags = r ? (r.cons_flags||[]) : [];
  if (allFlags.length) {
    html += `<div style="margin:4px 0 14px"><div class="detail-key" style="margin-bottom:6px">Условия хранения</div><div style="display:flex;flex-direction:column;gap:6px">`;
    allFlags.forEach(f => {
      const m = FLAG_META[f]; if (!m) return;
      const dismissed = over.has(f);
      html += `<div style="display:flex;align-items:flex-start;gap:8px;padding:8px 10px;border-radius:10px;background:${dismissed?'#f5f2ed':'#fdf9f2'};border:1px solid ${dismissed?'#e2ddd5':m.color+'40'};${dismissed?'opacity:0.55':''}">
        <span style="color:${m.color};flex-shrink:0">${m.icon}</span>
        <span style="flex:1;font-size:0.8rem;color:var(--stone)"><b>${m.label}.</b> ${m.note}${dismissed?' <i style="color:#9a8f7e">(флаг снят вами)</i>':''}</span>
        ${state.costUnlocked ? `<button class="ext-flag-toggle" data-flag="${f}" style="font-size:0.7rem;color:#8a8070;background:none;border:none;cursor:pointer;text-decoration:underline;flex-shrink:0">${dismissed?'вернуть':'снять'}</button>` : ''}
      </div>`;
    });
    html += `</div></div>`;
  }

  if (r) {
    const meta = [];
    if (r.struct_group) meta.push(['Структурная группа', escapeHtml(r.struct_group)]);
    if (r.crystal_system) meta.push(['Сингония (IMA)', escapeHtml(r.crystal_system)]);
    if (r.ima_symbol) meta.push(['Символ IMA', `<span class="mono">${escapeHtml(r.ima_symbol)}</span>${r.year_published?` · опубликован ${escapeHtml(r.year_published)}`:''}`]);
    if (meta.length) {
      html += `<div class="divide-y divide-stone-100" style="margin-bottom:14px">` +
        meta.map(([k,v])=>`<div class="detail-row"><div class="detail-key">${k}</div><div class="detail-val">${v}</div></div>`).join('') + `</div>`;
    }
    if (r.occurrence) html += `<div style="margin-bottom:12px"><div class="detail-key" style="margin-bottom:4px">Генезис</div><div style="font-size:0.85rem;color:#4a4238;line-height:1.5">${escapeHtml(r.occurrence)}</div></div>`;
    if (r.assoc_parsed && r.assoc_parsed.length) {
      html += `<div style="margin-bottom:12px"><div class="detail-key" style="margin-bottom:6px">Парагенезис <span style="text-transform:none;letter-spacing:0;color:#b8b0a0">(зелёные — есть в коллекции)</span></div>
        <div style="display:flex;flex-wrap:wrap;gap:5px">${r.assoc_parsed.map(chipHtml).join('')}</div></div>`;
    }
    if (r.etymology) html += `<div style="margin-bottom:12px"><div class="detail-key" style="margin-bottom:4px">Происхождение названия</div><div style="font-size:0.85rem;color:#4a4238;line-height:1.5;font-style:italic">${escapeHtml(r.etymology)}</div></div>`;

    // химические аналоги
    const an = (r.analogs||[]);
    if (an.length) {
      html += `<div style="margin-bottom:12px"><div class="detail-key" style="margin-bottom:6px">Химические аналоги <span style="text-transform:none;letter-spacing:0;color:#b8b0a0">(та же формула, один элемент замещён)</span></div><div style="display:flex;flex-direction:column;gap:4px">`;
      an.slice(0,14).forEach(([other, myEl, otherEl]) => {
        const rr = EXT.ref.get(other);
        const label = rr ? rr.name_display : other;
        const status = EXT.owned.has(other) ? ['есть','#2c5c3c','#e6f2ea'] : (EXT.sat.has(other) ? ['спутник','#8a6d1a','#fdf6e3'] : ['нет','#9a8f7e','#f5f2ed']);
        html += `<div style="display:flex;align-items:center;gap:8px;font-size:0.82rem">
          <span class="mono" style="color:#8a8070;width:70px;flex-shrink:0">${myEl} → ${otherEl}</span>
          <span style="color:var(--stone);font-weight:500;${EXT.owned.has(other)?'cursor:pointer;text-decoration:underline dotted':''}" ${EXT.owned.has(other)?`class="ext-chip" data-b="${other}"`:''}>${escapeHtml(label)}</span>
          ${rr && rr.formula_html ? `<span class="formula" style="color:#8a8070">${rr.formula_html}</span>` : ''}
          <span style="margin-left:auto;font-size:0.65rem;padding:2px 8px;border-radius:10px;background:${status[2]};color:${status[1]};flex-shrink:0">${status[0]}</span>
        </div>`;
      });
      if (an.length > 14) html += `<div style="font-size:0.72rem;color:#9a8f7e">+${an.length-14} ещё…</div>`;
      html += `</div></div>`;
    }
  } else if (x.ima_name && x.ima_name!=='0') {
    html += `<div style="font-size:0.78rem;color:#9a8f7e;margin-bottom:12px">Справочные данные для «${escapeHtml(x.ima_name)}» не найдены (вид вне списка IMA/справочника).</div>`;
  }

  // дата поступления и полка
  if (x.shelf) html += `<div class="detail-row"><div class="detail-key">Полка</div><div class="detail-val"><a href="#/shelf/${encodeURIComponent(x.shelf)}" style="text-decoration:underline">${escapeHtml(x.shelf)}${x.shelf_pos?` · поз. ${x.shelf_pos}`:''}</a></div></div>`;
  if (state.costUnlocked) {
    html += `<div class="detail-row"><div class="detail-key">Дата поступления</div><div class="detail-val" style="display:flex;gap:8px;align-items:center">
      <input type="date" id="extAcq" value="${x.acquired_at||''}" style="border:1.5px solid #d6d0c6;border-radius:8px;padding:4px 8px;font-size:0.82rem"/>
      <button id="extAcqSave" style="font-size:0.72rem;color:#8a8070;background:none;border:none;cursor:pointer;text-decoration:underline">сохранить</button>
    </div></div>`;
  } else if (x.acquired_at) {
    html += `<div class="detail-row"><div class="detail-key">Дата поступления</div><div class="detail-val">${new Date(x.acquired_at).toLocaleDateString('ru-RU')}</div></div>`;
  }

  host.innerHTML = html;
  mBody.appendChild(host);

  host.querySelectorAll('.ext-chip[data-b]').forEach(c => {
    c.onclick = () => { const obj = specimenByBase(c.dataset.b); if (obj) { closeDetailModal(); setTimeout(()=>openDetailModal(obj), 80); } };
  });
  host.querySelectorAll('.ext-flag-toggle').forEach(btn => {
    btn.onclick = async () => {
      const f = btn.dataset.flag;
      let ov = Array.isArray(x.flag_override) ? [...x.flag_override] : [];
      ov = ov.includes(f) ? ov.filter(z=>z!==f) : [...ov, f];
      x.flag_override = ov;
      const rec = state.data.find(d=>d.id===x.id); if (rec) rec.flag_override = ov;
      await fetch(`${SB_URL}/rest/v1/minerals?id=eq.${encodeURIComponent(x.id)}`, {
        method:'PATCH', headers:{...SB_HEADERS,'Content-Type':'application/json','Prefer':'return=minimal'},
        body: JSON.stringify({flag_override: ov})
      });
      enrichModal(x); decorateCards(true);
    };
  });
  const acqBtn = el('extAcqSave');
  if (acqBtn) acqBtn.onclick = async () => {
    const v = el('extAcq').value || null;
    x.acquired_at = v;
    const rec = state.data.find(d=>d.id===x.id); if (rec) rec.acquired_at = v;
    await fetch(`${SB_URL}/rest/v1/minerals?id=eq.${encodeURIComponent(x.id)}`, {
      method:'PATCH', headers:{...SB_HEADERS,'Content-Type':'application/json','Prefer':'return=minimal'},
      body: JSON.stringify({acquired_at: v})
    });
    acqBtn.textContent = '✓';
    setTimeout(()=>acqBtn.textContent='сохранить', 1200);
  };
}

// ---------- БЕЙДЖИ НА КАРТОЧКАХ ----------
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

// экспорт для части 2
window.__EXT_INTERNAL = { EXT, baseName, activeFlags, specimenByBase, loadReference, loadExtraColumns,
  computeOwnedSets, buildSlugMap, setHashFor, clearHash, handleRoute, enrichModal, decorateCards,
  FLAG_META, chipHtml, openShelfView };
})();

// ============================================================
// ЧАСТЬ 2: вкладки «Развитие» и «Раскладка», mindmap с ассоциациями,
// фильтр по флагам, сортировка «типовые сначала», здоровье данных,
// перехват функций, PWA, инициализация
// ============================================================
(function(){
'use strict';
const { EXT, baseName, activeFlags, specimenByBase, loadReference, loadExtraColumns,
  computeOwnedSets, buildSlugMap, setHashFor, clearHash, handleRoute, enrichModal, decorateCards,
  FLAG_META } = window.__EXT_INTERNAL;

// ---------- ВКЛАДКА «РАЗВИТИЕ» ----------
function renderCoverage(list, body){
  if (!EXT.refReady) { body.innerHTML = '<div class="text-sm text-stone-400 text-center py-6">Справочник ещё загружается…</div>'; return; }

  // 1) покрытие эталонного каркаса по классам
  const frame = new Map(); // canonical_class -> [{name, owned|sat|no}]
  EXT.ref.forEach(r => {
    if (!r.canonical_class) return;
    if (!frame.has(r.canonical_class)) frame.set(r.canonical_class, []);
    const st = EXT.owned.has(r.name) ? 2 : (EXT.sat.has(r.name) ? 1 : 0);
    frame.get(r.canonical_class).push({name: r.name, disp: r.name_display, st});
  });
  const rows = [...frame.entries()].map(([cls, arr]) => {
    const have = arr.filter(a=>a.st===2).length;
    return {cls, arr, have, total: arr.length, pct: Math.round(100*have/arr.length)};
  }).sort((a,b)=>a.pct-b.pct);

  let html = `<div class="text-xs text-stone-500 mb-3">Эталонный каркас: классические представители каждого класса. Клик по строке — список дыр.</div>`;
  rows.forEach((r,i) => {
    const missing = r.arr.filter(a=>a.st===0);
    html += `<div class="bar-row ext-cov-row" data-i="${i}">
      <div class="flex items-center justify-between gap-3 mb-1">
        <div class="text-sm font-medium text-stone-800">${r.cls}</div>
        <div class="mono text-xs text-stone-500">${r.have}/${r.total} · ${r.pct}%</div>
      </div>
      <div class="h-2 rounded-full bg-stone-100 overflow-hidden"><div class="bar-fill" style="width:${Math.max(4,r.pct)}%;background:${r.pct<50?'#c06a2c':r.pct<80?'#c8a96e':'#4fad7f'};opacity:0.8"></div></div>
      <div class="ext-cov-detail hidden" style="margin-top:6px;font-size:0.78rem;color:#6a6050">
        ${missing.length ? 'Не хватает: ' + missing.map(m=>`<a href="https://www.mindat.org/search.php?search=${encodeURIComponent(m.disp||m.name)}" target="_blank" rel="noopener" style="text-decoration:underline dotted">${m.disp||m.name}</a>`).join(', ') : 'Каркас закрыт полностью ✓'}
      </div>
    </div>`;
  });

  // 2) топ отсутствующих спутников (ассоциативный вишлист)
  const linkCnt = new Map(), via = new Map();
  EXT.owned.forEach(b => {
    const r = EXT.ref.get(b); if (!r) return;
    (r.assoc_parsed||[]).forEach(a => {
      if (EXT.owned.has(a) || EXT.sat.has(a) || !EXT.ref.has(a)) return;
      linkCnt.set(a,(linkCnt.get(a)||0)+1);
      if (!via.has(a)) via.set(a,new Set());
      if (via.get(a).size<6) via.get(a).add(EXT.ref.get(b)?.name_display||b);
    });
  });
  const wl = [...linkCnt.entries()]
    .map(([n,c]) => ({n, c, common: EXT.ref.get(n)?.assoc_count||0, disp: EXT.ref.get(n)?.name_display||n}))
    .sort((a,b)=> (b.c*2+Math.min(b.common,50)*0.2) - (a.c*2+Math.min(a.common,50)*0.2)).slice(0,20);
  html += `<div class="text-xs text-stone-500 mt-5 mb-2 font-semibold" style="font-family:'Syne',sans-serif;color:var(--stone);font-size:0.85rem">Топ отсутствующих видов-спутников</div>`;
  wl.forEach(w => {
    html += `<div style="display:flex;align-items:baseline;gap:8px;padding:5px 8px;font-size:0.82rem;border-bottom:1px solid #f5f2ed">
      <a href="https://www.mindat.org/search.php?search=${encodeURIComponent(w.disp)}" target="_blank" rel="noopener" style="font-weight:500;color:var(--stone);text-decoration:underline dotted">${w.disp}</a>
      <span class="mono" style="font-size:0.7rem;color:#8a8070">${w.c} связ.</span>
      <span style="font-size:0.7rem;color:#b8b0a0;flex:1;text-align:right" class="truncate">спутник: ${[...(via.get(w.n)||[])].slice(0,3).join(', ')}</span>
    </div>`;
  });

  // 3) топ отсутствующих химических аналогов
  const anCnt = new Map(), anVia = new Map();
  EXT.owned.forEach(b => {
    const r = EXT.ref.get(b); if (!r) return;
    (r.analogs||[]).forEach(([other,e1,e2]) => {
      if (EXT.owned.has(other) || EXT.sat.has(other)) return;
      anCnt.set(other,(anCnt.get(other)||0)+1);
      if (!anVia.has(other)) anVia.set(other, `${r.name_display||b} (${e1}→${e2})`);
    });
  });
  const an = [...anCnt.entries()].sort((a,b)=>b[1]-a[1]).slice(0,20);
  html += `<div class="text-xs text-stone-500 mt-5 mb-2 font-semibold" style="font-family:'Syne',sans-serif;color:var(--stone);font-size:0.85rem">Топ отсутствующих химических аналогов</div>`;
  an.forEach(([n,c]) => {
    const rr = EXT.ref.get(n);
    html += `<div style="display:flex;align-items:baseline;gap:8px;padding:5px 8px;font-size:0.82rem;border-bottom:1px solid #f5f2ed">
      <a href="https://www.mindat.org/search.php?search=${encodeURIComponent(rr?.name_display||n)}" target="_blank" rel="noopener" style="font-weight:500;color:var(--stone);text-decoration:underline dotted">${rr?.name_display||n}</a>
      ${rr?.formula_html ? `<span class="formula" style="color:#8a8070;font-size:0.72rem">${rr.formula_html}</span>` : ''}
      <span class="mono" style="font-size:0.7rem;color:#8a8070">${c} вид.</span>
      <span style="font-size:0.7rem;color:#b8b0a0;flex:1;text-align:right" class="truncate">напр.: ${anVia.get(n)||''}</span>
    </div>`;
  });

  body.innerHTML = html;
  body.querySelectorAll('.ext-cov-row').forEach(row => {
    row.addEventListener('click', () => row.querySelector('.ext-cov-detail')?.classList.toggle('hidden'));
  });
}

// ---------- ВКЛАДКА «РАСКЛАДКА» ----------
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
    remaining.forEach(d=>{ const s=score(last,d); if (s>bs || (s===bs && (d.name_ru||'')<(best.name_ru||''))) { bs=s; best=d; } });
    chain.push(best); remaining=remaining.filter(d=>d!==best);
  }
  return chain;
}
function renderOrder(list, body){
  if (!EXT.refReady) { body.innerHTML = '<div class="text-sm text-stone-400 text-center py-6">Справочник ещё загружается…</div>'; return; }
  const byClass = new Map();
  list.forEach(d => { const c=d.class||'без класса'; if(!byClass.has(c)) byClass.set(c,[]); byClass.get(c).push(d); });
  const classes=[...byClass.keys()].sort((a,b)=>(classSortKey(a)-classSortKey(b)));
  let html = `<div class="text-xs text-stone-500 mb-3">Теоретический порядок: класс → структурная группа → цепочка природных соседств (парагенезис + изоморфизм). Значки — условия хранения. <button id="extOrderCsv" style="text-decoration:underline;color:var(--stone);background:none;border:none;cursor:pointer;font-size:inherit">Скачать CSV текущей выборки</button></div>`;
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
    html += `<div style="margin-bottom:14px"><div style="font-family:'Syne',sans-serif;font-size:0.9rem;color:var(--stone);margin-bottom:6px;border-bottom:2px solid ${cardAccent(cls)}40;padding-bottom:3px">${cls} <span class="mono" style="font-size:0.7rem;color:#9a8f7e">${items.length}</span></div>`;
    gnames.forEach(g=>{
      const chain=greedyChain(groups.get(g));
      html += `<div style="margin-bottom:6px"><div style="font-size:0.72rem;color:#8a8070;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:3px">${g} · ${chain.length}</div><div style="display:flex;flex-wrap:wrap;gap:4px">`;
      chain.forEach((d,i)=>{
        const prev = i>0 ? chain[i-1] : null;
        const linked = prev && assocLinked(baseName(d.ima_name), baseName(prev.ima_name));
        const flags = activeFlags(d).map(f=>FLAG_META[f]?FLAG_META[f].icon:'').join('');
        html += `${i>0?`<span style="color:${linked?'#4fad7f':'#ddd8d0'};align-self:center;font-size:0.7rem">${linked?'—':'·'}</span>`:''}<span class="el-chip ext-order-chip" data-id="${escapeHtml(d.id)}" style="background:${cardAccent(d.class)}12;border-color:${cardAccent(d.class)}35" title="${escapeHtml(d.locality||'')}">${d.is_type_locality?'★ ':''}${escapeHtml(d.name_ru||d.id)}${flags?' '+flags:''}</span>`;
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

// ---------- MINDMAP: рёбра парагенезиса и изоморфизма ----------
const _renderMindmap = window.renderMindmap;
window.renderMindmap = function(center){
  const b0 = baseName(center.ima_name);
  const r0 = EXT.ref.get(b0);
  if (!EXT.refReady || !r0) return _renderMindmap(center);

  const panel = el('mindmapPanel'); if (!panel) return; panel.innerHTML='';
  const manualIds = new Set(Array.isArray(center.related_ids)?center.related_ids:[]);
  const links=[]; const seen=new Set();
  const push=(o,type)=>{ if(o.id!==center.id && !seen.has(o.id)){ links.push({o,type}); seen.add(o.id);} };
  state.data.forEach(o=>{ if(manualIds.has(o.id)) push(o,'manual'); });
  // парагенезис
  const assocSet = new Set(r0.assoc_parsed||[]);
  state.data.forEach(o=>{
    const b=baseName(o.ima_name); if(!b) return;
    const rr=EXT.ref.get(b);
    if (assocSet.has(b) || (rr && (rr.assoc_parsed||[]).includes(b0))) push(o,'assoc');
  });
  // изоморфные аналоги
  const isoSet = new Set((r0.analogs||[]).map(a=>a[0]));
  state.data.forEach(o=>{ const b=baseName(o.ima_name); if(b && isoSet.has(b)) push(o,'iso'); });
  // то же месторождение
  if (center.locality) state.data.filter(o=>o.locality===center.locality).slice(0,10).forEach(o=>push(o,'locality'));

  const shown = links.slice(0,16);
  if (!shown.length) return _renderMindmap(center);

  const typeColors={manual:'#c8a96e',assoc:'#4fad7f',iso:'#b86a9e',locality:'#7ba8a0'};
  const typeLabels={manual:'ручная связь',assoc:'парагенезис (справочник)',iso:'изоморфный аналог',locality:'то же месторождение'};
  const nodes=[{id:center.id,name:center.name_ru||'—',isCenter:true,accent:cardAccent(center.class),obj:center},
    ...shown.map(l=>({id:l.o.id,name:l.o.name_ru||'—',isCenter:false,accent:cardAccent(l.o.class),type:l.type,obj:l.o}))];
  const dlinks=shown.map(l=>({source:center.id,target:l.o.id,type:l.type}));
  const w=panel.clientWidth,h=panel.clientHeight;
  const svg=d3.select(panel).append('svg').attr('width',w).attr('height',h).style('display','block');
  const sim=d3.forceSimulation(nodes)
    .force('link',d3.forceLink(dlinks).id(d=>d.id).distance(115))
    .force('charge',d3.forceManyBody().strength(-280))
    .force('center',d3.forceCenter(w/2,h/2))
    .force('collide',d3.forceCollide().radius(40));
  const link=svg.append('g').selectAll('line').data(dlinks).join('line')
    .attr('class','mindmap-link').attr('stroke',d=>typeColors[d.type]||'#c8bfaf')
    .attr('stroke-dasharray',d=>d.type==='iso'?'4,3':null).attr('stroke-width',2);
  const node=svg.append('g').selectAll('g').data(nodes).join('g').style('cursor','pointer')
    .call(d3.drag()
      .on('start',(e,d)=>{if(!e.active)sim.alphaTarget(0.3).restart();d.fx=d.x;d.fy=d.y;})
      .on('drag',(e,d)=>{d.fx=e.x;d.fy=e.y;})
      .on('end',(e,d)=>{if(!e.active)sim.alphaTarget(0);d.fx=null;d.fy=null;}));
  node.append('circle').attr('r',d=>d.isCenter?26:18).attr('fill',d=>d.isCenter?'#8a8070':d.accent).attr('stroke','#fff').attr('stroke-width',2);
  node.append('text').attr('class','mindmap-node-label').attr('text-anchor','middle')
    .attr('dy',d=>d.isCenter?40:32).text(d=>d.name.length>16?d.name.slice(0,14)+'…':d.name);
  node.append('title').text(d=>d.isCenter?d.name:`${d.name} — ${typeLabels[d.type]||''}`);
  node.on('click',(e,d)=>{ if(d.isCenter) return; closeMindmapModal(); closeDetailModal(); setTimeout(()=>openDetailModal(d.obj),100); });
  sim.on('tick',()=>{
    link.attr('x1',d=>d.source.x).attr('y1',d=>d.source.y).attr('x2',d=>d.target.x).attr('y2',d=>d.target.y);
    node.attr('transform',d=>`translate(${d.x},${d.y})`);
  });
  // легенда: дополнить новыми типами (один раз)
  const legend = document.querySelector('#mindmapBody > div:last-child');
  if (legend && !legend.querySelector('.ext-legend')) {
    const s=document.createElement('span'); s.className='ext-legend';
    s.innerHTML=`<span style="display:inline-block;width:18px;height:2px;background:#4fad7f;vertical-align:middle"></span> парагенезис · <span style="display:inline-block;width:18px;height:2px;background:#b86a9e;vertical-align:middle;border-top:1px dashed #b86a9e"></span> изоморфизм`;
    legend.insertBefore(s, legend.firstChild);
  }
};

// ---------- ФИЛЬТР ПО ХРАНЕНИЮ + СОРТИРОВКА «ТИПОВЫЕ» ----------
function injectControls(){
  // sort options
  ['sort_d','sort_m'].forEach(id=>{
    const sel=el(id); if(!sel || sel.querySelector('option[value="tl_first"]')) return;
    const o=document.createElement('option'); o.value='tl_first'; o.textContent='★ Типовые сначала';
    sel.insertBefore(o, sel.children[1]);
  });
  // conservation filter
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
    document.querySelector('#sheet .space-y-3')?.insertBefore(wrap, document.querySelector('#sheet .space-y-3').lastElementChild);
    m.addEventListener('change', ()=>{ const d=el('fCons_d'); if(d) d.value=m.value; });
  }
}

// applyFilters: обёртка — базовая фильтрация + пост-фильтр по флагам + новая сортировка
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

// ---------- ЗДОРОВЬЕ ДАННЫХ ----------
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
function injectHealthRow(){
  const sheet=document.querySelector('#settingsSheet .settings-row')?.parentElement;
  if (!sheet || el('healthRowM')) return;
  const row=document.createElement('div');
  row.className='settings-row'; row.id='healthRowM';
  row.innerHTML=`<div><div class="settings-label">Здоровье данных</div><div class="settings-desc">Неполные записи, черновики, нераспознанные виды</div></div>
    <button id="openHealthM" class="settings-btn">Открыть</button>`;
  sheet.appendChild(row);
  row.querySelector('#openHealthM').onclick=()=>{ closeSettings(); openHealth(); };
  // и на десктопе — рядом с «Диапазоны ₽»
  const hdr=document.querySelector('.header-quick-buttons');
  if (hdr && !el('openHealthD')) {
    const b=document.createElement('button'); b.id='openHealthD';
    b.className='px-3 py-1.5 rounded-lg text-xs font-medium transition-all';
    b.style.cssText='background:rgba(200,169,110,0.15);color:var(--amber);border:1px solid rgba(200,169,110,0.3)';
    b.textContent='Данные'; b.onclick=openHealth; hdr.appendChild(b);
  }
}

// ---------- ПЕРЕХВАТЫ ----------
const _openDetailModal = window.openDetailModal;
window.openDetailModal = function(x){ _openDetailModal(x); try{ enrichModal(x); }catch(e){console.warn(e);} setHashFor(x); };
const _closeDetailModal = window.closeDetailModal;
window.closeDetailModal = function(){ _closeDetailModal(); clearHash(); };
const _renderCards = window.renderCards;
window.renderCards = function(reset){ _renderCards(reset); try{ decorateCards(); }catch(e){} };
const _renderViz = window.renderViz;
window.renderViz = function(list){
  const body = el('vizBody');
  if (state.currentViz==='coverage') { body.innerHTML=''; return renderCoverage(list, body); }
  if (state.currentViz==='order')    { body.innerHTML=''; return renderOrder(list, body); }
  return _renderViz(list);
};

function injectVizTabs(){
  const tabs=el('vizTabs'); if(!tabs || tabs.querySelector('[data-viz="coverage"]')) return;
  [['coverage','Развитие'],['order','Раскладка']].forEach(([k,label])=>{
    const b=document.createElement('button');
    b.className='viz-tab'; b.dataset.viz=k; b.textContent=label;
    b.addEventListener('click',()=>{
      tabs.querySelectorAll('.viz-tab').forEach(t=>t.classList.remove('active'));
      b.classList.add('active'); state.currentViz=k;
      el('showAllClasses').style.display='none';
      renderViz(state.filtered);
    });
    tabs.appendChild(b);
  });
}

// ---------- PWA ----------
function initPWA(){
  if (!document.querySelector('link[rel="manifest"]')) {
    const l=document.createElement('link'); l.rel='manifest'; l.href='manifest.json';
    document.head.appendChild(l);
  }
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(()=>{});
}

// ---------- ИНИЦИАЛИЗАЦИЯ ----------
async function extInit(){
  computeOwnedSets();
  buildSlugMap();
  injectControls();
  injectVizTabs();
  injectHealthRow();
  initPWA();
  loadExtraColumns().then(()=>decorateCards(true));
  await loadReference();
  decorateCards(true);
  window.addEventListener('hashchange', handleRoute);
  handleRoute();
  if (state.currentViz==='coverage'||state.currentViz==='order') renderViz(state.filtered);
  console.log('[ext] готово: справочник', EXT.ref.size, 'видов; основных', EXT.owned.size, '; спутников', EXT.sat.size);
}
const waiter = setInterval(()=>{
  if (window.state && state.data && state.data.length) { clearInterval(waiter); extInit().catch(e=>console.error('[ext]',e)); }
}, 200);
setTimeout(()=>clearInterval(waiter), 60000);
})();
