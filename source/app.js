// What's Cooking with Brian – shared cooking checklist PWA.
// Source file. The build step bundles this into ../app.js (see build.mjs).

import XLSX from 'xlsx/dist/xlsx.mini.min.js';
import { initializeApp } from 'firebase/app';
import {
  initializeAuth, indexedDBLocalPersistence, browserLocalPersistence,
  signInWithEmailAndPassword, sendPasswordResetEmail, onAuthStateChanged, signOut, connectAuthEmulator,
} from 'firebase/auth';
import {
  initializeFirestore, persistentLocalCache, persistentMultipleTabManager,
  collection, doc, onSnapshot, setDoc, deleteField, serverTimestamp, connectFirestoreEmulator,
} from 'firebase/firestore';

const APP_NAME = "What's Cooking with Brian";

// ---------------------------------------------------------------- helpers
const $ = (s, r = document) => r.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const norm = (v) => (v == null ? '' : String(v).replace(/\s+/g, ' ').trim());
const slug = (s) => norm(s).toLowerCase().normalize('NFKD').replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '') || 'x';
const ls = {
  get(k, d) { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* storage full or blocked */ } },
};
function hash(str) { // cyrb53, short stable id for a line of text
  let h1 = 0xdeadbeef, h2 = 0x41c6ce57;
  for (let i = 0; i < str.length; i++) { const c = str.charCodeAt(i); h1 = Math.imul(h1 ^ c, 2654435761); h2 = Math.imul(h2 ^ c, 1597334677); }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(36);
}
function normTime(s) { // "0:27", "0:27:00", "1:03" -> "0:27"; anything else unchanged
  const m = norm(s).match(/^(\d{1,2}):(\d{2})(?::00)?$/);
  return m ? `${+m[1]}:${m[2]}` : norm(s);
}
function toMinutes(s) { const m = String(s).match(/^(\d+):(\d{2})$/); return m ? +m[1] * 60 + +m[2] : null; }
function fmtDuration(min) { if (min == null) return ''; const h = Math.floor(min / 60), m = min % 60; return h ? `${h} h${m ? ` ${m} min` : ''}` : `${m} min`; }
const displayMealName = (s) => norm(s).replace(/^\d+\s*[.):-]?\s*/, '') || norm(s);
const stem = (w) => w.toLowerCase().replace(/[^a-z]/g, '').replace(/(es|s)$/, '');

let toastTimer;
function toast(msg, ms = 2600) {
  const t = $('#toast'); t.textContent = msg; t.classList.add('show');
  clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.remove('show'), ms);
}
function confirmDialog(text, okLabel = 'Clear') {
  const dlg = $('#confirm');
  $('#confirm-text').textContent = text; $('#confirm-ok').textContent = okLabel;
  return new Promise((resolve) => {
    dlg.addEventListener('close', () => resolve(dlg.returnValue === 'ok'), { once: true });
    dlg.returnValue = ''; dlg.showModal();
  });
}
const CHECK_SVG = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 12.5l4.2 4.2L19 7" fill="none" stroke="#fff" stroke-width="3.2" stroke-linecap="round" stroke-linejoin="round"/></svg>';

// ---------------------------------------------------------------- recipe parsing
function rowsFromWorkbook(wb) {
  const find = (name) => wb.SheetNames.find((n) => n.trim().toLowerCase() === name);
  const toRows = (n) => (n ? XLSX.utils.sheet_to_json(wb.Sheets[n], { defval: '', raw: false }) : []);
  const recipesName = find('recipes') || wb.SheetNames[0];
  return { recipes: toRows(recipesName), shopping: toRows(find('shopping')) };
}
function keyed(row) { // lower-case, trimmed column names
  const o = {}; for (const k in row) o[k.trim().toLowerCase()] = norm(row[k]); return o;
}

function buildModel({ recipes, shopping }) {
  const meals = new Map();
  let seq = 0;
  for (const raw of recipes) {
    const r = keyed(raw); seq++;
    const recipe = r.recipe, section = (r.section || '').toLowerCase();
    if (!recipe || !section) continue;
    const mealName = r.meal || recipe;
    let meal = meals.get(mealName);
    if (!meal) { meal = { name: mealName, title: displayMealName(mealName), id: 'm-' + slug(mealName), dishes: new Map(), planTotal: '', planFirst: '', plan: [] }; meals.set(mealName, meal); }
    let dish = meal.dishes.get(recipe);
    if (!dish) { dish = { name: recipe, type: /side/i.test(r.type) ? 'Side' : 'Main', makes: '', groups: [], steps: [], tips: [], seq }; meal.dishes.set(recipe, dish); }
    const order = r.order === '' ? NaN : parseFloat(r.order);
    switch (section) {
      case 'makes': dish.makes = r.text; break;
      case 'plan total': meal.planTotal = r.text; break;
      case 'plan first': meal.planFirst = r.text; break;
      case 'plan': meal.plan.push({ when: normTime(r.amount), order, text: r.text, seq }); break;
      case 'ingredient': {
        let g = dish.groups.find((x) => x.name === r.group);
        if (!g) { g = { name: r.group, items: [] }; dish.groups.push(g); }
        g.items.push({ amount: r.amount, text: r.text }); break;
      }
      case 'step': dish.steps.push({ order, amount: r.amount, text: r.text, seq }); break;
      case 'tip': if (r.text) dish.tips.push(r.text); break;
      default: break;
    }
  }
  const byOrder = (a, b) => (isNaN(a.order) ? 1e9 : a.order) - (isNaN(b.order) ? 1e9 : b.order) || a.seq - b.seq;

  const out = [];
  for (const meal of meals.values()) {
    const dishes = [...meal.dishes.values()].sort((a, b) => (a.type === b.type ? a.seq - b.seq : a.type === 'Main' ? -1 : 1));
    meal.dishes = dishes;
    const used = new Map();
    const uniq = (base) => { const n = (used.get(base) || 0) + 1; used.set(base, n); return hash(n > 1 ? `${base}#${n}` : base); };
    dishes.forEach((d, i) => {
      d.color = Math.min(i, 3); d.anchor = `${meal.id}-d${i}`;
      d.steps.sort(byOrder);
      d.steps.forEach((s, j) => { s.num = isNaN(s.order) ? j + 1 : s.order; s.key = uniq(`step|${d.name}|${s.text}`); });
      for (const g of d.groups) for (const it of g.items) it.key = uniq(`ing|${d.name}|${g.name}|${it.text}`);
    });
    meal.plan.sort(byOrder);
    meal.plan.forEach((p) => { p.key = uniq(`plan|${p.when}|${p.text}`); p.min = toMinutes(p.when); });
    meal.main = dishes.find((d) => d.type === 'Main') || dishes[0];
    meal.sides = dishes.filter((d) => d !== meal.main);
    const mins = meal.plan.map((p) => p.min).filter((m) => m != null);
    meal.totalMin = mins.length ? Math.max(...mins) : null;
    meal.keys = dishes.flatMap((d) => d.steps.map((s) => s.key));
    computeStarts(meal);
    out.push(meal);
  }
  out.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));

  const sections = new Map();
  const seen = new Map();
  for (const raw of shopping) {
    const r = keyed(raw); if (!r.item) continue;
    const sec = r.section || 'Other';
    if (!sections.has(sec)) sections.set(sec, []);
    const base = `shop|${r.item}`; const n = (seen.get(base) || 0) + 1; seen.set(base, n);
    sections.get(sec).push({ item: r.item, amount: r.amount, usedIn: r['used in'], key: hash(n > 1 ? `${base}#${n}` : base) });
  }
  const pref = ['meat', 'produce', 'dairy', 'pantry'];
  const rank = (s) => { const i = pref.indexOf(s.toLowerCase()); return i < 0 ? 99 : i; };
  const shop = [...sections.entries()].map(([name, items], i) => ({ name, items, i })).sort((a, b) => rank(a.name) - rank(b.name) || a.i - b.i);
  return { meals: out, shopping: shop, shopKeys: shop.flatMap((s) => s.items.map((x) => x.key)) };
}

// Find references like "(Carrot steps 2-6)" or "Main step 1" in plan text.
const REF_RE = /\b([A-Z][A-Za-z'-]*(?: [a-z][a-z'-]*)?) (steps?) (\d+)((?:\s*[-–]\s*\d+)?)/g;
function resolveRef(phrase, meal) {
  const p = phrase.toLowerCase();
  if (p === 'main') return meal.main;
  if (p === 'side' || p === 'sides') return meal.sides.length === 1 ? meal.sides[0] : null;
  const words = p.split(' ').map(stem).filter(Boolean);
  let best = null, bestScore = 0;
  for (const d of [...meal.sides, meal.main]) {
    const nameWords = new Set(d.name.split(/[\s-]+/).map(stem));
    const score = words.filter((w) => nameWords.has(w)).length + (d === meal.main ? -0.5 : 0);
    if (score > bestScore) { best = d; bestScore = score; }
  }
  return best;
}
function computeStarts(meal) {
  const start = new Map();
  for (const p of meal.plan) {
    p.dishes = new Set();
    for (const m of p.text.matchAll(REF_RE)) { const d = resolveRef(m[1], meal); if (d) p.dishes.add(d); }
    if (p.min == null) continue; // "Before" rows are just preheating
    for (const d of p.dishes) if (!start.has(d)) start.set(d, p.min);
  }
  meal.starts = meal.dishes.filter((d) => start.has(d)).map((d) => ({ dish: d, min: start.get(d) }))
    .sort((a, b) => a.min - b.min || a.dish.color - b.dish.color);
}
function linkRefs(text, meal) {
  return esc(text).replace(REF_RE, (all, name, _w, first) => {
    const d = resolveRef(name, meal);
    if (!d) return all;
    return `<a class="ref chip-${d.color}" href="#" data-jump="${d.anchor}-s${first}">${all}</a>`;
  });
}

// ---------------------------------------------------------------- state
const state = {
  model: null,
  recipeInfo: ls.get('recipes-info', null), // {source, at, error}
  ticks: {},               // scopeId -> { key: true }
  sync: 'none',            // ok | pending | offline | local | none | error
  user: null,
  sheetUrl: '',
  firebase: null,          // { auth, db } when configured
  authReady: false,
  dbError: '',
};
const CFG = window.APP_CONFIG || {};
const fbConfig = CFG.firebase && CFG.firebase.apiKey && CFG.firebase.projectId ? CFG.firebase : null;

// ---------------------------------------------------------------- ticks store
let unsubTicks = null, unsubCfg = null, lastSnapPending = false;

function setTick(scope, key, value) {
  if (!state.firebase) {
    const t = state.ticks[scope] || (state.ticks[scope] = {});
    if (value) t[key] = true; else delete t[key];
    ls.set('local-ticks', state.ticks); refreshTicks(); return;
  }
  const { db } = state.firebase;
  // optimistic local update; Firestore also fires a local snapshot immediately
  const t = state.ticks[scope] || (state.ticks[scope] = {});
  if (value) t[key] = true; else delete t[key];
  refreshTicks();
  setDoc(doc(db, 'ticks', scope), { k: { [key]: value ? true : deleteField() }, at: serverTimestamp() }, { merge: true })
    .catch((e) => { console.error(e); toast('Could not save tick: ' + (e.code || e.message)); });
}
async function clearScope(scope) {
  if (!state.firebase) { state.ticks[scope] = {}; ls.set('local-ticks', state.ticks); refreshTicks(); return; }
  state.ticks[scope] = {}; refreshTicks();
  setDoc(doc(state.firebase.db, 'ticks', scope), { k: {}, at: serverTimestamp() })
    .catch((e) => toast('Could not clear: ' + (e.code || e.message)));
}
function startFirestore() {
  const { db } = state.firebase;
  stopFirestore();
  unsubTicks = onSnapshot(collection(db, 'ticks'), { includeMetadataChanges: true }, (snap) => {
    const ticks = {};
    snap.forEach((d) => { ticks[d.id] = { ...(d.data().k || {}) }; });
    state.ticks = ticks; state.dbError = '';
    lastSnapPending = snap.metadata.hasPendingWrites;
    updateSync(snap.metadata.fromCache);
    refreshTicks();
  }, (err) => {
    console.error(err);
    state.dbError = err.code === 'permission-denied'
      ? 'This account is not allowed to read the shared ticks. Check the two email addresses in the Firestore rules (SETUP.md step 4).'
      : 'Sync error: ' + (err.code || err.message);
    state.sync = 'error'; renderSync(); route();
  });
  unsubCfg = onSnapshot(doc(db, 'config', 'app'), (snap) => {
    const url = (snap.exists() && snap.data().sheetUrl) || '';
    if (url !== state.sheetUrl) { state.sheetUrl = url; ls.set('sheet-url', url); loadRecipes(); }
  }, () => {});
}
function stopFirestore() { if (unsubTicks) unsubTicks(); if (unsubCfg) unsubCfg(); unsubTicks = unsubCfg = null; }

function updateSync(fromCache) {
  if (!state.firebase) state.sync = 'local';
  else if (!state.user) state.sync = 'none';
  else if (state.dbError) state.sync = 'error';
  else if (!navigator.onLine) state.sync = 'offline';
  else if (lastSnapPending) state.sync = 'pending';
  else if (fromCache) state.sync = 'pending';
  else state.sync = 'ok';
  renderSync();
}
function renderSync() {
  const el = $('#sync');
  const labels = { ok: 'Synced', pending: 'Syncing…', offline: lastSnapPending ? 'Offline · will sync' : 'Offline', local: 'This phone only', none: 'Signed out', error: 'Sync error' };
  el.className = 'sync ' + state.sync; el.lastElementChild.textContent = labels[state.sync] || '';
}
addEventListener('online', () => updateSync(false));
addEventListener('offline', () => updateSync(true));

// ---------------------------------------------------------------- recipe loading
function sheetSources(url) {
  const pub = url.match(/\/spreadsheets\/d\/e\/([\w-]+)/);
  if (pub) return { kind: 'xlsx', url: `https://docs.google.com/spreadsheets/d/e/${pub[1]}/pub?output=xlsx` };
  const id = url.match(/\/spreadsheets\/d\/([\w-]{20,})/);
  if (id) {
    const base = `https://docs.google.com/spreadsheets/d/${id[1]}/gviz/tq?tqx=out:csv&headers=1&sheet=`;
    return { kind: 'csv', recipes: base + 'Recipes', shopping: base + 'Shopping' };
  }
  return null;
}
async function fetchOk(url, type) {
  const r = await fetch(url, { cache: 'no-store' });
  if (!r.ok) throw new Error(`HTTP ${r.status}`);
  return type === 'buf' ? r.arrayBuffer() : r.text();
}
async function fetchRows(src) {
  if (src === 'bundled') return rowsFromWorkbook(XLSX.read(await fetchOk('recipes.xlsx', 'buf'), { type: 'array' }));
  if (src.kind === 'xlsx') return rowsFromWorkbook(XLSX.read(await fetchOk(src.url, 'buf'), { type: 'array' }));
  const csv = async (u) => {
    const text = await fetchOk(u, 'text');
    if (/^\s*</.test(text)) throw new Error('Google returned a web page instead of data. Is the sheet shared as "Anyone with the link"?');
    const wb = XLSX.read(text, { type: 'string', raw: true });
    return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '', raw: false });
  };
  const [recipes, shopping] = await Promise.all([csv(src.recipes), csv(src.shopping).catch(() => [])]);
  return { recipes, shopping };
}
function setRows(rows, info) {
  const model = buildModel(rows);
  if (!model.meals.length) throw new Error('No recipes found. Is there a tab named "Recipes" with the usual columns?');
  const sig = JSON.stringify(rows);
  const changed = sig !== state.rowsSig;
  state.rowsSig = sig; state.model = model;
  ls.set('recipes-cache', rows);
  state.recipeInfo = info; ls.set('recipes-info', info);
  if (changed) route();
}
let loading = null;
async function loadRecipes() {
  if (loading) return loading;
  loading = (async () => {
    const url = state.sheetUrl;
    const src = url ? sheetSources(url) : 'bundled';
    try {
      if (!src) throw new Error('That does not look like a Google Sheets link.');
      setRows(await fetchRows(src), { source: url ? 'sheet' : 'bundled', at: Date.now(), error: '' });
    } catch (e) {
      console.warn('Recipe load failed', e);
      const info = { ...(state.recipeInfo || {}), error: (url ? 'Could not load the Google Sheet: ' : 'Could not load recipes: ') + e.message };
      state.recipeInfo = info; ls.set('recipes-info', info);
      if (!state.model) { // nothing on screen yet: fall back to the built-in file
        try { setRows(await fetchRows('bundled'), { source: 'bundled', at: Date.now(), error: info.error }); } catch { route(); }
      } else if (location.hash.startsWith('#/settings')) route();
    } finally { loading = null; }
  })();
  return loading;
}

// ---------------------------------------------------------------- views
const view = () => $('#view');
function setTitle(text, back) {
  $('#title').textContent = text; $('#back').hidden = !back;
  document.title = text === APP_NAME ? APP_NAME : `${text} · ${APP_NAME}`;
}
function setTab(name) {
  document.querySelectorAll('#tabs a').forEach((a) => a.classList.toggle('active', a.dataset.tab === name));
}
const countDone = (scope, keys) => { const t = state.ticks[scope] || {}; let n = 0; for (const k of keys) if (t[k]) n++; return n; };
const progressHtml = (scope, keys) => `<span class="progress" data-progress="${scope}"><span class="bar"><b></b></span><span class="pct"></span></span>`;
const itemHtml = (scope, key, inner, extraClass = '', id = '') =>
  `<li class="item ${extraClass}" data-scope="${scope}" data-key="${key}" role="checkbox" aria-checked="false" tabindex="0"${id ? ` id="${id}"` : ''}><span class="box">${CHECK_SVG}</span>${inner}</li>`;

function banners() {
  let h = '';
  if (state.dbError) h += `<div class="banner">${esc(state.dbError)}</div>`;
  return h;
}

function renderMeals() {
  setTitle(APP_NAME, false); setTab('meals');
  const m = state.model;
  if (!m) { view().innerHTML = '<p class="loading">Loading recipes…</p>'; return; }
  let h = banners();
  if (!state.firebase) h += '<div class="banner">Sharing is not set up yet, so ticks stay on this phone. See SETUP.md.</div>';
  for (const meal of m.meals) {
    const sides = meal.sides.map((d) => esc(d.name)).join(' &amp; ');
    const time = meal.planTotal ? esc(meal.planTotal.split(/[.,;]/)[0]) : meal.totalMin != null ? `≈ ${fmtDuration(meal.totalMin)}` : '';
    h += `<a class="card meal-card" href="#/meal/${meal.id}">
      <div class="eyebrow">${esc(meal.title)}</div>
      <h2>${esc(meal.main.name)}</h2>
      ${sides ? `<p class="sides">with ${sides}</p>` : ''}
      <div class="meta">${time ? `<span class="time">⏱ ${time}</span>` : ''}${progressHtml(meal.id, meal.keys)}</div>
    </a>`;
  }
  view().innerHTML = h || '<p class="empty">No meals yet.</p>';
  refreshTicks();
}

function renderMeal(id) {
  const meal = state.model && state.model.meals.find((x) => x.id === id);
  if (!meal) { if (state.model) location.replace('#/'); else view().innerHTML = '<p class="loading">Loading…</p>'; return; }
  setTitle(meal.title, true); setTab('meals');
  const S = meal.id;
  let h = '<nav class="jump"><a href="#" data-jump="' + S + '-plan">Plan</a>' +
    meal.dishes.map((d) => `<a href="#" class="chip-${d.color}" data-jump="${d.anchor}">${esc(shortName(d))}</a>`).join('') + '</nav>';
  h += banners();
  h += `<div class="listhead">${progressHtml(S, meal.keys)}</div>`;

  // ---- Plan first
  if (meal.planTotal || meal.planFirst || meal.starts.length > 1) {
    h += `<section class="card plan" id="${S}-plan"><p class="section-label">The plan</p><h2>Total time: ${esc(meal.planTotal || fmtDuration(meal.totalMin))}</h2>`;
    if (meal.planFirst) h += `<div class="startfirst"><strong>⚠ Start first</strong>${esc(meal.planFirst)}</div>`;
    if (meal.starts.length > 1) {
      const mainStart = (meal.starts.find((s) => s.dish === meal.main) || {}).min;
      h += '<p class="section-label">Start order</p><ol class="startorder">' + meal.starts.map((s, i) => {
        const early = i === 0 && s.dish !== meal.main && mainStart != null && s.min < mainStart;
        return `<li class="chip-${s.dish.color}" data-jump="${s.dish.anchor}">${early ? '<span class="first-badge">First</span>' : ''}${esc(s.dish.name)}<span class="at">${s.min === 0 ? 'at 0:00' : 'at ' + Math.floor(s.min / 60) + ':' + String(s.min % 60).padStart(2, '0')}</span></li>`;
      }).join('') + '</ol>';
    }
    h += '</section>';
  }

  // ---- dishes
  for (const d of meal.dishes) {
    h += `<section class="card dish" id="${d.anchor}"><div class="dish-head c${d.color}"><p class="section-label">${d.type}</p><h2>${esc(d.name)}</h2>${d.makes ? `<p class="makes">${esc(d.makes)}</p>` : ''}</div>`;
    if (d.groups.length) {
      h += '<h3>Ingredients</h3>';
      for (const g of d.groups) {
        if (g.name) h += `<h4>${esc(g.name)}</h4>`;
        h += '<ul class="list ings">' + g.items.map((it) =>
          `<li class="ing">${it.amount ? `<span class="amt">${esc(it.amount)}</span> ` : ''}${esc(it.text)}</li>`).join('') + '</ul>';
      }
    }
    if (d.steps.length) {
      h += '<h3>Steps</h3><ul class="list">' + d.steps.map((s) => itemHtml(S, s.key,
        `<span class="body"><span class="num">${esc(s.num)}.</span>${esc(s.text)}${s.amount ? `<br><span class="used">${esc(s.amount)}</span>` : ''}</span>`, '', `${d.anchor}-s${s.num}`)).join('') + '</ul>';
    }
    for (const t of d.tips) h += `<div class="tip"><strong>Tip:</strong> ${esc(t)}</div>`;
    h += '</section>';
  }
  h += `<button class="btn outline-danger" data-clear="${S}" data-label="${esc(meal.title)}">Clear ticks for this meal</button>`;
  view().innerHTML = h;
  refreshTicks();
}
function shortName(d) { // short label for the jump bar
  const n = d.name.replace(/\s+with\s+.*$/i, '');
  return n.length > 26 ? n.slice(0, 24) + '…' : n;
}

function renderShopping() {
  setTitle('Shopping list', false); setTab('shopping');
  const m = state.model;
  if (!m) { view().innerHTML = '<p class="loading">Loading…</p>'; return; }
  let h = banners() + `<div class="listhead">${progressHtml('shopping', m.shopKeys)}</div>`;
  if (!m.shopping.length) h += '<p class="empty">No shopping list found. Add a "Shopping" tab to the spreadsheet.</p>';
  for (const sec of m.shopping) {
    h += `<section class="card"><p class="section-label">${esc(sec.name)}</p><ul class="list">` + sec.items.map((x) => itemHtml('shopping', x.key,
      `<span class="body"><strong>${esc(x.item)}</strong>${x.amount ? ` · ${esc(x.amount)}` : ''}${x.usedIn ? `<span class="sub">For: ${esc(x.usedIn)}</span>` : ''}</span>`)).join('') + '</ul></section>';
  }
  if (m.shopping.length) h += '<button class="btn outline-danger" data-clear="shopping" data-label="the shopping list">Clear shopping ticks</button>';
  view().innerHTML = h;
  refreshTicks();
}

function renderSettings() {
  setTitle('Settings', false); setTab('settings');
  const info = state.recipeInfo || {};
  const when = info.at ? new Date(info.at).toLocaleString() : 'never';
  let h = banners();
  h += '<section class="card"><p class="section-label">Sharing</p>';
  if (!state.firebase) h += '<p>Not set up yet. Ticks are saved on this phone only. Follow SETUP.md steps 2–5 to share ticks between phones.</p>';
  else if (state.user) h += `<p>Signed in as <strong>${esc(state.user.email)}</strong>.</p><p class="small">Status: ${esc($('#sync').textContent)}</p><button class="btn ghost" id="signout">Sign out</button>`;
  h += '</section>';

  h += `<section class="card"><p class="section-label">Recipes</p>
    <p>Source: <strong>${info.source === 'sheet' ? 'Google Sheet' : 'Built-in file (recipes.xlsx)'}</strong><br>
    <span class="small">Last loaded: ${esc(when)} · ${state.model ? state.model.meals.length + ' meals' : 'none'}</span></p>
    ${info.error ? `<p class="err">${esc(info.error)}</p>` : ''}
    <form id="sheetform">
      <label class="field">Google Sheet link${state.firebase ? ' (shared with both phones)' : ''}
        <input type="url" id="sheeturl" placeholder="https://docs.google.com/spreadsheets/…" value="${esc(state.sheetUrl)}">
      </label>
      <div class="row-btns" style="justify-content:flex-start;margin-top:0">
        <button class="btn" type="submit">Save link</button>
        <button class="btn ghost" type="button" id="reload">Reload recipes</button>
      </div>
      <p class="small">Leave the link empty to use the built-in recipes.xlsx file.</p>
    </form></section>`;

  h += `<section class="card"><p class="section-label">While cooking</p>
    <label class="switch"><span>Keep screen on in a meal</span><input type="checkbox" id="wake" ${ls.get('wake', true) ? 'checked' : ''}></label></section>`;
  h += `<section class="card"><p class="section-label">App updates</p>
    <p class="small">After uploading a new version to GitHub, wait 2 minutes, then tap this.</p>
    <button class="btn ghost" type="button" id="update">Check for app update</button></section>`;
  h += `<p class="small" style="text-align:center">${esc(APP_NAME)} · version ${esc(window.APP_VERSION || 'dev')}</p>`;
  view().innerHTML = h;
}

function renderLogin(err = '') {
  setTitle(APP_NAME, false);
  $('#tabs').hidden = true;
  view().innerHTML = `<section class="card login"><h2>Sign in</h2>
    <p class="small">Use the email and password you set up in Firebase (SETUP.md step 2). You only do this once per phone.</p>
    <form id="loginform">
      <label class="field">Email<input type="email" id="email" autocomplete="username" required value="${esc(ls.get('last-email', ''))}"></label>
      <label class="field">Password<input type="password" id="password" autocomplete="current-password" required></label>
      ${err ? `<p class="err">${esc(err)}</p>` : ''}
      <button class="btn block" type="submit">Sign in</button>
    </form>
    <p style="text-align:center"><a href="#" id="forgot">Forgot password?</a></p></section>`;
}

// ---------------------------------------------------------------- tick rendering
function refreshTicks() {
  document.querySelectorAll('.item[data-key]').forEach((el) => {
    const on = !!(state.ticks[el.dataset.scope] || {})[el.dataset.key];
    el.classList.toggle('done', on); el.setAttribute('aria-checked', on);
  });
  // ticked items move to the bottom of their list, unticked ones keep their original order
  document.querySelectorAll('ul.list').forEach((ul) => {
    const items = [...ul.children].filter((li) => li.dataset.key);
    if (!items.length) return;
    items.forEach((li, i) => { if (li.dataset.i == null) li.dataset.i = i; });
    const want = [...items].sort((a, b) => a.classList.contains('done') - b.classList.contains('done') || a.dataset.i - b.dataset.i);
    if (want.some((li, i) => li !== items[i])) want.forEach((li) => ul.appendChild(li));
  });
  document.querySelectorAll('[data-progress]').forEach((el) => {
    const scope = el.dataset.progress, m = state.model; if (!m) return;
    const keys = scope === 'shopping' ? m.shopKeys : (m.meals.find((x) => x.id === scope) || {}).keys || [];
    const done = countDone(scope, keys);
    el.querySelector('b').style.width = keys.length ? (100 * done / keys.length) + '%' : 0;
    el.querySelector('.pct').textContent = `${done} / ${keys.length}`;
  });
}

// ---------------------------------------------------------------- events
document.addEventListener('click', async (e) => {
  const jump = e.target.closest('[data-jump]');
  if (jump) {
    e.preventDefault();
    const target = document.getElementById(jump.dataset.jump) || document.getElementById(jump.dataset.jump.replace(/-s\d+$/, ''));
    if (target) {
      const y = target.getBoundingClientRect().top + scrollY - 120;
      scrollTo({ top: y, behavior: 'smooth' });
      if (target.classList.contains('item')) { target.classList.remove('flash'); void target.offsetWidth; target.classList.add('flash'); }
    }
    return;
  }
  const item = e.target.closest('.item[data-key]');
  if (item) { setTick(item.dataset.scope, item.dataset.key, !item.classList.contains('done')); return; }
  const clr = e.target.closest('[data-clear]');
  if (clr) {
    if (await confirmDialog(`Clear all ticks for ${clr.dataset.label}? This clears them on both phones.`)) { clearScope(clr.dataset.clear); toast('Ticks cleared'); }
    return;
  }
  if (e.target.closest('#back')) { if (navCount > 1) history.back(); else location.hash = '#/'; return; }
  if (e.target.id === 'signout') { stopFirestore(); await signOut(state.firebase.auth); return; }
  if (e.target.id === 'reload') { toast('Loading recipes…'); await loadRecipes(); toast(state.recipeInfo && state.recipeInfo.error ? 'Could not load, see the message' : 'Recipes loaded'); route(); return; }
  if (e.target.id === 'update') {
    toast('Checking for an update…');
    try {
      const r = await fetch('sw.js', { cache: 'no-store' });
      const v = ((await r.text()).match(/VERSION = '([^']+)'/) || [])[1];
      if (!v) throw new Error('no version found');
      if (v === window.APP_VERSION) { toast('You have the latest version (' + v + ')'); return; }
      const reg = await navigator.serviceWorker.getRegistration();
      if (reg) {
        await reg.update();
        const sw = reg.installing || reg.waiting;
        if (sw && sw.state !== 'activated') await new Promise((res) => { sw.addEventListener('statechange', () => sw.state === 'activated' && res()); setTimeout(res, 15000); });
      }
      toast('Updating to ' + v + '…'); setTimeout(() => location.reload(), 600);
    } catch (err) { toast('Could not check: ' + err.message); }
    return;
  }
  if (e.target.id === 'forgot') {
    e.preventDefault();
    const email = $('#email').value.trim();
    if (!email) { toast('Type your email first'); return; }
    try { await sendPasswordResetEmail(state.firebase.auth, email); toast('Reset email sent'); } catch (err) { toast('Could not send: ' + (err.code || err.message)); }
  }
});
document.addEventListener('keydown', (e) => {
  if ((e.key === ' ' || e.key === 'Enter') && e.target.matches('.item[data-key]')) { e.preventDefault(); e.target.click(); }
});
document.addEventListener('submit', async (e) => {
  if (e.target.id === 'loginform') {
    e.preventDefault();
    const email = $('#email').value.trim(), pw = $('#password').value;
    ls.set('last-email', email);
    const btn = e.target.querySelector('button'); btn.disabled = true; btn.textContent = 'Signing in…';
    try { await signInWithEmailAndPassword(state.firebase.auth, email, pw); } catch (err) {
      const msg = { 'auth/invalid-credential': 'Wrong email or password.', 'auth/network-request-failed': 'No connection. Sign-in needs internet the first time.', 'auth/too-many-requests': 'Too many tries. Wait a few minutes.' }[err.code] || (err.code || err.message);
      renderLogin(msg);
    }
  }
  if (e.target.id === 'sheetform') {
    e.preventDefault();
    const url = $('#sheeturl').value.trim();
    if (url && !sheetSources(url)) { toast('That is not a Google Sheets link'); return; }
    state.sheetUrl = url; ls.set('sheet-url', url);
    if (state.firebase) setDoc(doc(state.firebase.db, 'config', 'app'), { sheetUrl: url }, { merge: true }).catch((err) => toast('Could not save: ' + (err.code || err.message)));
    toast('Saved. Loading recipes…');
    await loadRecipes(); route();
  }
});
document.addEventListener('change', (e) => { if (e.target.id === 'wake') { ls.set('wake', e.target.checked); updateWakeLock(); } });

// keep screen on while a meal is open
let wakeLock = null;
async function updateWakeLock() {
  const want = ls.get('wake', true) && location.hash.startsWith('#/meal/') && document.visibilityState === 'visible';
  try {
    if (want && !wakeLock && 'wakeLock' in navigator) { wakeLock = await navigator.wakeLock.request('screen'); wakeLock.addEventListener('release', () => { wakeLock = null; }); }
    else if (!want && wakeLock) { await wakeLock.release(); wakeLock = null; }
  } catch { wakeLock = null; }
}
document.addEventListener('visibilitychange', () => { updateWakeLock(); if (document.visibilityState === 'visible' && state.sheetUrl) loadRecipes(); });

// ---------------------------------------------------------------- router
let lastRoute = '', navCount = 0;
function route() {
  if (state.firebase && !state.authReady) { view().innerHTML = '<p class="loading">Loading…</p>'; return; }
  if (state.firebase && !state.user) { renderLogin(); return; }
  $('#tabs').hidden = false;
  const hash = location.hash.replace(/^#\/?/, '');
  const [page, arg] = hash.split('/');
  const keepScroll = hash === lastRoute;
  const y = scrollY;
  if (page === 'meal') renderMeal(decodeURIComponent(arg || ''));
  else if (page === 'shopping') renderShopping();
  else if (page === 'settings') renderSettings();
  else renderMeals();
  if (keepScroll) scrollTo(0, y); else scrollTo(0, 0);
  if (hash !== lastRoute) navCount++;
  lastRoute = hash;
  updateWakeLock();
}
addEventListener('hashchange', route);

// ---------------------------------------------------------------- start
function start() {
  state.sheetUrl = ls.get('sheet-url', '');
  const cached = ls.get('recipes-cache', null);
  if (cached) { try { state.model = buildModel(cached); state.rowsSig = JSON.stringify(cached); } catch { /* ignore bad cache */ } }

  if (fbConfig) {
    const app = initializeApp(fbConfig);
    const auth = initializeAuth(app, { persistence: [indexedDBLocalPersistence, browserLocalPersistence] });
    const db = initializeFirestore(app, { localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }) });
    if (CFG.emulator) { // local testing only
      connectAuthEmulator(auth, 'http://127.0.0.1:9099', { disableWarnings: true });
      connectFirestoreEmulator(db, '127.0.0.1', 8080);
    }
    state.firebase = { auth, db };
    onAuthStateChanged(auth, (user) => {
      state.user = user; state.authReady = true; state.dbError = '';
      if (user) startFirestore(); else { stopFirestore(); state.ticks = {}; }
      updateSync(true); route();
    });
  } else {
    state.ticks = ls.get('local-ticks', {});
    state.sync = 'local'; renderSync();
  }
  route();
  loadRecipes();
  if ('serviceWorker' in navigator && location.protocol !== 'file:') {
    navigator.serviceWorker.register('sw.js').catch((e) => console.warn('SW failed', e));
  }
}
start();
