// What's Cooking with Brian – shared cooking checklist PWA.
// Source file. The build step bundles this into ../app.js (see build.mjs).

import XLSX from 'xlsx/dist/xlsx.mini.min.js';
import { initializeApp } from 'firebase/app';
import {
  initializeAuth, indexedDBLocalPersistence, browserLocalPersistence,
  signInWithEmailAndPassword, sendPasswordResetEmail, onAuthStateChanged, signOut, connectAuthEmulator,
} from 'firebase/auth';
import {
  initializeFirestore, persistentLocalCache, persistentMultipleTabManager, terminate, clearIndexedDbPersistence,
  collection, doc, onSnapshot, setDoc, deleteDoc, deleteField, FieldPath, serverTimestamp, connectFirestoreEmulator,
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
// Color themes: chosen per phone in Settings (stored on the phone only). index.html applies it before the page draws.
const THEMES = [['teal', 'Teal', '#1f6e68'], ['rust', 'Rust', '#b8472a'], ['blueberry', 'Blueberry', '#2f4f8f'], ['plum', 'Plum', '#7a3b6e'], ['forest', 'Forest', '#2f6b3a'], ['slate', 'Slate', '#44525e'], ['pink', 'Hot Pink', '#d6336c'], ['violet', 'Violet', '#7c3aed'], ['ocean', 'Ocean', '#0b78a8']];
function applyColor() {
  const c = ls.get('color-theme', 'teal');
  if (c && c !== 'teal') document.documentElement.dataset.color = c; else delete document.documentElement.dataset.color;
  const bar = document.querySelector('.topbar'), meta = document.querySelector('meta[name="theme-color"]');
  if (bar && meta) meta.content = getComputedStyle(bar).backgroundColor; // phone status bar matches the header
}
try { matchMedia('(prefers-color-scheme: dark)').addEventListener('change', applyColor); } catch { /* old browser */ }
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
let undoFn = null;
function toast(msg, ms = 2600) {
  const t = $('#toast'); t.textContent = msg; t.classList.remove('hasundo'); undoFn = null; t.classList.add('show');
  clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.remove('show'), ms);
}
// Deleted straight away, with a few seconds to undo it (instead of an "are you sure?" question).
function undoToast(msg, fn) {
  const t = $('#toast'); t.innerHTML = `<span>${esc(msg)}</span><button type="button" class="toastundo">Undo</button>`;
  undoFn = fn; t.classList.add('show', 'hasundo');
  clearTimeout(toastTimer); toastTimer = setTimeout(() => { t.classList.remove('show', 'hasundo'); undoFn = null; }, 6000);
}
document.addEventListener('click', (e) => {
  if (!e.target.closest('.toastundo')) return;
  const f = undoFn; undoFn = null; $('#toast').classList.remove('show', 'hasundo'); if (f) { f(); toast('Restored'); }
});
function confirmDialog(text, okLabel = 'Clear') {
  const dlg = $('#confirm');
  $('#confirm-text').textContent = text; $('#confirm-ok').textContent = okLabel;
  return new Promise((resolve) => {
    dlg.addEventListener('close', () => resolve(dlg.returnValue === 'ok'), { once: true });
    dlg.returnValue = ''; dlg.showModal();
  });
}
// Small line icons in the tab-bar style; they take the color of the text around them (so they follow the theme).
const IC = (d) => `<svg class="ic" viewBox="0 0 24 24" aria-hidden="true"><path d="${d}" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
const ICONS = {
  cal: IC('M4.5 6.5h15v13h-15zM4.5 10.5h15M8.5 4v4M15.5 4v4'),
  timer: IC('M12 21a7.5 7.5 0 1 0 0-15 7.5 7.5 0 0 0 0 15zM12 10v3.5l2.5 1.5M9.5 3h5'),
  cook: IC('M7.5 17.5h9v3h-9zM7.5 17.5v-4A4 4 0 0 1 8.4 6a4 4 0 0 1 7.2 0 4 4 0 0 1 .9 7.5v4'),
  camera: IC('M4 8.5h3.2l1.8-2.5h6l1.8 2.5H20V19H4zM12 16.5a3.2 3.2 0 1 0 0-6.4 3.2 3.2 0 0 0 0 6.4z'),
  warn: IC('M12 4 2.8 19.5h18.4zM12 10v4.2M12 17v.01'),
  bell: IC('M6.5 16.5V11a5.5 5.5 0 0 1 11 0v5.5l1.5 2h-14zM10 21h4'),
  note: IC('M6 3.5h8.5L19 8v12.5H6zM14 3.5V8h5M9 12h7M9 16h5'),
  list: IC('M9 6.5h11M9 12h11M9 17.5h11M4.5 6.5v.01M4.5 12v.01M4.5 17.5v.01'),
};
const CHECK_SVG = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 12.5l4.2 4.2L19 7" fill="none" stroke="#fff" stroke-width="3.2" stroke-linecap="round" stroke-linejoin="round"/></svg>';

// ---------------------------------------------------------------- recipe parsing
function rowsFromWorkbook(wb) {
  const find = (name) => wb.SheetNames.find((n) => n.trim().toLowerCase() === name);
  const toRows = (n) => (n ? XLSX.utils.sheet_to_json(wb.Sheets[n], { defval: '', raw: false }) : []);
  const recipesName = find('recipes') || wb.SheetNames[0];
  return { recipes: toRows(recipesName), shopping: toRows(find('shopping')), archive: toRows(find('archive')) };
}
function keyed(row) { // lower-case, trimmed column names
  const o = {}; for (const k in row) o[k.trim().toLowerCase()] = norm(row[k]); return o;
}

function buildModel({ recipes, shopping, archive = [] }) {
  const current = buildMeals(recipes);
  const names = new Set(current.map((x) => x.name));
  const past = buildMeals(archive).filter((x) => !names.has(x.name)); // a meal in both tabs counts as current
  past.forEach((x) => { x.archived = true; });
  return { meals: current, past, ...buildShopping(shopping) };
}
// Shopping details on an ingredient row (columns Shop item / Shop qty / Shop unit / Shop section).
// null = not filled in yet; { none: true } = never bought (e.g. water).
function shopInfo(r) {
  const item = r['shop item'];
  if (!item) return null;
  if (/^\(?none\)?$/i.test(item)) return { none: true };
  const q = parseFloat(String(r['shop qty'] || '').replace(',', '.'));
  return { item, qty: isNaN(q) ? null : q, unit: (r['shop unit'] || '').trim(), section: r['shop section'] || 'Other' };
}
function buildMeals(recipes) {
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
    if (!dish) { dish = { name: recipe, id: '', type: /side/i.test(r.type) ? 'Side' : /dessert/i.test(r.type) ? 'Dessert' : 'Main', makes: '', groups: [], steps: [], tips: [], seq }; meal.dishes.set(recipe, dish); }
    if (!dish.id && r.id) dish.id = r.id;
    if (/^(yes|y|true)$/i.test(r['hide from recipe book'] || '')) dish.hideBook = true; // simple sides kept out of the Recipe Book
    const order = r.order === '' ? NaN : parseFloat(r.order);
    switch (section) {
      case 'makes': dish.makes = r.text; break;
      case 'plan total': meal.planTotal = r.text; break;
      case 'plan first': meal.planFirst = r.text; break;
      case 'plan': meal.plan.push({ when: normTime(r.amount), order, text: r.text, seq }); break;
      case 'ingredient': {
        let g = dish.groups.find((x) => x.name === r.group);
        if (!g) { g = { name: r.group, items: [] }; dish.groups.push(g); }
        g.items.push({ amount: r.amount, text: r.text, shop: shopInfo(r) }); break;
      }
      case 'step': dish.steps.push({ order, amount: r.amount, text: r.text, timer: r.timer || '', prep: r['prep task'] || '', prepAmt: r['prep amount'] || '', combo: r['combined order'] === '' || r['combined order'] == null ? NaN : parseFloat(r['combined order']), ahead: r.ahead || '', seq }); break;
      case 'tip': if (r.text) dish.tips.push(r.text); break;
      case 'ahead': if (r.text) (meal.aheadNotes || (meal.aheadNotes = [])).push({ when: r.amount || '', text: r.text }); break;
      case 'made': dish.madeBefore = Math.max(0, parseInt(r.amount, 10) || 0); dish.madeNote = r.text || ''; break;
      default: break;
    }
  }
  const byOrder = (a, b) => (isNaN(a.order) ? 1e9 : a.order) - (isNaN(b.order) ? 1e9 : b.order) || a.seq - b.seq;

  const out = [];
  for (const meal of meals.values()) {
    const ORDER = { Main: 0, Side: 1, Dessert: 2 };
    const dishes = [...meal.dishes.values()].sort((a, b) => ORDER[a.type] - ORDER[b.type] || a.seq - b.seq);
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
    meal.kind = meal.main.type; // 'Main' or 'Side' (past entries are one dish each)
    meal.rid = meal.main.id || 'd-' + slug(meal.main.name); // ratings and calendar entries are keyed on the main dish's ID
    // A meal named after its dish (Recipe Book entries) gets a short title, like "Chicken Pot Pie" for a planned meal:
    // "Apple Muffins with Cinnamon Crumb" → "Apple Muffins". The dish keeps its full name.
    if (meal.title === meal.main.name) meal.title = meal.title.replace(/\s+with\s+.*$/i, '') || meal.title;
    const mins = meal.plan.map((p) => p.min).filter((m) => m != null);
    meal.totalMin = mins.length ? Math.max(...mins) : null;
    meal.keys = dishes.flatMap((d) => d.steps.map((s) => s.key));
    // Combined plan: every step of every dish in cooking order (the spreadsheet's "Combined order" column).
    // The same steps (same ticks) as the dish-by-dish view; steps without a number go last, dish by dish.
    const all = dishes.flatMap((d, di) => d.steps.map((st) => ({ d, s: st, di })));
    meal.combined = all.some((x) => !isNaN(x.s.combo)) && dishes.length > 1
      ? all.sort((a, b) => (isNaN(a.s.combo) ? 1e9 : a.s.combo) - (isNaN(b.s.combo) ? 1e9 : b.s.combo) || a.di - b.di || a.s.num - b.s.num) : null;
    computeStarts(meal);
    out.push(meal);
  }
  out.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));
  return out;
}
function buildShopping(shopping) {
  const sections = new Map();
  let list = '';
  const seen = new Map();
  for (const raw of shopping) {
    const r = keyed(raw); if (!r.item) continue;
    if (!list && r.list) list = r.list;
    const sec = r.section || 'Other';
    if (!sections.has(sec)) sections.set(sec, []);
    const base = `shop|${r.item}`; const n = (seen.get(base) || 0) + 1; seen.set(base, n);
    sections.get(sec).push({ item: r.item, amount: r.amount, usedIn: r['used in'], key: hash(n > 1 ? `${base}#${n}` : base) });
  }
  const pref = ['meat', 'produce', 'dairy', 'pantry'];
  const rank = (s) => { if (s.toLowerCase() === 'other') return -1; const i = pref.indexOf(s.toLowerCase()); return i < 0 ? 99 : i; }; // Other first
  const shop = [...sections.entries()].map(([name, items], i) => ({ name, items, i })).sort((a, b) => rank(a.name) - rank(b.name) || a.i - b.i);
  return { shopping: shop, shopKeys: shop.flatMap((s) => s.items.map((x) => x.key)), shopList: list };
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
  firebase: null,          // { auth, db } when configured
  authReady: false,
  dbError: '',
  plan: {},
  planX: {},               // calendar/plan .x: day -> { breakfast, lunch, dessert } (dinner is in plan)                // 'YYYY-MM-DD' -> meal name (shared calendar)
  ratings: {},             // meal id -> { name, main, sides, votes: { userId: 1-5 } }
  notes: {},               // shopping notes: id -> { text, by, at }
  thumbs: {},              // photos/_thumbs: dish ID -> tiny JPEG for the Recipe Book list
  photos: {},              // photos/<dish ID>: dish ID -> shrunk JPEG (data URL), or null when there's none
  nextTime: {},            // notes/nexttime: main dish ID -> { text, email, at } (note for next time)
  madeAdjust: {},          // history/adjust: dish ID -> change (+/-) to the counted times made
  ideas: {},               // Planning tab: recipe ideas, id -> { text, email, at, sent }
  session: { dinner: {}, timers: {} }, // cooking/session: dinner times and running timers
  staples: {},             // shopping/staples: items we usually have, by name key
  bought: {},              // shopping/bought: meal ID -> { since: 'YYYY-MM-DD', i: { shop item key: true } }
  myName: '',              // this person's name (Settings)
  names: {},               // email (lower-case) -> name, from Firestore profiles/*; used to label notes and added items
  extras: {},              // items added to the shopping list in the app: id -> { name, amount, email, at }
  calWeek: 0,              // calendar page: 0 = this week
};
const CFG = window.APP_CONFIG || {};
const fbConfig = CFG.firebase && CFG.firebase.apiKey && CFG.firebase.projectId ? CFG.firebase : null;

// ---------------------------------------------------------------- ticks store
let unsubSession = null, unsubStaples = null, unsubBought = null, unsubIdeas = null, unsubAdjust = null, unsubNext = null, unsubThumbs = null, thumbsLoaded = false, boughtLoaded = false, ticksLoaded = false, extrasLoaded = false, unsubTicks = null, unsubPlan = null, unsubRatings = null, unsubNotes = null, unsubProfile = null, unsubExtras = null, lastSnapPending = false;

// The shopping page fills at least the screen, so its buttons sit just above the tabs when the list is short
// (and after the last item, as usual, when it's long).
function fitShopPage() {
  const bar = document.querySelector('.topbar'), pad = parseFloat(getComputedStyle(document.body).paddingBottom) || 0;
  view().style.minHeight = `calc(100dvh - ${Math.round((bar ? bar.offsetHeight : 56) + pad)}px)`;
}
addEventListener('resize', () => { if (document.querySelector('.shopfoot')) fitShopPage(); });
// Tick everything on the shopping list at once (staples stay as they are), with Undo.
function markAllBought() {
  const before = JSON.parse(JSON.stringify(state.bought)), today = dayKey(new Date()), ticked = [];
  for (const l of shopLines()) {
    if (l.done || isStaple(l.item)) continue;
    for (const x of l.open) {
      const old = state.bought[x.rid], cur = !old || cooked(x.rid) ? { since: today, i: {} } : { since: old.since || today, i: { ...(old.i || {}) } };
      cur.i[l.key] = true; state.bought[x.rid] = cur;
    }
  }
  const t = state.ticks[SHOP()] || {};
  for (const [id, x] of Object.entries(state.extras)) if (!isStaple(x.name) && !t[extraKey(id)]) ticked.push(extraKey(id));
  const saveBought = () => {
    if (state.firebase) setDoc(doc(state.firebase.db, 'shopping', 'bought'), { b: state.bought }).catch((err) => toast('Could not save: ' + (err.code || err.message)));
    else ls.set('local-bought', state.bought);
  };
  saveBought(); ticked.forEach((k) => setTick(SHOP(), k, true)); route();
  undoToast('Everything marked as bought', () => { state.bought = before; saveBought(); ticked.forEach((k) => setTick(SHOP(), k, false)); route(); });
}
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
// Shopping ticks are saved per list (the Shopping tab's List date), so a new weekly list starts unticked.
const SHOP = () => (autoShop() ? 'shopping-auto' : state.model && state.model.shopList ? 'shopping-' + state.model.shopList : 'shopping');
const isShopScope = (scope) => scope === 'shopping' || scope.startsWith('shopping-');
async function clearScope(scope, k = {}) { // k: the ticks to put back (for Undo); empty clears them
  if (!state.firebase) { state.ticks[scope] = { ...k }; ls.set('local-ticks', state.ticks); refreshTicks(); return; }
  state.ticks[scope] = { ...k }; refreshTicks();
  setDoc(doc(state.firebase.db, 'ticks', scope), { k: { ...k }, at: serverTimestamp() })
    .catch((e) => toast('Could not clear: ' + (e.code || e.message)));
}
function startFirestore() {
  const { db } = state.firebase;
  stopFirestore();
  unsubTicks = onSnapshot(collection(db, 'ticks'), { includeMetadataChanges: true }, safely('ticks', (snap) => {
    const ticks = {};
    snap.forEach((d) => { ticks[d.id] = { ...(d.data().k || {}) }; });
    state.ticks = ticks; state.dbError = '';
    lastSnapPending = snap.metadata.hasPendingWrites;
    updateSync(snap.metadata.fromCache);
    if (!snap.metadata.fromCache) { ticksLoaded = true; tidyOldLists(); }
    refreshTicks();
  }), (err) => {
    console.error(err);
    state.dbError = err.code === 'permission-denied'
      ? 'This account is not allowed to read the shared ticks. Check the two email addresses in the Firestore rules (SETUP.md step 4).'
      : 'Sync error: ' + (err.code || err.message);
    state.sync = 'error'; renderSync(); route();
  });
  unsubPlan = onSnapshot(doc(db, 'calendar', 'plan'), { includeMetadataChanges: true }, safely('calendar', (snap) => {
    state.plan = { ...((snap.exists() && snap.data().d) || {}) }; syncLoaded('calendar', snap);
    state.planX = { ...((snap.exists() && snap.data().x) || {}) }; // breakfast / lunch / dessert per day
    if (changed('calendar', [state.plan, state.planX]) && /^#?\/?(calendar|past|planning)?$/.test(location.hash)) route(); // re-draw Calendar, Meals or Recipe Book
  }), (err) => listenerFailed('calendar', err));
  unsubRatings = onSnapshot(collection(db, 'ratings'), { includeMetadataChanges: true }, safely('ratings', (snap) => {
    const r = {}; snap.forEach((d) => { r[d.id] = d.data(); });
    state.ratings = r; syncLoaded('ratings', snap);
    if (changed('ratings', r) && /^#?\/?(calendar|past|planning|meal\/.*)?$/.test(location.hash)) route();
  }), (err) => listenerFailed('ratings', err));
  unsubNotes = onSnapshot(doc(db, 'notes', 'shopping'), { includeMetadataChanges: true }, safely('notes', (snap) => {
    state.notes = { ...((snap.exists() && snap.data().n) || {}) }; syncLoaded('notes', snap);
    if (changed('notes', state.notes)) renderNotes();
  }), (err) => listenerFailed('notes', err));
  // everyone's names, so notes and added items show the latest name for each email
  unsubProfile = onSnapshot(collection(db, 'profiles'), { includeMetadataChanges: true }, safely('names', (snap) => {
    const names = {};
    snap.forEach((d) => {
      const x = d.data();
      if (x.email && x.name) names[String(x.email).toLowerCase()] = x.name;
      if (d.id === state.user.uid) state.myName = x.name || '';
    });
    state.names = names; syncLoaded('names', snap);
    const inp = $('#myname'); if (inp && document.activeElement !== inp) inp.value = state.myName;
    if (changed('names', names)) { renderNotes(); renderExtras(); }
  }), (err) => listenerFailed('names', err));
  unsubExtras = onSnapshot(doc(db, 'shopping', 'extra'), { includeMetadataChanges: true }, safely('added items', (snap) => {
    state.extras = { ...((snap.exists() && snap.data().i) || {}) }; syncLoaded('added items', snap);
    if (!snap.metadata.fromCache) { extrasLoaded = true; tidyOldLists(); }
    if (changed('added items', state.extras)) renderExtras();
  }), (err) => listenerFailed('added items', err));
  // dinner times, running timers and who's making which dish (one shared record)
  unsubSession = onSnapshot(doc(db, 'cooking', 'session'), { includeMetadataChanges: true }, safely('cooking', (snap) => {
    const d = (snap.exists() && snap.data()) || {};
    state.session = { dinner: { ...(d.dinner || {}) }, timers: { ...(d.timers || {}) } }; syncLoaded('cooking', snap);
    renderTimers();
  }), (err) => listenerFailed('cooking', err));
  unsubStaples = onSnapshot(doc(db, 'shopping', 'staples'), { includeMetadataChanges: true }, safely('staples', (snap) => {
    state.staples = { ...((snap.exists() && snap.data().s) || {}) }; syncLoaded('staples', snap);
    if (changed('staples', state.staples)) renderExtras();
  }), (err) => listenerFailed('staples', err));
  unsubBought = onSnapshot(doc(db, 'shopping', 'bought'), { includeMetadataChanges: true }, safely('bought', (snap) => {
    state.bought = { ...((snap.exists() && snap.data().b) || {}) }; syncLoaded('bought', snap);
    if (!snap.metadata.fromCache) { boughtLoaded = true; tidyBought(); }
    if (changed('bought', state.bought)) renderExtras();
  }), (err) => listenerFailed('bought', err));
  unsubIdeas = onSnapshot(doc(db, 'ideas', 'list'), { includeMetadataChanges: true }, safely('ideas', (snap) => {
    state.ideas = { ...((snap.exists() && snap.data().i) || {}) }; syncLoaded('ideas', snap);
    if (changed('ideas', state.ideas) && /^#\/ideas/.test(location.hash)) route();
  }), (err) => listenerFailed('ideas', err));
  unsubThumbs = onSnapshot(doc(db, 'photos', '_thumbs'), { includeMetadataChanges: true }, safely('photos', (snap) => {
    state.thumbs = { ...((snap.exists() && snap.data().t) || {}) }; if (!snap.metadata.fromCache) thumbsLoaded = true;
    if (changed('thumbs', state.thumbs) && /^#?\/?(past|planning)?$/.test(location.hash)) route();
  }), (err) => listenerFailed('photos', err));
  unsubNext = onSnapshot(doc(db, 'notes', 'nexttime'), { includeMetadataChanges: true }, safely('next-time notes', (snap) => {
    state.nextTime = { ...((snap.exists() && snap.data().n) || {}) }; syncLoaded('next-time notes', snap);
    if (changed('next-time notes', state.nextTime) && /^#\/meal\//.test(location.hash)) route();
  }), (err) => listenerFailed('next-time notes', err));
  unsubAdjust = onSnapshot(doc(db, 'history', 'adjust'), { includeMetadataChanges: true }, safely('times made', (snap) => {
    state.madeAdjust = { ...((snap.exists() && snap.data().a) || {}) }; syncLoaded('times made', snap);
    if (changed('times made', state.madeAdjust) && /^#?\/?(past|planning)?$/.test(location.hash)) route();
  }), (err) => listenerFailed('times made', err));
}
// Redraw only when a list's contents change, not when only its "confirmed by server" status changes.
const lastSeen = {};
function changed(what, value) { const sig = JSON.stringify(value); if (lastSeen[what] === sig) return false; lastSeen[what] = sig; return true; }
// Runs a snapshot handler; if it throws, the error is reported instead of silently stopping later updates.
const safely = (what, fn) => (snap) => {
  try { fn(snap); } catch (e) {
    console.error('error handling', what, e);
    syncInfo[what] = { ok: false, error: 'app error: ' + (e && e.message), at: Date.now() };
    state.listenErr = { what, code: 'app error: ' + (e && e.message) };
    toast(`Problem showing ${what}: ${e && e.message}`, 8000);
  }
};
// Sync health: what each shared list last loaded, and any error (shown in Settings and as a banner).
const syncInfo = {};
function syncLoaded(what, snap) { syncInfo[what] = { ok: true, server: !snap.metadata.fromCache, at: Date.now() }; if (state.listenErr && state.listenErr.what === what) { state.listenErr = null; route(); } }
function listenerFailed(what, err) {
  console.error('sync failed:', what, err);
  syncInfo[what] = { ok: false, error: err.code || err.message, at: Date.now() };
  state.listenErr = { what, code: err.code || err.message };
  toast(`Couldn't load ${what}: ${err.code || err.message}`, 6000);
  route();
}
// Throw away this phone's saved copy of the shared data and download it all again from Firebase.
// (Uninstalling the app on Android doesn't clear this copy: it belongs to Chrome.)
async function resetLocalData() {
  if (!state.firebase) return;
  if (!(await confirmDialog('Reload all shared data on this phone from Firebase? Anything ticked or added while offline that hasn\'t synced yet will be lost.', 'Reload'))) return;
  toast('Reloading data…');
  try {
    stopFirestore();
    // shutting down and clearing can wait forever while another copy of the app holds the data: give up after 6 seconds
    const busy = new Promise((_, no) => setTimeout(() => no(Object.assign(new Error('busy'), { code: 'failed-precondition' })), 6000));
    await Promise.race([terminate(state.firebase.db), busy]);
    await Promise.race([clearIndexedDbPersistence(state.firebase.db), busy]);
    location.reload();
  } catch (e) {
    toast(e.code === 'failed-precondition' ? 'Close any other copy of the app (for example a Chrome tab), then try again.' : 'Could not reset: ' + (e.code || e.message), 8000);
    setTimeout(() => location.reload(), 8000);
  }
}
function stopFirestore() { for (const k in lastSeen) delete lastSeen[k]; for (const k in photoUnsubs) { photoUnsubs[k](); delete photoUnsubs[k]; } for (const u of [unsubTicks, unsubPlan, unsubRatings, unsubNotes, unsubProfile, unsubExtras, unsubSession, unsubStaples, unsubBought, unsubIdeas, unsubAdjust, unsubThumbs, unsubNext]) if (u) u(); unsubTicks = unsubPlan = unsubRatings = unsubNotes = unsubProfile = unsubExtras = unsubSession = unsubStaples = unsubBought = unsubIdeas = unsubAdjust = unsubThumbs = unsubNext = null; }

// ---------------------------------------------------------------- Planning: recipe ideas
// Stored in Firestore ideas/list as { i: { id: { text, email, at, sent } } } (one document, so one read per app open).
// "Copy for recipe Claude" copies the unsent ideas as a ready-to-paste message, then marks them sent.
function writeIdeas(changes) {
  for (const [id, v] of Object.entries(changes)) { if (v) state.ideas[id] = v; else delete state.ideas[id]; }
  if (/^#\/ideas/.test(location.hash)) route();
  if (!state.firebase) { ls.set('local-ideas', state.ideas); return; }
  const i = {}; for (const [id, v] of Object.entries(changes)) i[id] = v || deleteField();
  setDoc(doc(state.firebase.db, 'ideas', 'list'), { i }, { merge: true }).catch((e) => toast('Could not save idea: ' + (e.code || e.message)));
}
const IDEA_MAX = 1000;
const cleanIdea = (v) => String(v || '').replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n').trim().slice(0, IDEA_MAX);
function linkify(text) { // escaped text with web addresses turned into links that open outside the app
  return esc(text).replace(/https?:\/\/[^\s<]+/g, (u) => {
    const trail = (u.match(/[).,!?;:]+$/) || [''])[0]; const url = u.slice(0, u.length - trail.length);
    let label = url; try { const x = new URL(url.replace(/&amp;/g, '&')); label = esc(x.hostname.replace(/^www\./, '') + (x.pathname.length > 1 ? x.pathname : '')); } catch { /* keep as typed */ }
    if (label.length > 60) label = label.slice(0, 58) + '…';
    return `<a class="idealink" href="${url}" target="_blank" rel="noopener">${label}</a>${trail}`;
  });
}
const ideaList = (sent) => Object.entries(state.ideas).filter(([, x]) => !!x.sent === sent)
  .sort((a, b) => (sent ? (b[1].sent - a[1].sent) || (a[1].at - b[1].at) : (a[1].at || 0) - (b[1].at || 0)));
// An idea is { name, url } (older ones are a single { text }): the recipe name big, the link under it,
// then who added it, then the date small and grey.
const ideaName = (x) => x.name || x.text || '';
const fixUrl = (u) => { u = String(u || '').trim(); return !u ? '' : /^https?:\/\//i.test(u) ? u : 'https://' + u; };
function ideaRowHtml([id, x]) {
  const mine = (x.email || '').toLowerCase() === myEmail();
  return `<li class="idea"><div>
    <span class="ideaname">${x.name ? esc(x.name) : linkify(x.text || '')}</span>
    ${x.url ? `<span class="idealinkline">${linkify(x.url)}</span>` : ''}
    <span class="ideaby">${esc(nameFor(x.email))}${x.at ? `<span class="ideadate"> – ${esc(ideaDay(x.at))}${x.sent ? ` · sent ${esc(ideaDay(x.sent))}` : ''}</span>` : ''}</span></div>
    ${mine ? `<button class="delnote" data-editidea="${esc(id)}" aria-label="Edit idea">✎</button>` : `<button class="delnote" data-delidea="${esc(id)}" aria-label="Delete idea">×</button>`}</li>`;
}
const ideaDay = (at) => new Date(at).toLocaleDateString(undefined, { month: 'short', day: 'numeric' });
// An idea counts as added once a recipe with that name (or containing it) is in the recipes.
const normName = (x) => String(x || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
function ideaRecipe(x) {
  const n = normName(x.name || ''); if (n.length < 4) return null;
  for (const meal of allMeals()) for (const d of meal.dishes) {
    const dn = normName(d.name);
    if (dn === n || (dn.includes(n)) || (dn.length >= 8 && n.includes(dn))) return meal;
  }
  return null;
}
function renderPlanning() {
  setTitle('Planning', false); setTab('planning');
  // keep a half-typed idea when the list redraws
  const added = Object.entries(state.ideas).filter(([, x]) => ideaRecipe(x)).sort((a, b) => (b[1].at || 0) - (a[1].at || 0));
  const isAdded = ([, x]) => !!ideaRecipe(x);
  const open = ideaList(false).filter((e) => !isAdded(e)), sent = ideaList(true).filter((e) => !isAdded(e));
  let h = banners() + planningSwitch('ideas');
  h += '<button class="btn addidea" id="addidea">＋ Add idea</button>';
  h += `<section class="card"><p class="section-label">Ideas${open.length ? ` (${open.length})` : ''}</p>`;
  h += open.length ? `<ul class="list idealist">${open.map(ideaRowHtml).join('')}</ul>`
    : '<p class="small">No ideas waiting. Tap Add idea; it shows up on both phones.</p>';
  h += '</section>';
  if (open.length) h += `<section class="card sendcard"><p class="section-label">Send to the recipe Claude</p>
    <p class="small">Copies the ${open.length === 1 ? 'idea' : `${open.length} ideas`} above as a message to paste to the recipe Claude, then marks ${open.length === 1 ? 'it' : 'them'} as sent.</p>
    <button class="btn" type="button" id="copyideas">Copy for recipe Claude</button></section>`;
  if (sent.length) h += `<details class="card ideasent"${ideasSentOpen ? ' open' : ''}><summary class="section-label">Sent to Claude (${sent.length})</summary>
    <ul class="list idealist">${sent.map(ideaRowHtml).join('')}</ul>
    <div class="sentbtns"><button class="btn ghost" type="button" id="unsendall">Move all back to Ideas</button><button class="btn outline-danger" type="button" id="clearsent">Clear sent ideas</button></div></details>`;
  if (added.length) h += `<details class="card ideasent"${ideasSentOpen ? ' open' : ''}><summary class="section-label">In the Recipe Book ✓ (${added.length})</summary>
    <ul class="list idealist">${added.map(([id, x]) => { const meal = ideaRecipe(x); return ideaRowHtml([id, x]).replace('<span class="ideaname">', `<a class="ideaadded" href="#/meal/${esc(meal.id)}">Open recipe ›</a><span class="ideaname">`); }).join('')}</ul></details>`;
  view().innerHTML = h;
}
let ideasSentOpen = false;
function ideasMessage(list) {
  return `Recipe ideas from the app's Planning tab. Please add each to the Archive tab of recipes.xlsx (not Recipes), so it shows in the app's Recipe Book under "Want to try", following the "Recipe ideas from the app" section of CLAUDE.md: the Main as its own entry with every column filled (Shop, Timer, Prep task), plus any side as a separate entry, using the idea's name in the dish name. Then run the checker and push. For a link, open the page and use that recipe. If an idea is too vague, suggest a recipe and check with me first.\n\n`
    + list.map(([, x], n) => `${n + 1}. ${ideaName(x).replace(/\n+/g, ' / ')}${x.url ? ` - ${x.url}` : ''} (from ${nameFor(x.email)})`).join('\n');
}
async function copyIdeas() {
  const list = ideaList(false).filter(([, x]) => !ideaRecipe(x)); if (!list.length) return;
  const msg = ideasMessage(list);
  let copied = false;
  try { await navigator.clipboard.writeText(msg); copied = true; } catch { /* shown below to copy by hand */ }
  const r = await formDialog(`<p class="picktitle">${copied ? 'Copied' : 'Copy this message'}</p>
    <p class="small">${copied ? 'Paste it to the recipe Claude.' : 'Select all of it, copy, and paste it to the recipe Claude.'} Mark ${list.length === 1 ? 'this idea' : `these ${list.length} ideas`} as sent?</p>
    <textarea class="dlg-input" rows="6" readonly>${esc(msg)}</textarea>
    <div class="row-btns"><button class="btn ghost" value="">Not yet</button><button class="btn" value="save">Mark as sent</button></div>`);
  if (r.value !== 'save') return;
  const now = Date.now(), ch = {};
  for (const [id, x] of list) ch[id] = { ...x, sent: now };
  writeIdeas(ch); toast('Marked as sent');
}
async function addIdea(name = '', url = '') { // the Add idea dialog (stays near the top: it opens the keyboard)
  const r = await formDialog(`<p class="picktitle">New recipe idea</p>
    <label class="field">Recipe name<input class="dlg-input" name="name" maxlength="120" placeholder="e.g. Salmon teriyaki rice bowls" value="${esc(name)}" autocomplete="off"></label>
    <label class="field">Link <span class="small">(optional)</span><input class="dlg-input" name="url" type="text" inputmode="url" autocapitalize="off" spellcheck="false" maxlength="${IDEA_MAX}" placeholder="Paste a recipe link" value="${esc(url)}" autocomplete="off"></label>
    <div class="row-btns"><button class="btn ghost" value="">Cancel</button><button class="btn" value="save">Add idea</button></div>`);
  if (r.value !== 'save') return;
  const n = clean(r.form.name.value, 120), u = fixUrl(r.form.url.value);
  if (!n) { toast('Type the recipe name first'); addIdea('', r.form.url.value); return; }
  writeIdeas({ [Date.now().toString(36) + Math.random().toString(36).slice(2, 6)]: { name: n, url: u, email: myEmail(), at: Date.now() } });
  toast('Idea added');
}
async function editIdea(id) {
  const x = state.ideas[id]; if (!x) return;
  const r = await formDialog(`<p class="picktitle">Edit idea</p>
    <label class="field">Recipe name<input class="dlg-input" name="name" maxlength="120" value="${esc(ideaName(x))}" autocomplete="off"></label>
    <label class="field">Link <span class="small">(optional)</span><input class="dlg-input" name="url" type="text" inputmode="url" autocapitalize="off" spellcheck="false" maxlength="${IDEA_MAX}" value="${esc(x.url || '')}" autocomplete="off"></label>
    <div class="row-btns"><button class="btn outline-danger" value="delete" style="width:auto;margin-right:auto">Delete</button>${x.sent ? '<button class="btn ghost" value="unsend">Not sent</button>' : ''}<button class="btn ghost" value="">Cancel</button><button class="btn" value="save">Save</button></div>`);
  if (r.value === 'delete') { writeIdeas({ [id]: null }); undoToast('Idea deleted', () => writeIdeas({ [id]: x })); return; }
  if (r.value === 'unsend') { writeIdeas({ [id]: { ...x, sent: null } }); toast('Moved back to Ideas'); return; } // null, not left out: a saved merge would keep the old date
  if (r.value !== 'save') return;
  const name = clean(r.form.name.value, 120), url = fixUrl(r.form.url.value);
  if (!name) { toast('Type the recipe name, or tap Delete'); return; }
  if (name !== ideaName(x) || url !== (x.url || '')) { const { text, ...rest } = x; writeIdeas({ [id]: { ...rest, name, url } }); toast('Idea updated'); }
}

// ---------------------------------------------------------------- dish photos
// One photo per dish in Firestore photos/<dish ID> as { img: JPEG data URL, email, at }. The phone shrinks it first
// (longest side 1100 px, about 100-250 KB) so it fits in a document and the free plan. A dish's photo is only
// loaded when its recipe is opened (1 read), then kept on the phone.
const photoUnsubs = {};
function watchPhoto(key) {
  if (!state.firebase || photoUnsubs[key]) return;
  photoUnsubs[key] = onSnapshot(doc(state.firebase.db, 'photos', key), safely('photos', (snap) => {
    state.photos[key] = (snap.exists() && snap.data().img) || null; paintPhotos();
    if (state.photos[key] && thumbsLoaded && !state.thumbs[key] && !snap.metadata.fromCache) makeThumb(state.photos[key]).then((t) => writeThumb(key, t)).catch(() => {});
  }), (err) => { console.warn('photo', key, err); });
}
function paintPhotos() { // fill each dish's photo slot and its Add photo button
  document.querySelectorAll('[data-photo]').forEach((slot) => {
    const key = slot.dataset.photo; watchPhoto(key);
    const img = state.photos[key];
    if (img) { if (slot.dataset.src !== img) { slot.innerHTML = `<img src="${img}" alt="Photo of ${esc(slot.dataset.name)}">`; slot.dataset.src = img; } }
    else { slot.innerHTML = ''; slot.dataset.src = ''; }
    slot.hidden = !img;
    const unknown = state.photos[key] === undefined && !!state.firebase; // still loading
    document.querySelectorAll(`[data-addphoto="${CSS.escape(key)}"]`).forEach((b) => { b.hidden = !!img || unknown; });
    document.querySelectorAll(`[data-hasphoto="${CSS.escape(key)}"]`).forEach((b) => { b.hidden = !img; }); // Replace / Remove in the recipe menu
  });
}
const bookThumbs = () => Object.keys(state.thumbs).length > 0; // show the thumbnail column once any dish has a photo
function makeThumb(src) { // a 100 px square, centre-cropped, about 3-5 KB
  return new Promise((res, rej) => { const im = new Image(); im.onload = () => {
    const c = document.createElement('canvas'); c.width = c.height = 100; const m = Math.min(im.width, im.height);
    c.getContext('2d').drawImage(im, (im.width - m) / 2, (im.height - m) / 2, m, m, 0, 0, 100, 100); res(c.toDataURL('image/jpeg', 0.6)); };
  im.onerror = rej; im.src = src; });
}
function writeThumb(key, t) {
  if (t) state.thumbs[key] = t; else delete state.thumbs[key];
  if (!state.firebase) { try { ls.set('local-thumbs', state.thumbs); } catch { /* full */ } return; }
  setDoc(doc(state.firebase.db, 'photos', '_thumbs'), { t: { [key]: t || deleteField() } }, { merge: true }).catch(() => {});
}
async function shrinkPhoto(file) {
  const bmp = await createImageBitmap(file); // turned the right way up from the camera's orientation
  const k = Math.min(1, 1100 / Math.max(bmp.width, bmp.height));
  const c = document.createElement('canvas'); c.width = Math.round(bmp.width * k); c.height = Math.round(bmp.height * k);
  c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
  let url = '';
  for (const q of [0.8, 0.7, 0.6, 0.5, 0.4]) { url = c.toDataURL('image/jpeg', q); if (url.length < 340000) break; }
  return url;
}
function pickPhoto(key) {
  const inp = document.createElement('input'); inp.type = 'file'; inp.accept = 'image/*';
  inp.addEventListener('change', async () => {
    const f = inp.files && inp.files[0]; if (!f) return;
    toast('Saving photo…', 10000);
    try { savePhoto(key, await shrinkPhoto(f)); } catch (e) { toast('Couldn’t use that photo: ' + e.message, 4000); }
  });
  inp.click();
}
function removePhoto(key) {
  const old = state.photos[key]; if (!old) return;
  savePhoto(key, null); undoToast('Photo removed', () => savePhoto(key, old));
}
function savePhoto(key, img) {
  state.photos[key] = img; paintPhotos();
  if (img) makeThumb(img).then((t) => writeThumb(key, t)).catch(() => {}); else writeThumb(key, null);
  if (!state.firebase) { try { ls.set('local-photos', state.photos); } catch { /* full */ } if (img) toast('Photo saved'); return; }
  const ref = doc(state.firebase.db, 'photos', key);
  (img ? setDoc(ref, { img, email: myEmail(), at: Date.now() }) : deleteDoc(ref))
    .then(() => { if (img) toast('Photo saved'); })
    .catch((e) => toast('Could not save the photo: ' + (e.code || e.message), 5000));
}
async function photoMenu(key, name) {
  const dlg = formDialog(`<p class="picktitle">${esc(name)}</p>
    <div class="picklist"><button class="pickopt" value="replace"><strong>Replace photo</strong></button><button class="pickopt" value="remove"><strong>Remove photo</strong></button></div>
    <div class="row-btns"><button class="btn ghost" value="">Cancel</button></div>`);
  $('#formdlg').classList.add('centered');
  const r = await dlg;
  if (r.value === 'replace') pickPhoto(key);
  else if (r.value === 'remove') removePhoto(key);
}

// ---------------------------------------------------------------- shopping notes
// Shared notes at the top of the shopping list, stored in Firestore notes/shopping as { n: { id: { text, email, at } } }.
// Notes and added items store who wrote them by email; the name shown is looked up live from profiles,
// so setting or changing a name in Settings relabels everything that person wrote.
const myEmail = () => (state.user && state.user.email ? state.user.email.toLowerCase() : 'local');
function nameFor(email, fallback) {
  if (!email) return fallback || '';
  const e = String(email).toLowerCase();
  if (e === 'local') return state.myName || 'Me';
  return state.names[e] || (e === myEmail() && state.myName) || email;
}
const fmtWhen = (at) => (at ? new Date(at).toLocaleDateString(undefined, { month: 'short', day: 'numeric' }) : ''); // just the date, e.g. "Nov 5"
function writeNotes(changes) {
  for (const [id, v] of Object.entries(changes)) { if (v) state.notes[id] = v; else delete state.notes[id]; }
  renderNotes();
  if (!state.firebase) { ls.set('local-notes', state.notes); return; }
  const n = {}; for (const [id, v] of Object.entries(changes)) n[id] = v || deleteField();
  setDoc(doc(state.firebase.db, 'notes', 'shopping'), { n }, { merge: true }).catch((e) => toast('Could not save note: ' + (e.code || e.message)));
}
const NOTE_LIMIT = 3; // notes per person
const myNoteCount = () => Object.values(state.notes).filter((x) => (x.email || '').toLowerCase() === myEmail()).length;
function addNote(text) {
  if (myNoteCount() >= NOTE_LIMIT) { toast(`You already have ${NOTE_LIMIT} notes. Delete one to add another.`); return false; }
  const id = Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  writeNotes({ [id]: { text, email: myEmail(), at: Date.now() } });
  return true;
}
function notesListHtml() {
  const notes = Object.entries(state.notes).sort((a, b) => (b[1].at || 0) - (a[1].at || 0));
  if (!notes.length) return '';
  // notes tick like shopping items (stored with the shopping ticks, not counted in the shopping progress)
  return notes.map(([id, x]) => itemHtml(SHOP(), 'n-' + id,
    `<span class="body"><span class="notetext">${esc(x.text)}</span>
    <span class="sub">${esc(nameFor(x.email, x.by))}${x.at ? `<span class="notedate"> – ${esc(fmtWhen(x.at))}</span>` : ''}</span></span>${(x.email || '').toLowerCase() === myEmail()
      ? `<button class="delnote" data-editnote="${esc(id)}" aria-label="Edit note">✎</button>`
      : `<button class="delnote" data-delnote="${esc(id)}" aria-label="Delete note">×</button>`}`, 'extra note')).join('');
}
// Folded, the notes card is one line: the newest note, and how many more. Tap to see and tick them all.
let notesOpen = false;
function notesSummaryHtml() {
  const notes = Object.values(state.notes).sort((a, b) => (b.at || 0) - (a.at || 0)); if (!notes.length) return '';
  return `${ICONS.note}<span class="ns-closed"><span class="ns-first">${esc(notes[0].text)}</span>${notes.length > 1 ? `<span class="ns-more">+${notes.length - 1} more</span>` : ''}</span><span class="ns-open">Notes (${notes.length})</span>`;
}
document.addEventListener('toggle', (e) => { if (e.target.id === 'notescard') notesOpen = e.target.open; }, true);
function renderNotes() {
  const ul = $('#notelist'); if (!ul) return;
  ul.innerHTML = notesListHtml(); $('#notesummary').innerHTML = notesSummaryHtml();
  $('#notescard').hidden = !Object.keys(state.notes).length;
  refreshTicks();
}

// ---------------------------------------------------------------- items added to the shopping list
// Stored in Firestore shopping/extra as { i: { id: { name, amount, email, at } } }. Ticked like any other item.
const extraKey = (id) => 'x-' + id;
const extraKeys = () => Object.keys(state.extras).map(extraKey);
function writeExtras(changes) {
  for (const [id, v] of Object.entries(changes)) { if (v) state.extras[id] = v; else delete state.extras[id]; }
  renderExtras();
  if (!state.firebase) { ls.set('local-extras', state.extras); return; }
  const i = {}; for (const [id, v] of Object.entries(changes)) i[id] = v || deleteField();
  setDoc(doc(state.firebase.db, 'shopping', 'extra'), { i }, { merge: true }).catch((e) => toast('Could not save item: ' + (e.code || e.message)));
}
// Categories for added items: the shopping list's own sections, plus "Other".
const shopSections = () => (state.model ? state.model.shopping.map((sec) => sec.name) : []);
const extraSection = (x) => (shopSections().includes(x.section) ? x.section : 'Other');
function extrasListHtml(section, staplesOnly) {
  const items = Object.entries(state.extras).filter(([, x]) => (staplesOnly ? isStaple(x.name) : extraSection(x) === section && !isStaple(x.name))).sort((a, b) => (a[1].at || 0) - (b[1].at || 0));
  if (!items.length) return '';
  return items.map(([id, x]) => itemHtml(SHOP(), extraKey(id),
    `<span class="body"><strong>${esc(x.name)}</strong>${x.amount ? ` · ${esc(x.amount)}` : ''}<span class="sub">Added by ${esc(nameFor(x.email))}</span></span>${
      (x.email || '').toLowerCase() === myEmail()
        ? `<button class="delnote" data-editextra="${esc(id)}" aria-label="Edit ${esc(x.name)}">✎</button>`
        : `<button class="delnote" data-delextra="${esc(id)}" aria-label="Remove ${esc(x.name)}">×</button>`}`, 'extra', '', x.name)).join('');
}
function renderExtras() { // added items sit inside the shopping sections, so redraw the shopping list
  if (/^#\/shopping/.test(location.hash)) route();
}

// ---------------------------------------------------------------- "+" in the top bar (shopping list)
function formDialog(html) { // shows a small form dialog; resolves with the submit button's value and the form
  let dlg = $('#formdlg');
  if (!dlg) { dlg = document.createElement('dialog'); dlg.id = 'formdlg'; document.body.appendChild(dlg); }
  dlg.className = ''; dlg.innerHTML = `<form method="dialog">${html}</form>`;
  dlg.returnValue = '';
  // Enter in a text box saves (otherwise the first button, Cancel or Remove, would be used)
  dlg.querySelectorAll('input').forEach((inp) => inp.addEventListener('keydown', (ev) => {
    if (ev.key === 'Enter') { ev.preventDefault(); const save = dlg.querySelector('button[value="save"]'); if (save) save.click(); }
  }));
  // Nothing to type in? Then no keyboard will open, so it sits in the middle of the screen with nothing focused.
  const typing = dlg.querySelector('input:not([type=checkbox]):not([readonly]), textarea:not([readonly]), select');
  if (!typing) dlg.classList.add('mid');
  return new Promise((resolve) => { dlg.addEventListener('close', () => resolve({ value: dlg.returnValue, form: dlg.querySelector('form') }), { once: true }); dlg.showModal(); if (!typing && document.activeElement) document.activeElement.blur(); });
}
const clean = (v, max) => String(v || '').replace(/\s+/g, ' ').trim().slice(0, max);
async function addMenu() {
  const full = myNoteCount() >= NOTE_LIMIT;
  const dlg = formDialog(`<p class="picktitle">Add to the shopping list</p>
    <div class="picklist">
      <button class="pickopt" value="item"><strong>Item</strong><span class="small">Something to buy, with an amount</span></button>
      <button class="pickopt" value="note"${full ? ' disabled' : ''}><strong>Note</strong><span class="small">${full ? `You have ${NOTE_LIMIT} notes. Delete one to add another.` : `A note for both of you (up to ${NOTE_LIMIT} each)`}</span></button>
    </div><div class="row-btns"><button class="btn ghost" value="">Cancel</button></div>`);
  $('#formdlg').classList.add('mid'); // no typing here, so it sits in the middle
  const r = await dlg;
  if (r.value === 'item') editItem(null);
  else if (r.value === 'note') {
    const nd = formDialog(`<p class="picktitle">New note</p>
      <input class="dlg-input" name="text" maxlength="200" placeholder="e.g. Coupon in the glovebox" autocomplete="off">
      <div class="row-btns"><button class="btn ghost" value="">Cancel</button><button class="btn" value="save">Add note</button></div>`);
    $('#formdlg').classList.add('mid'); // small enough to sit in the middle
    const n = await nd;
    const text = clean(n.form.text.value, 200);
    if (n.value === 'save' && text && addNote(text)) toast('Note added');
  }
}
async function editNote(id) { // your own notes: change the text or delete
  const x = state.notes[id]; if (!x) return;
  const r = await formDialog(`<p class="picktitle">Edit note</p>
    <input class="dlg-input" name="text" maxlength="200" value="${esc(x.text)}" autocomplete="off">
    <div class="row-btns"><button class="btn outline-danger" value="delete" style="width:auto;margin-right:auto">Delete</button><button class="btn ghost" value="">Cancel</button><button class="btn" value="save">Save</button></div>`);
  if (r.value === 'delete') { writeNotes({ [id]: null }); undoToast('Note deleted', () => writeNotes({ [id]: x })); return; }
  if (r.value !== 'save') return;
  const text = clean(r.form.text.value, 200);
  if (!text) { toast('Type the note, or tap Delete'); return; }
  if (text !== x.text) { writeNotes({ [id]: { ...x, text } }); toast('Note updated'); }
}
async function editItem(id) { // id null = new item; otherwise edit your own item
  const x = id ? state.extras[id] : null;
  const r = await formDialog(`<p class="picktitle">${x ? 'Edit item' : 'New item'}</p>
    <label class="field">Item<input class="dlg-input" name="name" maxlength="80" placeholder="e.g. Paper towels" value="${esc(x ? x.name : '')}" autocomplete="off"></label>
    <label class="field">Amount <span class="small">(optional)</span><input class="dlg-input" name="amount" maxlength="40" placeholder="e.g. 1 pack" value="${esc(x ? x.amount : '')}" autocomplete="off"></label>
    <label class="field">Category<select class="dlg-input" name="section">${[...shopSections(), 'Other'].map((sec) => `<option${(x ? extraSection(x) : 'Other') === sec ? ' selected' : ''}>${esc(sec)}</option>`).join('')}</select></label>
    <div class="row-btns">${x ? '<button class="btn outline-danger" value="delete" style="width:auto;margin-right:auto">Remove</button>' : ''}<button class="btn ghost" value="">Cancel</button><button class="btn" value="save">${x ? 'Save' : 'Add item'}</button></div>`);
  if (r.value === 'delete') { writeExtras({ [id]: null }); undoToast(`${x.name} removed`, () => writeExtras({ [id]: x })); return; }
  if (r.value !== 'save') return;
  const name = clean(r.form.name.value, 80), amount = clean(r.form.amount.value, 40), section = r.form.section.value;
  if (!name) { toast('Type the item name'); return; }
  if (x) writeExtras({ [id]: { ...x, name, amount, section } });
  else writeExtras({ [Date.now().toString(36) + Math.random().toString(36).slice(2, 6)]: { name, amount, section, email: myEmail(), at: Date.now() } });
  toast(x ? 'Item updated' : 'Item added');
}

// ---------------------------------------------------------------- ratings
// One record per meal in Firestore (ratings/<meal id>). It keeps the meal's names too,
// so the calendar can still show the rating after the recipe is removed from the sheet.
const myId = () => (state.user ? state.user.uid : 'local');
const votesNeeded = () => (state.firebase ? 2 : 1);
function ratingDoc(key) { // key: meal ID, meal name, or calendar value; ratings are stored under the main dish's ID
  const meal = mealByName(key);
  return (meal ? state.ratings[meal.rid] : state.ratings[key]) || null;
}
function ratingInfo(key) {
  const r = ratingDoc(key);
  const votes = Object.values((r && r.votes) || {}).filter((v) => typeof v === 'number');
  const mine = r && r.votes ? r.votes[myId()] || 0 : 0;
  const done = votes.length >= votesNeeded();
  const avg = votes.length ? votes.reduce((a, b) => a + b, 0) / votes.length : 0;
  return { record: r || null, count: votes.length, mine, done, avg };
}
const starsText = (avg) => `★ ${avg.toFixed(1)} / 5.0`;
function setRating(meal, stars) {
  const id = meal.rid;
  const r = state.ratings[id] || (state.ratings[id] = { votes: { ...((ratingDoc(meal) || {}).votes || {}) } });
  Object.assign(r, { name: meal.name, main: meal.main.name, sides: meal.sides.map((d) => d.name) });
  r.votes = { ...(r.votes || {}) };
  if (stars) r.votes[myId()] = stars; else delete r.votes[myId()];
  route();
  if (!state.firebase) { ls.set('local-ratings', state.ratings); return; }
  setDoc(doc(state.firebase.db, 'ratings', id), {
    name: meal.name, main: meal.main.name, sides: meal.sides.map((d) => d.name),
    votes: { [myId()]: stars || deleteField() }, at: serverTimestamp(),
  }, { merge: true }).catch((e) => toast('Could not save rating: ' + (e.code || e.message)));
}
// ---------------------------------------------------------------- notes for next time
// One shared note per recipe (keyed on the main dish's ID). Written under the rating once you've rated;
// from the next day on it also shows at the top of the recipe, so you read it before cooking.
function writeNextTime(rid, v) {
  if (v) state.nextTime[rid] = v; else delete state.nextTime[rid];
  route();
  if (!state.firebase) { ls.set('local-nexttime', state.nextTime); return; }
  setDoc(doc(state.firebase.db, 'notes', 'nexttime'), { n: { [rid]: v || deleteField() } }, { merge: true }).catch((e) => toast('Could not save: ' + (e.code || e.message)));
}
const nextTimeBy = (x) => `${esc(nameFor(x.email))}<span class="notedate"> – ${esc(new Date(x.at).toLocaleDateString(undefined, { month: 'short', day: 'numeric' }))}</span>`;
function nextTimeTopHtml(meal) { // shown at the top once the note is from an earlier day
  const x = state.nextTime[meal.rid]; if (!x || !x.text || dayKey(new Date(x.at)) >= dayKey(new Date())) return '';
  return `<section class="card nexttop">${ICONS.note}<div><p class="section-label">From last time</p><p class="nt-text">${esc(x.text)}</p><p class="small">${nextTimeBy(x)}</p></div></section>`;
}
function nextTimeCardHtml(meal) { // under the rating: add or edit the note (after you've rated, or if there already is one)
  const x = state.nextTime[meal.rid];
  if (!x && !ratingInfo(meal.name).mine) return '';
  return `<section class="card nextcard"><p class="section-label">Note for next time</p>${x
    ? `<div class="nt-row"><div><p class="nt-text">${esc(x.text)}</p><p class="small">${nextTimeBy(x)}</p></div><button class="delnote" data-editnext="${esc(meal.id)}" aria-label="Edit note">✎</button></div>`
    : `<button class="btn ghost" data-editnext="${esc(meal.id)}">＋ Add a note for next time</button>`}</section>`;
}
async function editNextTime(meal) {
  const x = state.nextTime[meal.rid];
  const r = await formDialog(`<p class="picktitle">Note for next time</p><p class="small" style="margin:0 0 8px">${esc(meal.main.name)}</p>
    <textarea class="dlg-input" name="text" rows="4" maxlength="400">${esc(x ? x.text : '')}</textarea>
    <div class="row-btns">${x ? '<button class="btn outline-danger" value="delete" style="width:auto;margin-right:auto">Delete</button>' : ''}<button class="btn ghost" value="">Cancel</button><button class="btn" value="save">Save</button></div>`);
  if (r.value === 'delete') { writeNextTime(meal.rid, null); undoToast('Note deleted', () => writeNextTime(meal.rid, x)); return; }
  if (r.value !== 'save') return;
  const text = String(r.form.text.value || '').replace(/[ \t]+/g, ' ').trim().slice(0, 400);
  if (!text) { if (x) { writeNextTime(meal.rid, null); undoToast('Note deleted', () => writeNextTime(meal.rid, x)); } return; }
  if (!x || text !== x.text) { writeNextTime(meal.rid, { text, email: myEmail(), at: Date.now() }); toast('Saved for next time'); }
}
function ratingCardHtml(meal) {
  const info = ratingInfo(meal.name);
  let status;
  if (info.done) status = `<strong class="avg">${starsText(info.avg)}</strong> average of ${info.count} votes`;
  else if (info.mine) status = 'Your vote is in. Waiting for the other vote.';
  else if (info.count) status = 'The other vote is in. Add yours to see the average.';
  else status = 'Tap the stars to rate this meal. The average shows once you have both voted.';
  return `<section class="card rate" id="${meal.id}-rate"><p class="section-label">Rate this meal</p>
    <div class="stars-input" role="radiogroup" aria-label="Your rating">${[1, 2, 3, 4, 5].map((n) =>
      `<button class="star${n <= info.mine ? ' on' : ''}" data-rate="${n}" data-meal="${meal.id}" aria-label="${n} star${n > 1 ? 's' : ''}">★</button>`).join('')}</div>
    <p class="small">${info.mine ? `You gave ${info.mine} star${info.mine > 1 ? 's' : ''}. Tap it again to remove your vote. ` : ''}${status}</p></section>`;
}

// ---------------------------------------------------------------- calendar store
const dayKey = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const keyToDate = (k) => { const [y, m, d] = k.split('-').map(Number); return new Date(y, m - 1, d, 12); };
const fmtDay = (d, opts = { weekday: 'short', month: 'short', day: 'numeric' }) => d.toLocaleDateString(undefined, opts);
function weekStart(offset) { const d = new Date(); d.setHours(12, 0, 0, 0); d.setDate(d.getDate() - d.getDay() + offset * 7); return d; }
const allMeals = () => (state.model ? [...state.model.meals, ...(state.model.past || [])] : []);
const mealByName = (v) => (v && typeof v === 'string' ? allMeals().find((x) => x.rid === v || x.name === v) || null : null); // by ID (or name)
const planKey = (v) => { const meal = mealByName(v); return meal ? meal.rid : v; }; // calendar value -> meal ID
const mealById = (id) => allMeals().find((x) => x.id === id) || null;
// A day holds dinner (state.plan, as always) plus optional breakfast, lunch and dessert (state.planX).
const SLOTS = [['breakfast', 'Breakfast'], ['lunch', 'Lunch'], ['dinner', 'Dinner'], ['dessert', 'Dessert']];
const slotName = (slot) => (SLOTS.find((x) => x[0] === slot) || [, 'Dinner'])[1];
const slotValue = (day, slot) => (slot === 'dinner' ? state.plan[day] : (state.planX[day] || {})[slot]);
function planList() { // every planned thing: [day, value, slot]
  const out = Object.entries(state.plan).map(([k, v]) => [k, v, 'dinner']);
  for (const [k, slots] of Object.entries(state.planX)) for (const [slot, v] of Object.entries(slots || {})) if (v) out.push([k, v, slot]);
  return out;
}
function setPlan(day, mealName, slot = 'dinner') { // mealName: recipe meal name or ID, { text } for a custom entry, or null to remove
  if (slot === 'dinner') { writePlan({ [day]: mealName || null }); return; }
  const cur = { ...(state.planX[day] || {}) };
  if (mealName) cur[slot] = mealName; else delete cur[slot];
  if (Object.keys(cur).length) state.planX[day] = cur; else delete state.planX[day];
  route();
  if (!state.firebase) { ls.set('local-planx', state.planX); return; }
  setDoc(doc(state.firebase.db, 'calendar', 'plan'), { x: { [day]: { [slot]: mealName || deleteField() } }, at: serverTimestamp() }, { merge: true })
    .catch((e) => toast('Could not save: ' + (e.code || e.message)));
}
function writePlan(changes) { // { 'YYYY-MM-DD': value or null }, saved together
  const d = {};
  for (const [day, v] of Object.entries(changes)) {
    if (v) state.plan[day] = v; else delete state.plan[day];
    d[day] = v || deleteField();
  }
  route();
  if (!state.firebase) { ls.set('local-plan', state.plan); return; }
  setDoc(doc(state.firebase.db, 'calendar', 'plan'), { d, at: serverTimestamp() }, { merge: true })
    .catch((e) => toast('Could not save: ' + (e.code || e.message)));
}
function movePlan(from, to) { // drag and drop: move, or swap if the target day already has something
  if (from === to || !state.plan[from]) return;
  writePlan({ [to]: state.plan[from], [from]: state.plan[to] || null });
}
function nextPlanned(mealName) { // soonest date (today or later) this meal is planned for
  const today = dayKey(new Date());
  const key = planKey(mealName);
  return planList().filter(([k, v]) => k >= today && planKey(v) === key).map(([k]) => k).sort()[0] || null;
}

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
  el.title = labels[state.sync] || ''; el.setAttribute('aria-label', el.title); // when synced, only the green dot shows
}
addEventListener('online', () => updateSync(false));
addEventListener('offline', () => updateSync(true));

// ---------------------------------------------------------------- recipe loading
async function fetchOk(url, type) {
  const r = await fetch(url, { cache: 'no-store' });
  if (!r.ok) throw new Error(`HTTP ${r.status}`);
  return type === 'buf' ? r.arrayBuffer() : r.text();
}
// Recipes come only from recipes.xlsx in the GitHub repository (see CLAUDE.md).
async function fetchRows() {
  return rowsFromWorkbook(XLSX.read(await fetchOk('recipes.xlsx', 'buf'), { type: 'array' }));
}
function setRows(rows, info) {
  const model = buildModel(rows);
  if (!model.meals.length && !model.past.length) throw new Error('No recipes found in recipes.xlsx. Is there a tab named "Recipes" with the usual columns?');
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
    try {
      setRows(await fetchRows(), { at: Date.now(), error: '' });
      state.recipesFresh = true; tidyOldLists();
    } catch (e) {
      // keep showing the last recipes that loaded (they're saved on the phone)
      console.warn('Recipe load failed', e);
      const info = { ...(state.recipeInfo || {}), error: 'Could not load the latest recipes: ' + e.message };
      state.recipeInfo = info; ls.set('recipes-info', info);
      if (!state.model || location.hash.startsWith('#/settings')) route();
    } finally { loading = null; }
  })();
  return loading;
}

// ---------------------------------------------------------------- weekly shopping list turnover
// When a new list (newer List date) arrives: items added in the app that were ticked on the old list are
// removed (bought), unticked ones carry over, and the old list's ticks are deleted. Only lists OLDER than the
// one this phone has loaded are touched, so a phone that hasn't loaded the new list can't wipe its ticks.
// Runs only with freshly loaded recipes and ticks.
const tidied = new Set();
function tidyOldLists() {
  const cur = autoShop() ? 'auto' : state.model && state.model.shopList; // switching to the automatic list retires the dated lists
  if (!cur || !state.recipesFresh || (state.firebase && !(ticksLoaded && extrasLoaded))) return;
  for (const id of Object.keys(state.ticks)) {
    const m = id === 'shopping' ? [id, ''] : id.match(/^shopping-(.+)$/); // 'shopping' = record from before lists had dates
    if (!m || !(m[1] < cur) || tidied.has(id)) continue;
    tidied.add(id);
    const ticked = state.ticks[id] || {};
    const bought = Object.keys(state.extras).filter((x) => ticked[extraKey(x)]);
    if (bought.length) writeExtras(Object.fromEntries(bought.map((x) => [x, null])));
    deleteTicks(id);
  }
}
function deleteTicks(id) {
  delete state.ticks[id];
  if (!state.firebase) { ls.set('local-ticks', state.ticks); return; }
  deleteDoc(doc(state.firebase.db, 'ticks', id)).catch((e) => console.warn('could not delete ticks', id, e));
}

// ---------------------------------------------------------------- cooking together: dinner time, timers, who makes what
// One shared record, cooking/session: { timers: { id: {...} } }
function writeSession(part, key, value) { // part: 'timers'; value null = delete
  const cur = state.session[part] || (state.session[part] = {});
  if (value == null) delete cur[key]; else cur[key] = value;
  if (!state.firebase) { ls.set('local-session', state.session); renderTimers(); route(); return; }
  renderTimers();
  setDoc(doc(state.firebase.db, 'cooking', 'session'), { [part]: { [key]: value == null ? deleteField() : value } }, { merge: true })
    .catch((e) => toast('Could not save: ' + (e.code || e.message)));
}
const dishKey = (d) => d.id || 'd-' + slug(d.name);

// Step timers: "Roast 20-25 minutes" gets a 20-minute timer button. Timers are shared, so both phones see them.
function stepMinutes(text) {
  const t = String(text || '');
  let m = /(\d+)(?:\s*(?:-|–|to)\s*\d+)?\s*(?:minutes?|mins?)\b/i.exec(t);
  if (m) return +m[1];
  m = /(\d+)(?:\s*(?:-|–|to)\s*\d+)?\s*(?:hours?|hrs?)\b/i.exec(t);
  return m ? +m[1] * 60 : 0;
}
const timerLabel = (d, st) => st.timer || `${shortName(d)}: step ${st.num}`;
function runningTimerFor(step) { return Object.entries(state.session.timers || {}).find(([, t]) => t.step === step) || null; }
async function timerTap(label, mins, mealId, step) {
  const run = runningTimerFor(step);
  if (run) { if (await confirmDialog(`Cancel the "${run[1].label}" timer?`, 'Cancel timer')) writeSession('timers', run[0], null); return; }
  startTimer(label, mins, mealId, step);
}
function startTimer(label, mins, mealId, step) {
  const id = Date.now().toString(36) + Math.random().toString(36).slice(2, 5);
  writeSession('timers', id, { label, mins, ends: Date.now() + mins * 60000, meal: mealId, step: step || '', by: myEmail() });
  audioUnlock(); updateWakeLock();
  toast(`Timer started: ${mins} min`);
}
const alarmed = new Set();
let timerTick = null;
function renderTimers() {
  const bar = $('#timerbar'); if (!bar) return;
  const list = Object.entries(state.session.timers || {}).sort((a, b) => a[1].ends - b[1].ends);
  bar.hidden = !list.length;
  document.body.classList.toggle('has-timers', !!list.length);
  const cookBar = $('#cooktimers');
  if (cookBar) cookBar.hidden = !list.length;
  if (!list.length) { clearInterval(timerTick); timerTick = null; bar.innerHTML = ''; if (cookBar) cookBar.innerHTML = ''; markTimerButtons(); return; }
  const now = Date.now();
  bar.innerHTML = list.map(([id, t]) => {
    const left = Math.round((t.ends - now) / 1000), done = left <= 0;
    if (done && !alarmed.has(id)) { alarmed.add(id); alarm(t.label, (t.by || '').toLowerCase() === myEmail()); }
    const mm = Math.floor(Math.abs(left) / 60), ss = String(Math.abs(left) % 60).padStart(2, '0');
    const theirs = t.by && t.by.toLowerCase() !== myEmail();
    return `<div class="timer${done ? ' done' : ''}"><a href="#/meal/${esc(t.meal || '')}" class="tlabel">${esc(t.label)}${theirs ? ` <span class="towner">· ${esc(nameFor(t.by))}</span>` : ''}</a><span class="tleft">${done ? 'Done!' : `${mm}:${ss}`}</span><button class="tclose" data-stoptimer="${esc(id)}" aria-label="${done ? 'Dismiss' : 'Cancel'} timer">×</button></div>`;
  }).join('');
  if (cookBar) cookBar.innerHTML = bar.innerHTML; // cook mode covers the screen, so it shows the timers too
  if (!timerTick) timerTick = setInterval(renderTimers, 1000);
  markTimerButtons();
}
// timer buttons on steps: filled in and counting down while their timer runs
function markTimerButtons() {
  const now = Date.now();
  document.querySelectorAll('.cooktimer[data-cookstep]').forEach((btn) => { // the timer button in cook mode
    const run = runningTimerFor(btn.dataset.cookstep);
    const left = run ? Math.round((run[1].ends - now) / 1000) : 0;
    btn.classList.toggle('running', !!run && left > 0);
    btn.classList.toggle('rang', !!run && left <= 0);
    btn.innerHTML = !run ? `${ICONS.timer} Start ${btn.dataset.mins} min timer` : left > 0 ? `${ICONS.timer} ${Math.floor(left / 60)}:${String(left % 60).padStart(2, '0')} · tap to cancel` : `${ICONS.bell} Done · tap to clear`;
  });
  document.querySelectorAll('[data-timer][data-step]').forEach((btn) => {
    const run = runningTimerFor(btn.dataset.step);
    btn.classList.toggle('running', !!run && run[1].ends > now);
    btn.classList.toggle('rang', !!run && run[1].ends <= now);
    const side = btn.classList.contains('tb-side'); // the round button on the right of a step: short labels
    if (!run) { btn.innerHTML = side ? `${ICONS.timer}<span>${btn.dataset.timer}m</span>` : `${ICONS.timer} ${btn.dataset.timer} min`; return; }
    const left = Math.round((run[1].ends - now) / 1000), clock = `${Math.floor(left / 60)}:${String(left % 60).padStart(2, '0')}`;
    btn.innerHTML = side ? (left > 0 ? `${ICONS.timer}<span>${clock}</span>` : `${ICONS.bell}<span>Done</span>`)
      : left > 0 ? `${ICONS.timer} ${clock} · tap to cancel` : `${ICONS.bell} Done · tap to clear`;
  });
}
let audioCtx = null;
// Sound and vibration only on the phone that started the timer; the other phone just shows it.
function audioUnlock() { try { audioCtx = audioCtx || new (window.AudioContext || window.webkitAudioContext)(); if (audioCtx.state === 'suspended') audioCtx.resume(); } catch { /* no sound */ } }
function alarm(label, mine) {
  toast(`Time's up: ${label}`, 8000);
  if (!mine) return;
  if (navigator.vibrate) navigator.vibrate([400, 150, 400, 150, 800]);
  try {
    audioCtx = audioCtx || new (window.AudioContext || window.webkitAudioContext)();
    for (let i = 0; i < 6; i++) {
      const o = audioCtx.createOscillator(), g = audioCtx.createGain(), t0 = audioCtx.currentTime + i * 0.5;
      o.frequency.value = 880; g.gain.setValueAtTime(0.0001, t0); g.gain.exponentialRampToValueAtTime(0.4, t0 + 0.02); g.gain.exponentialRampToValueAtTime(0.0001, t0 + 0.35);
      o.connect(g); g.connect(audioCtx.destination); o.start(t0); o.stop(t0 + 0.4);
    }
  } catch { /* sound not available */ }
}

// Cook mode: one step at a time, big text
let cook = null; // { meal, dish }
// Cook mode's "All" tab (dish -1): the combined order as one list, each step still the dish's own (same ticks).
function combinedDish(meal) {
  return { name: 'All dishes', type: 'Combined', color: 0, anchor: `${meal.id}-combined`, combined: true,
    steps: meal.combined.map((x, i) => ({ ...x.s, num: i + 1, dishNum: x.s.num, dish: x.d })),
    groups: meal.dishes.flatMap((x) => x.groups.map((g) => ({ name: `${cookTabName(meal, x)}${g.name ? ' · ' + g.name : ''}`, items: g.items }))) };
}
const cookDishOf = (meal) => (cook.dish === -1 && meal.combined ? combinedDish(meal) : meal.dishes[Math.max(0, cook.dish)]);
const stepDish = (d, st) => (st.dish ? [st.dish, { ...st, num: st.dishNum }] : [d, st]); // the real dish and step behind a combined step
function openCook(mealId, dish = 0) {
  const meal = mealById(mealId); if (!meal) return;
  if (dish === -1 && !meal.combined) dish = 0;
  cook = { meal: meal.id, dish: dish === -1 ? -1 : Math.min(dish, meal.dishes.length - 1), step: null };
  renderCook();
}
function renderCook() {
  let el = $('#cookmode');
  if (!cook) { if (el) el.remove(); document.body.classList.remove('cooking'); return; }
  const meal = mealById(cook.meal); if (!meal) { cook = null; renderCook(); return; }
  if (!el) { el = document.createElement('div'); el.id = 'cookmode'; document.body.appendChild(el); }
  document.body.classList.add('cooking');
  const d = cookDishOf(meal), ticks = state.ticks[meal.id] || {}, n = d.steps.length;
  if (cook.step == null) { const i = d.steps.findIndex((x) => !ticks[x.key]); cook.step = i < 0 ? n : i; }
  const others = meal.dishes.length > 1 ? `<div class="cooknext"><p class="small">Cook another part of the meal:</p>${meal.dishes.map((x, i) => { if (i === cook.dish || !x.steps.length) return '';
    const left = x.steps.filter((st2) => !ticks[st2.key]).length;
    return `<button class="btn ghost chip-${x.color}" data-cook="dish:${i}">${esc(x.name)}<span class="small">${left ? `${left} step${left > 1 ? 's' : ''} to go` : 'Done ✓'}</span></button>`; }).join('')}</div>` : '';
  // A wheel of steps: the current one in the middle at full size; the ones around it smaller and fainter
  // (done ones in green). Scroll it, tap a nearby step, or use Back / Next.
  el.innerHTML = `<div class="cookhead"><button class="cookx" data-cook="close" aria-label="Close cook mode">×</button>
      <div class="cooktabs">${meal.combined ? `<button class="chip-all${cook.dish === -1 ? ' on' : ''}" data-cook="dish:-1">All</button>` : ''}${meal.dishes.map((x, i) => `<button class="chip-${x.color}${i === cook.dish ? ' on' : ''}" data-cook="dish:${i}">${esc(cookTabName(meal, x))}</button>`).join('')}</div></div>
    <div class="cookdishrow"><p class="cookdish c${d.color}"><span>${esc(d.type)}</span><b class="cd-name">${esc(d.name)}</b></p>${d.groups.length ? `<button class="btn ghost cookingsbtn" data-cook="ings">${ICONS.list} Ingredients</button>` : ''}</div>
    <div class="cooksheet" id="cooksheet" hidden><div class="cs-back" data-cook="ingsclose"></div><div class="cs-panel"><div class="cs-head"><strong>Ingredients · ${esc(shortName(d))}</strong><button class="cookx" data-cook="ingsclose" aria-label="Close">×</button></div>
      <div class="cs-body">${d.groups.map((g) => `${g.name ? `<h4>${esc(g.name)}</h4>` : ''}<ul class="list ings">${g.items.map((it) => `<li class="ing">${it.amount ? `<span class="amt">${esc(it.amount)}</span> ` : ''}${esc(it.text)}</li>`).join('')}</ul>`).join('')}</div></div></div>
    <div class="cookwheel" id="cookwheel"><div class="cw-pad"></div>
      ${d.steps.map((st, i) => { const mins = stepMinutes(st.text); return `<div class="cw-step" data-cwi="${i}" data-cwkey="${esc(st.key)}">
        <p class="cookcount">${st.dish ? `<span class="cw-dish chip-${st.dish.color}">${esc(cookTabName(meal, st.dish))}</span>` : ''}Step ${esc(st.num)} of ${n}<span class="cw-donemark"> · done</span></p>
        <p class="cooktext">${esc(st.text)}</p>${st.amount ? `<p class="cookused">${esc(st.amount)}</p>` : ''}
        ${mins ? `<button class="btn ghost cooktimer" data-cook="timer" data-cookstep="${esc(st.key)}" data-mins="${mins}">${ICONS.timer} Start ${mins} min timer</button>` : ''}
        <button class="btn ghost cw-untick" data-cook="untick">↺ Mark not done</button></div>`; }).join('')}
      <div class="cw-step cw-end" data-cwi="${n}"><p class="cookdone">${esc(d.name)} is done 🎉</p>${others}</div>
      <div class="cw-pad"></div></div>
    <div class="timerbar cooktimers" id="cooktimers" hidden></div>
    <div class="cookfoot"><button class="btn ghost" data-cook="back" id="cookback">‹ Back</button><button class="btn" data-cook="done" id="cooknextbtn">Next ›</button></div>`;
  const on = el.querySelector('.cooktabs .on'); // keep the chosen dish's tab in view
  if (on) on.scrollIntoView({ block: 'nearest', inline: 'center' });
  renderTimers();
  fitCookName();
  sizeWheel(); wheelTo(cook.step, false);
  const w = $('#cookwheel');
  w.addEventListener('scroll', onWheelScroll, { passive: true });
  // a hard swipe keeps all its momentum; once the wheel stops, it settles on the nearest step
  if ('onscrollend' in window) w.addEventListener('scrollend', settleWheel);
  else w.addEventListener('scroll', () => { clearTimeout(settleTimer); settleTimer = setTimeout(settleWheel, 140); }, { passive: true });
  w.addEventListener('touchstart', () => { fingerDown = true; }, { passive: true });
  w.addEventListener('touchend', () => { fingerDown = false; setTimeout(() => { if (Date.now() - lastWheelScroll > 120) settleWheel(); }, 150); }, { passive: true }); // let go without a fling
}
// Cook mode's dish tabs use just the key food word: "Lemon-Herb Chicken Cutlets" → "Chicken", "Roasted Baby
// Potatoes" → "Potatoes", "Chicken Pot Pie with Biscuit Topping" → "Pot Pie". If two dishes in the meal would
// get the same word, they keep their longer names.
const TAB_PAIRS = ['green beans', 'pot pie', 'garlic bread', 'lettuce wraps', 'pork belly', 'sweet potatoes', 'sweet potato', 'fried rice', 'black beans', 'brussels sprouts', 'mashed potatoes', 'french fries', 'mac and cheese'];
const TAB_SHAPES = new Set(['wedges', 'fries', 'cutlets', 'bites', 'skewers', 'strips', 'fillets', 'filets', 'tenders', 'medallions', 'kebabs', 'chunks', 'slices', 'pieces', 'patties', 'thighs', 'breasts']);
const titleCase = (t) => t.replace(/\b\w/g, (c) => c.toUpperCase());
function keyWord(name) {
  const w = name.replace(/\s+with\s+.*$/i, '').split(',')[0].trim().split(/\s+/);
  while (w.length > 1 && TAB_SHAPES.has(w[w.length - 1].toLowerCase())) w.pop(); // "Steak Bites" → "Steak", "Potato Wedges" → "Potato"
  const n = w.join(' '), low = n.toLowerCase();
  for (const p of TAB_PAIRS) if (low.endsWith(p)) return titleCase(n.slice(n.length - p.length));
  return titleCase(w[w.length - 1].replace(/^[^A-Za-z]+|[^A-Za-z]+$/g, '')) || shortName({ name });
}
function cookTabName(meal, d) {
  const k = keyWord(d.name);
  return meal.dishes.filter((x) => keyWord(x.name) === k).length > 1 ? shortName(d) : k;
}
function fitCookName() { // shrink the dish name (20px down to 16px) to fit one line; if it still can't, wrap it evenly
  const el = $('#cookmode .cd-name'); if (!el) return;
  el.style.fontSize = ''; el.style.whiteSpace = 'nowrap';
  const ps = getComputedStyle(el.parentElement), avail = el.parentElement.clientWidth - parseFloat(ps.paddingLeft) - parseFloat(ps.paddingRight) - 2;
  const base = parseFloat(getComputedStyle(el).fontSize), need = el.scrollWidth;
  if (need > avail) el.style.fontSize = Math.max(16, Math.floor(base * avail / need * 10) / 10) + 'px';
  el.style.whiteSpace = '';
}
function sizeWheel() { // room above the first step and below the last, so each can sit in the middle
  const w = $('#cookwheel'); if (!w) return;
  w.querySelectorAll('.cw-pad').forEach((x) => { x.style.height = Math.max(0, w.clientHeight / 2 - 60) + 'px'; });
}
addEventListener('resize', () => { if (cook) { sizeWheel(); wheelTo(cook.step, false); } });
function wheelTo(i, smooth) {
  const w = $('#cookwheel'); const el = w && w.querySelector(`[data-cwi="${i}"]`); if (!el) return;
  w.scrollTo({ top: el.offsetTop - (w.clientHeight - el.offsetHeight) / 2, behavior: smooth ? 'smooth' : 'auto' });
  markWheel(i);
}
function markWheel(cur) { // size/fade each step by its distance from the current one; done ones turn green
  if (!cook) return; const w = $('#cookwheel'); if (!w) return;
  cook.step = cur;
  const t = state.ticks[cook.meal] || {}, last = w.querySelectorAll('.cw-step').length - 1;
  w.querySelectorAll('.cw-step').forEach((el) => {
    const dist = Math.min(3, Math.abs(+el.dataset.cwi - cur));
    for (let k = 0; k < 4; k++) el.classList.toggle('cw-d' + k, k === dist);
    if (el.dataset.cwkey) el.classList.toggle('ticked', !!t[el.dataset.cwkey]);
  });
  const back = $('#cookback'), next = $('#cooknextbtn');
  if (back) back.disabled = cur === 0;
  if (next) next.hidden = cur >= last;
}
let wheelRaf = 0, settleTimer = 0, fingerDown = false, lastWheelScroll = 0;
function settleWheel() {
  const w = $('#cookwheel'); if (!w || !cook || fingerDown) return;
  const el = w.querySelector(`[data-cwi="${cook.step}"]`); if (!el) return;
  const off = el.offsetTop - (w.clientHeight - el.offsetHeight) / 2 - w.scrollTop;
  if (Math.abs(off) > 2) w.scrollTo({ top: w.scrollTop + off, behavior: 'smooth' });
}
function onWheelScroll() { // the step nearest the middle becomes the current one
  lastWheelScroll = Date.now();
  cancelAnimationFrame(wheelRaf);
  wheelRaf = requestAnimationFrame(() => {
    const w = $('#cookwheel'); if (!w || !cook) return;
    const mid = w.scrollTop + w.clientHeight / 2; let best = cook.step, bd = Infinity;
    w.querySelectorAll('.cw-step').forEach((el) => { const dd = Math.abs(el.offsetTop + el.offsetHeight / 2 - mid); if (dd < bd) { bd = dd; best = +el.dataset.cwi; } });
    if (best !== cook.step) markWheel(best);
  });
}
function cookAction(a) {
  if (!cook) return;
  const meal = mealById(cook.meal); const d = cookDishOf(meal);
  if (a === 'ings') { $('#cooksheet').hidden = false; return; }
  if (a === 'ingsclose') { $('#cooksheet').hidden = true; return; }
  if (a === 'close') { // back to the recipe, at the step you were on
    const cur = d.steps[cook.step], anchor = cur ? (cur.dish ? `${cur.dish.anchor}-s${cur.dishNum}` : `${d.anchor}-s${cur.num}`) : d.combined ? `${meal.id}-combined` : d.anchor; cook = null; renderCook(); route();
    const sec = document.getElementById(anchor); if (sec) scrollTo(0, sec.getBoundingClientRect().top + scrollY - 120);
    return;
  }
  if (a.startsWith('dish:')) { cook.dish = +a.slice(5); cook.step = null; renderCook(); return; }
  if (a === 'back') { wheelTo(Math.max(0, cook.step - 1), true); return; }
  const st = d.steps[cook.step]; if (!st) return;
  if (a === 'timer') { timerTap(timerLabel(...stepDish(d, st)), stepMinutes(st.text), meal.id, st.key); return; }
  if (a === 'untick') { setTick(meal.id, st.key, false); markWheel(cook.step); return; }
  if (a === 'done') { if (!(state.ticks[meal.id] || {})[st.key]) setTick(meal.id, st.key, true); wheelTo(cook.step + 1, true); }
}

// Last made and times made: calendar days before today, plus the dish's "Made" row in the spreadsheet
// (times made before the calendar). A side served in a current meal also counts that meal's calendar days.
const pastDays = (rid) => { const today = dayKey(new Date()); return [...new Set(planList().filter(([k, v]) => k < today && planKey(v) === rid).map(([k]) => k))].sort(); };
function lastMade(meal) { return pastDays(meal.rid).pop() || null; }
// calendar + spreadsheet. A side in a meal without its own Made row shares the main's (they were made together).
const countedMade = (meal, dish = meal.main) => pastDays(meal.rid).length + (dish.madeBefore != null ? dish.madeBefore : (meal.main.madeBefore || 0));
const madeCount = (meal, dish = meal.main) => Math.max(0, countedMade(meal, dish) + (Number(state.madeAdjust[dishKey(dish)]) || 0)); // plus your change
function writeAdjust(id, delta) {
  if (delta) state.madeAdjust[id] = delta; else delete state.madeAdjust[id];
  route();
  if (!state.firebase) { ls.set('local-adjust', state.madeAdjust); return; }
  setDoc(doc(state.firebase.db, 'history', 'adjust'), { a: { [id]: delta || deleteField() } }, { merge: true }).catch((e) => toast('Could not save: ' + (e.code || e.message)));
}
async function madeMenu(meal, dish) { // press and hold a Recipe Book line: times made (and your rating)
  const id = dishKey(dish), counted = countedMade(meal, dish);
  const rateable = dish === meal.main; // a side served in a current meal shares the meal's rating
  const before = rateable ? ratingInfo(meal.name).mine : 0;
  let n = madeCount(meal, dish), stars = before;
  const dlg = formDialog(`<p class="picktitle">${esc(dish.name)}</p>
    <p class="small menulabel">Times made</p>
    <div class="stepper"><button type="button" class="btn ghost" data-step="-1" aria-label="One less">−</button><output id="madeval">${n}</output><button type="button" class="btn ghost" data-step="1" aria-label="One more">+</button></div>
    <p class="small menulabel">Photo</p>
    <div class="menuphoto" data-photo="${esc(id)}" data-name="${esc(dish.name)}" hidden></div>
    <div class="menuphotobtns"><button type="button" class="photobtn" data-addphoto="${esc(id)}" hidden>${ICONS.camera} Add photo</button><button type="button" class="photobtn" data-hasphoto="${esc(id)}" data-replacephoto="${esc(id)}" hidden>Replace</button><button type="button" class="photobtn" data-hasphoto="${esc(id)}" data-removephoto="${esc(id)}" data-name="${esc(dish.name)}" hidden>Remove</button></div>
    ${rateable ? `<p class="small menulabel">Your rating</p><div class="menustars">${[1, 2, 3, 4, 5].map((k) => `<button type="button" data-menustar="${k}" aria-label="${k} star${k > 1 ? 's' : ''}">★</button>`).join('')}</div>` : ''}
    <div class="row-btns"><button class="btn ghost" value="">Cancel</button><button class="btn" value="save">Save</button></div>`);
  const f = $('#formdlg');
  f.classList.add('centered'); // no keyboard here, so it can sit in the middle
  paintPhotos();
  const paint = () => f.querySelectorAll('[data-menustar]').forEach((b) => b.classList.toggle('on', +b.dataset.menustar <= stars));
  paint();
  f.querySelectorAll('[data-menustar]').forEach((b) => b.addEventListener('click', () => { const k = +b.dataset.menustar; stars = stars === k ? 0 : k; paint(); }));
  f.querySelectorAll('[data-step]').forEach((b) => b.addEventListener('click', () => { n = Math.max(0, Math.min(999, n + Number(b.dataset.step))); $('#madeval').textContent = n; }));
  const r = await dlg;
  if (r.value !== 'save') return;
  let saved = false;
  if (n !== madeCount(meal, dish)) { writeAdjust(id, n - counted); saved = true; }
  if (rateable && stars !== before) { setRating(meal, stars || null); saved = true; }
  if (saved) toast('Saved');
}
const timesText = (n) => (n === 1 ? 'Made once' : n === 2 ? 'Made twice' : `Made ${n} times`);
function lastMadeText(meal, dish = meal.main) {
  const k = lastMade(meal), n = madeCount(meal, dish);
  if (!n) return 'Not made yet';
  if (!k) return timesText(n) + (dish.madeNote ? ` · last ${dish.madeNote}` : '');
  const days = Math.round((keyToDate(dayKey(new Date())) - keyToDate(k)) / 86400000);
  const ago = days <= 1 ? 'yesterday' : days < 14 ? `${days} days ago` : days < 60 ? `${Math.round(days / 7)} weeks ago` : `${Math.round(days / 30)} months ago`;
  return `${timesText(n)} · last ${ago}`;
}
// Suggestions for a day: well rated, and not had in a while (never made counts as a while)
function suggestions(day) {
  const today = keyToDate(dayKey(new Date()));
  return allMeals().filter((x) => x.kind === 'Main' && !plannedElsewhere(x.rid, day) && planKey(state.plan[day]) !== x.rid)
    .map((x) => { const k = lastMade(x), ri = ratingInfo(x.name); const weeks = k ? (today - keyToDate(k)) / 604800000 : 8;
      return { x, score: (ri.done ? ri.avg : 3.5) + Math.min(weeks, 8) * 0.3 }; })
    .sort((a, b) => b.score - a.score).slice(0, 3).map((o) => o.x);
}

// whether the folded "Usually have" card is open, kept across redraws (ticking redraws the list)
let staplesOpen = false, shoppedOpen = false;
document.addEventListener('toggle', (e) => { if (e.target.matches && e.target.matches('details.shoppedcard')) shoppedOpen = e.target.open; }, true);
document.addEventListener('toggle', (e) => { if (e.target.matches && e.target.matches('details.staples')) staplesOpen = e.target.open; }, true);
// Staples: items we usually have. They sit in their own folded card and don't count toward the list.
const stapleId = (name) => 's-' + slug(name);
const isStaple = (name) => !!state.staples[stapleId(name)];
function toggleStaple(name) {
  const id = stapleId(name), on = !isStaple(name);
  if (on) state.staples[id] = { name, by: myEmail() }; else delete state.staples[id];
  toast(on ? `"${name}" moved to Usually have` : `"${name}" is back on the list`);
  if (!state.firebase) { ls.set('local-staples', state.staples); route(); return; }
  route();
  setDoc(doc(state.firebase.db, 'shopping', 'staples'), { s: { [id]: on ? { name, by: myEmail() } : deleteField() } }, { merge: true })
    .catch((e) => toast('Could not save: ' + (e.code || e.message)));
}
// press and hold a shopping item to mark it as a staple. Any movement or scrolling cancels the hold:
// once the page starts scrolling the phone stops reporting finger movement, so scroll and cancel events are watched too.
let holdTimer = null, held = false;
document.addEventListener('pointerdown', (e) => {
  const it = e.target.closest('.item[data-name], .bookrow[data-dish]'); if (!it) return;
  held = false; const x0 = e.clientX, y0 = e.clientY, sy = scrollY;
  const stop = () => { clearTimeout(holdTimer); for (const [t, f] of offs) document.removeEventListener(t, f, true); removeEventListener('scroll', stop, true); };
  const moved = (ev) => { const p = ev.touches ? ev.touches[0] : ev; if (!p || Math.hypot(p.clientX - x0, p.clientY - y0) > 8) stop(); };
  const offs = [['pointermove', moved], ['touchmove', moved], ['pointercancel', stop], ['pointerup', stop], ['touchend', stop]];
  for (const [t, f] of offs) document.addEventListener(t, f, { capture: true, passive: true });
  addEventListener('scroll', stop, { capture: true, passive: true });
  clearTimeout(holdTimer);
  holdTimer = setTimeout(() => { stop(); if (Math.abs(scrollY - sy) > 2) return; held = true; if (navigator.vibrate) navigator.vibrate(25);
    if (it.matches('.bookrow')) { const meal = mealById(it.dataset.meal); if (meal) madeMenu(meal, meal.dishes[+it.dataset.dish]); } else stapleMenu(it.dataset.name); }, 600);
});
document.addEventListener('contextmenu', (e) => { if (e.target.closest('.item[data-name], .bookrow')) e.preventDefault(); });
async function stapleMenu(name) {
  const on = isStaple(name);
  const r = await formDialog(`<p class="picktitle">${esc(name)}</p><p class="small" style="margin:0 0 12px">${on ? 'You marked this as something you usually have.' : 'Usually have this at home? Staples move to a folded "Usually have" card and don’t count toward the list.'}</p>
    <div class="row-btns"><button class="btn ghost" value="">Cancel</button><button class="btn" value="save">${on ? 'Put it back on the list' : 'We usually have this'}</button></div>`);
  if (r.value === 'save') toggleStaple(name);
}

// ---------------------------------------------------------------- automatic shopping list
// Built from the ingredients of every current meal, plus anything planned for today or later, combined by
// shopping name. Ticking a line marks it "bought" for the meals it's for (shopping/bought). A meal's bought
// marks are used up once one of its calendar days has passed since they were set (it's been cooked).
function shopMeals() {
  if (!state.model) return [];
  const today = dayKey(new Date()), seen = new Set(), out = [];
  const add = (m) => { if (m && !seen.has(m.rid)) { seen.add(m.rid); out.push(m); } };
  state.model.meals.filter(isPlannedMeal).forEach(add); // made meals drop off until they're planned again
  planList().filter(([k]) => k >= today).sort((a, b) => a[0].localeCompare(b[0])).forEach(([, v]) => add(mealByName(v)));
  return out;
}
// The automatic list is used once every included dish has shopping details on all its ingredients.
function autoShop() {
  const meals = shopMeals();
  return meals.length > 0 && meals.every((m) => m.dishes.every((d) => d.groups.every((g) => g.items.every((it) => it.shop))));
}
const shopKey = (item) => 'i-' + slug(item);
function cooked(rid) { // a calendar day for this meal has passed since it was bought for
  const b = state.bought[rid]; if (!b) return false;
  const today = dayKey(new Date());
  return planList().some(([k, v]) => k >= (b.since || '') && k < today && planKey(v) === rid);
}
const isBought = (rid, key) => !!(state.bought[rid] && state.bought[rid].i && state.bought[rid].i[key]) && !cooked(rid);
function shopLines() {
  const lines = new Map();
  for (const meal of shopMeals()) {
    for (const d of meal.dishes) for (const g of d.groups) for (const it of g.items) {
      if (!it.shop || it.shop.none) continue;
      const key = shopKey(it.shop.item);
      if (!lines.has(key)) lines.set(key, { key, item: it.shop.item, section: it.shop.section, parts: [] });
      lines.get(key).parts.push({ rid: meal.rid, meal, qty: it.shop.qty, unit: it.shop.unit });
    }
  }
  for (const l of lines.values()) {
    l.open = l.parts.filter((x) => !isBought(x.rid, l.key));
    l.done = l.open.length === 0;
    l.partly = !l.done && l.open.length < l.parts.length;
  }
  // meals marked as shopped for (every item bought): their lines leave the list and the counter
  const rids = new Set([...lines.values()].flatMap((l) => l.parts.map((x) => x.rid)));
  const shopped = new Set([...rids].filter((rid) => [...lines.values()].every((l) => !l.parts.some((x) => x.rid === rid) || isBought(rid, l.key))));
  for (const l of lines.values()) l.mealDone = l.parts.every((x) => shopped.has(x.rid));
  return [...lines.values()];
}
// Mark a whole meal as shopped for: all its items count as bought (until its calendar day has passed), with Undo.
function mealShopKeys(rid) { return shopLines().filter((l) => l.parts.some((x) => x.rid === rid)).map((l) => l.key); }
function mealShopped(rid) { const ks = mealShopKeys(rid); return ks.length > 0 && ks.every((k) => isBought(rid, k)); }
function saveMealBought(rid) {
  if (!state.firebase) { ls.set('local-bought', state.bought); return; }
  setDoc(doc(state.firebase.db, 'shopping', 'bought'), { b: { [rid]: state.bought[rid] || deleteField() } }, { mergeFields: [new FieldPath('b', rid)] })
    .catch((e) => toast('Could not save: ' + (e.code || e.message)));
}
function toggleMealShopped(rid) {
  const meal = allMeals().find((x) => x.rid === rid); if (!meal) return;
  const before = state.bought[rid] ? JSON.parse(JSON.stringify(state.bought[rid])) : null, on = !mealShopped(rid);
  if (on) {
    const old = state.bought[rid], today = dayKey(new Date());
    const cur = !old || cooked(rid) ? { since: today, i: {} } : { since: old.since || today, i: { ...(old.i || {}) } };
    for (const k of mealShopKeys(rid)) cur.i[k] = true;
    state.bought[rid] = cur;
  } else delete state.bought[rid];
  saveMealBought(rid); route();
  undoToast(on ? `${mealLabel(meal)}: all bought` : `${mealLabel(meal)} back on the list`, () => { if (before) state.bought[rid] = before; else delete state.bought[rid]; saveMealBought(rid); route(); });
}
function mealChipsHtml() {
  const meals = shopMeals().filter((x) => mealShopKeys(x.rid).length); if (!meals.length) return '';
  return `<div class="mealchips"><span class="small">Shopped for a meal? Tap it:</span>${meals.map((x) => { const on = mealShopped(x.rid);
    return `<button class="mchip${on ? ' on' : ''}" data-mealshopped="${esc(x.rid)}" aria-pressed="${on}">${on ? '✓ ' : ''}${esc(mealLabel(x))}</button>`; }).join('')}</div>`;
}
const FRACTIONS = [[0.25, '¼'], [0.5, '½'], [0.75, '¾'], [1 / 3, '⅓'], [2 / 3, '⅔']];
function fmtQty(q) {
  const whole = Math.floor(q + 1e-9), rest = q - whole;
  const f = FRACTIONS.find(([v]) => Math.abs(rest - v) < 0.02);
  if (rest < 0.02) return String(whole);
  return f ? (whole ? whole + f[1] : f[1]) : String(Math.round(q * 100) / 100);
}
const PLURAL = { clove: 'cloves', head: 'heads', bunch: 'bunches', can: 'cans', package: 'packages', jar: 'jars', bag: 'bags', box: 'boxes', stalk: 'stalks', sprig: 'sprigs', slice: 'slices', cup: 'cups', bottle: 'bottles', loaf: 'loaves', piece: 'pieces' };
const WHOLE_UNITS = new Set(['each', 'head', 'bunch', 'can', 'package', 'jar', 'bag', 'box', 'bottle', 'loaf', 'piece']);
function amountText(parts) {
  const byUnit = new Map(); let loose = false;
  for (const x of parts) {
    if (x.qty == null) { loose = true; continue; }
    byUnit.set(x.unit, (byUnit.get(x.unit) || 0) + x.qty);
  }
  const bits = [...byUnit.entries()].map(([u, q]) => {
    if (WHOLE_UNITS.has(u || 'each')) q = Math.ceil(q - 0.01); // half an onion or half a package: buy a whole one
    const unit = !u || u === 'each' ? '' : q > 1 && PLURAL[u] ? PLURAL[u] : u;
    return `${fmtQty(q)}${unit ? ' ' + unit : ''}`;
  });
  if (loose && !bits.length) bits.push('as needed');
  return bits.join(' + ');
}
const mealLabel = (m) => m.title; // short meal name, e.g. "Chicken Pot Pie", "Apple Muffins"
function toggleLine(key) {
  const l = shopLines().find((x) => x.key === key); if (!l) return;
  const today = dayKey(new Date()), changes = {};
  const set = (rid, on) => {
    const cur = cooked(rid) ? { since: today, i: {} } : { since: (state.bought[rid] && state.bought[rid].since) || today, i: { ...((state.bought[rid] && state.bought[rid].i) || {}) } };
    if (on) cur.i[key] = true; else delete cur.i[key];
    state.bought[rid] = cur; changes[rid] = cur;
  };
  if (l.done) l.parts.forEach((x) => set(x.rid, false)); else l.open.forEach((x) => set(x.rid, true));
  route();
  if (!state.firebase) { ls.set('local-bought', state.bought); return; }
  const b = {}; for (const [rid, v] of Object.entries(changes)) b[rid] = Object.keys(v.i).length ? v : deleteField();
  // replace each changed meal's record whole, so a reset after cooking doesn't keep old marks
  setDoc(doc(state.firebase.db, 'shopping', 'bought'), { b }, { mergeFields: Object.keys(b).map((rid) => new FieldPath('b', rid)) })
    .catch((e) => toast('Could not save: ' + (e.code || e.message)));
}
// once both lists are loaded: drop bought marks for meals that have been cooked
function tidyBought() {
  if (!state.model || (state.firebase && !boughtLoaded)) return;
  const gone = Object.keys(state.bought).filter((rid) => cooked(rid));
  if (!gone.length) return;
  gone.forEach((rid) => delete state.bought[rid]);
  if (!state.firebase) { ls.set('local-bought', state.bought); return; }
  setDoc(doc(state.firebase.db, 'shopping', 'bought'), { b: Object.fromEntries(gone.map((rid) => [rid, deleteField()])) }, { merge: true }).catch(() => {});
}
function renderAutoShopping(m) {
  const lines = shopLines();
  const order = ['other', 'meat', 'produce', 'dairy', 'frozen', 'bakery', 'pantry'];
  const rank = (x) => { const i = order.indexOf(x.toLowerCase()); return i < 0 ? 50 : i; };
  const sections = [...new Set([...lines.map((l) => l.section), ...Object.values(state.extras).map((x) => x.section || 'Other')])].sort((a, b) => rank(a) - rank(b));
  const row = (l) => `<li class="item gline${l.done ? ' done' : ''}" data-line="${esc(l.key)}" data-name="${esc(l.item)}" role="checkbox" aria-checked="${l.done}" tabindex="0"><span class="box">${CHECK_SVG}</span>
    <span class="body"><strong>${esc(l.item)}</strong>${amountText(l.done ? l.parts : l.open) ? ` · ${esc(amountText(l.done ? l.parts : l.open))}` : ''}${l.partly ? ' <span class="more">more</span>' : ''}
    <span class="sub">For: ${esc([...new Set((l.done ? l.parts : l.open).map((x) => mealLabel(x.meal)))].join(', '))}</span></span></li>`;
  let h = '';
  for (const sec of sections) {
    const mine = lines.filter((l) => l.section === sec && !isStaple(l.item) && !l.mealDone).sort((a, b) => a.done - b.done || a.item.localeCompare(b.item));
    const rows = mine.map(row).join('') + extrasListHtml(sec);
    if (rows) h += `<section class="card"><p class="section-label">${esc(sec)}</p><ul class="list">${rows}</ul></section>`;
  }
  const got = lines.filter((l) => l.mealDone && !isStaple(l.item)).sort((a, b) => a.item.localeCompare(b.item));
  if (got.length) h += `<details class="card shoppedcard"${shoppedOpen ? ' open' : ''}><summary class="section-label">Already shopped for (${got.length})</summary>
    <p class="small">Items for meals you marked as shopped for. They don't count toward the list.</p><ul class="list">${got.map(row).join('')}</ul></details>`;
  const staples = lines.filter((l) => isStaple(l.item)).sort((a, b) => a.item.localeCompare(b.item));
  const stapleExtras = extrasListHtml(null, true);
  if (staples.length || stapleExtras) h += `<details class="card staples"${staplesOpen ? ' open' : ''}><summary class="section-label">Usually have (${staples.length + (stapleExtras.match(/data-key=/g) || []).length})</summary>
    <p class="small">Things you usually have at home. They don't count toward the list. Tick one if you're buying it this time.</p><ul class="list">${staples.map(row).join('')}${stapleExtras}</ul></details>`;
  if (!lines.length && !Object.keys(state.extras).length) h += '<p class="empty">Nothing to buy. Plan a meal on the calendar, or tap + to add an item.</p>';
  return h;
}
function autoCounts() { // [done, total] for the counter: list lines (not staples) and items you added (not staples)
  const lines = shopLines().filter((l) => !isStaple(l.item) && !l.mealDone);
  const extras = Object.entries(state.extras).filter(([, x]) => !isStaple(x.name));
  const t = state.ticks[SHOP()] || {};
  return [lines.filter((l) => l.done).length + extras.filter(([id]) => t[extraKey(id)]).length, lines.length + extras.length];
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
const itemHtml = (scope, key, inner, extraClass = '', id = '', name = '') =>
  `<li class="item ${extraClass}" data-scope="${scope}" data-key="${key}" role="checkbox" aria-checked="false" tabindex="0"${id ? ` id="${id}"` : ''}${name ? ` data-name="${esc(name)}"` : ''}><span class="box">${CHECK_SVG}</span>${inner}</li>`;

function banners() {
  let h = '';
  if (state.dbError) h += `<div class="banner">${esc(state.dbError)}</div>`;
  if (state.listenErr) h += `<div class="banner">Couldn't load ${esc(state.listenErr.what)} (${esc(state.listenErr.code)}). Try closing and reopening the app; if it keeps happening, tell Claude this message.</div>`;
  return h;
}

function planningSwitch(active) {
  return `<nav class="seg" aria-label="Planning">
    <a href="#/planning" class="${active === 'book' ? 'on' : ''}">Recipe Book</a>
    <a href="#/ideas" class="${active === 'ideas' ? 'on' : ''}">Ideas</a></nav>`;
}
// Planned = on the calendar for today or later, or a current meal (Recipes tab) that hasn't been made yet.
// Used by the Meals list, the Recipe Book's Planned section and the shopping list.
const isPlannedMeal = (meal) => !!nextPlanned(meal.name) || (!meal.archived && !madeCount(meal));
const kindBadge = (meal) => `<span class="kind chip-${meal.kind === 'Side' ? 1 : meal.kind === 'Dessert' ? 3 : 0}">${meal.kind}</span>`;
function plannedChip(meal) {
  const n = nextPlanned(meal.name);
  return n ? `<span class="planned">${ICONS.cal} ${n === dayKey(new Date()) ? 'Tonight' : esc(fmtDay(keyToDate(n), { weekday: 'short', month: 'short', day: 'numeric' }))}</span>` : '';
}
// Recipe Book: every dish on its own line (mains and sides split), as a short index. A search box, then sections:
// Planned (on the Planned list or the calendar), Made before (best rated first) and Want to try, each with Mains then Sides.
function bookRowHtml(meal, dish, inPlanned) {
  const ri = dish === meal.main ? ratingInfo(meal.name) : { done: false };
  const n = madeCount(meal, dish), next = nextPlanned(meal.name);
  const words = [dish.name, ...dish.groups.flatMap((g) => g.items.map((x) => (x.shop && x.shop.item ? x.shop.item + ' ' : '') + x.text))].join(' ').toLowerCase();
  const meta = [dish !== meal.main ? `with ${esc(meal.main.name)}` : '', n ? `<b class="bk-made">${esc(lastMadeText(meal, dish))}</b>` : inPlanned ? '' : 'Not made yet',
    next ? `${ICONS.cal} ${next === dayKey(new Date()) ? 'Tonight' : esc(fmtDay(keyToDate(next), { weekday: 'short', month: 'short', day: 'numeric' }))}` : ''].filter(Boolean).join(' · ');
  const starsHtml = ri.done ? `<span class="bk-stars">★ ${ri.avg.toFixed(1)}</span>`
    : ri.mine ? `<span class="bk-stars mine" title="Your rating (waiting for the other vote)">★ ${ri.mine}</span>` : ''; // only your vote so far: paler
  const di = meal.dishes.indexOf(dish), th = state.thumbs[dishKey(dish)];
  return `<li class="bookli">${bookThumbs() ? `<span class="bk-thumb">${th ? `<img src="${th}" alt="">` : ''}</span>` : ''}<a class="bookrow" href="#/meal/${meal.id}" data-meal="${esc(meal.id)}" data-dish="${di}" data-search="${esc(words)}">
    <span class="bk-name">${esc(dish.name)}</span>${starsHtml}
    <span class="bk-meta">${meta}</span></a><button class="bookmore" data-bookmenu="1" data-meal="${esc(meal.id)}" data-dish="${di}" aria-label="More for ${esc(dish.name)}">⋯</button></li>`;
}
let bookQuery = '';
const BOOK_SORTS = [['rated', 'Best rated'], ['made', 'Most made'], ['ago', 'Not made lately']];
function madeSort() { // how "Made before" is ordered (this phone's choice)
  const by = ls.get('book-sort', 'rated'), name = (a, b) => a.dish.name.localeCompare(b.dish.name);
  if (by === 'made') return (a, b) => b.made - a.made || b.stars - a.stars || name(a, b);
  if (by === 'ago') return (a, b) => (lastMade(a.meal) || '').localeCompare(lastMade(b.meal) || '') || name(a, b); // never on the calendar counts as longest ago
  return (a, b) => b.stars - a.stars || b.made - a.made || name(a, b);
}
function renderPast() {
  setTitle('Planning', false); setTab('planning');
  const m = state.model;
  if (!m) { view().innerHTML = '<p class="loading">Loading recipes…</p>'; return; }
  let h = banners() + planningSwitch('book');
  // one entry per dish: current meals give their main and each side; archived entries are one dish already
  const all = [...m.meals.flatMap((meal) => meal.dishes.map((d) => ({ meal, dish: d, kind: d.type }))),
    ...(m.past || []).map((x) => ({ meal: x, dish: x.main, kind: x.kind }))].filter((e) => !e.dish.hideBook);
  all.forEach((e, i) => { e.i = i; e.made = madeCount(e.meal, e.dish); e.next = nextPlanned(e.meal.name); e.planned = isPlannedMeal(e.meal); // on the calendar soon, or its meal is on the Planned list and hasn't been made yet (a side skipped when the meal was made goes to Want to try)
    const ri = e.dish === e.meal.main ? ratingInfo(e.meal.name) : null; e.stars = ri && ri.done ? ri.avg : 0; });
  const groups = [
    ['planned', 'Planned', all.filter((e) => e.planned), (a, b) => (a.next || '9').localeCompare(b.next || '9') || a.i - b.i],
    ['made', 'Made before', all.filter((e) => !e.planned && e.made), madeSort()],
    ['try', 'Want to try', all.filter((e) => !e.planned && !e.made), (a, b) => a.dish.name.localeCompare(b.dish.name)],
  ];
  h += `<div class="booksearch"><input type="search" id="booksearch" placeholder="Search recipes or ingredients" value="${esc(bookQuery)}" autocomplete="off" enterkeyhint="search"></div>`;
  for (const [id, title, list, sort] of groups) {
    if (!list.length) continue;
    const part = (label, rows) => (rows.length ? `<li class="bk-sub" data-sub>${label}</li>${rows.sort(sort).map((e) => bookRowHtml(e.meal, e.dish, id === 'planned')).join('')}` : '');
    const sortRow = id === 'made' && list.length > 1 ? `<div class="booksort">${BOOK_SORTS.map(([v, label]) => `<button class="filter${ls.get('book-sort', 'rated') === v ? ' on' : ''}" data-booksort="${v}">${label}</button>`).join('')}</div>` : '';
    h += `<details class="card booksec" data-sec="${id}"${id !== 'planned' || ls.get('book-planned-open', false) ? ' open' : ''}><summary><span class="section-label">${title}</span><span class="bk-count">${list.length}</span></summary>${sortRow}
      <ul class="booklist">${part('Mains', list.filter((e) => e.kind === 'Main'))}${part('Sides', list.filter((e) => e.kind === 'Side'))}${part('Desserts', list.filter((e) => e.kind === 'Dessert'))}</ul></details>`;
  }
  h += '<p class="empty" id="booknone" hidden>No recipes match.</p>';
  h += '<p class="small staplehint">Tip: tap ⋯ (or press and hold a recipe) to rate it, add a photo or change how many times it’s been made.</p>';
  view().innerHTML = h;
  filterBook();
}
function filterBook() { // hide rows (and empty sections) that don't match the search
  const q = bookQuery.trim().toLowerCase().split(/\s+/).filter(Boolean);
  let any = false;
  document.querySelectorAll('.booksec').forEach((sec) => {
    let n = 0;
    sec.querySelectorAll('.bookrow').forEach((a) => { const ok = q.every((w) => a.dataset.search.includes(w)); a.parentElement.hidden = !ok; if (ok) n++; });
    sec.querySelectorAll('[data-sub]').forEach((li) => { let x = li.nextElementSibling, vis = false; while (x && !x.hasAttribute('data-sub')) { if (!x.hidden) vis = true; x = x.nextElementSibling; } li.hidden = !vis; });
    sec.hidden = !n; sec.querySelector('.bk-count').textContent = n; if (q.length && n) sec.open = true;
    any = any || n > 0;
  });
  const none = $('#booknone'); if (none) none.hidden = any;
}
document.addEventListener('toggle', (e) => { if (e.target.matches && e.target.matches('.booksec[data-sec="planned"]') && !bookQuery) ls.set('book-planned-open', e.target.open); }, true);
document.addEventListener('input', (e) => { if (e.target.id === 'booksearch') { bookQuery = e.target.value; filterBook(); } });

// Shrink each "with …" sides line just enough to fit on one line (between 12.5px and 13.5px); wraps if still too long.
function fitSides() {
  document.querySelectorAll('.meal-card .sides').forEach((el) => {
    el.style.fontSize = ''; el.style.whiteSpace = 'nowrap';
    const base = parseFloat(getComputedStyle(el).fontSize), need = base * el.clientWidth / el.scrollWidth;
    el.style.whiteSpace = '';
    if (need < base) el.style.fontSize = Math.max(12.5, Math.floor(need * 10) / 10) + 'px';
  });
}
addEventListener('resize', fitSides);
function mealCardHtml(meal, tonight, label = 'Tonight') { // tonight: today's highlighted card; label: Tonight / For lunch / Dessert ...
  const ri = ratingInfo(meal.name);
  const sides = meal.sides.map((d) => esc(d.name)).join(' &amp; ');
  const time = meal.planTotal ? esc(meal.planTotal.split(/[.,;]/)[0]) : meal.totalMin != null ? `≈ ${fmtDuration(meal.totalMin)}` : '';
  const timeAndProgress = `${time ? `<span class="time">${ICONS.timer} ${time}</span>` : ''}${progressHtml(meal.id, meal.keys)}`;
  return `<a class="card meal-card${tonight ? ' tonight-card' : ri.done ? ' rated' : ''}" href="#/meal/${meal.id}">
      <div class="eyebrow">${tonight ? `<span class="tonight-label">${esc(label)}</span>` : `${esc(meal.title)}${plannedChip(meal)}`}</div>
      <h2>${esc(meal.main.name)}</h2>
      ${sides ? `<p class="sides">with ${sides}</p>` : ''}
      ${tonight || !madeCount(meal) ? '' : `<p class="lastmade">${esc(lastMadeText(meal))}</p>`}
      ${tonight || nextPlanned(meal.name) ? '' : `<span class="pickday" role="button" tabindex="0" data-pickday="${esc(meal.id)}">${ICONS.cal} Pick a day</span>`}
      <div class="meta">${tonight ? timeAndProgress : ri.done ? `<span class="avg">${starsText(ri.avg)}</span>` : timeAndProgress}</div>
      ${tonight ? headStartReminder(meal) : ''}
      ${tonight ? (() => { const first = meal.starts.length ? meal.starts[0].dish : meal.dishes.find((x) => x.steps.length); const i = first ? meal.dishes.indexOf(first) : -1;
        return i < 0 ? '' : `<span class="tonightcook" role="button" tabindex="0" data-tonightcook="${esc(meal.id)}" data-dish="${i}"><span class="tc-main">${ICONS.cook} Start cooking</span>${meal.dishes.length > 1 ? `<span class="tc-sub">${esc(first.name)} first</span>` : ''}</span>`; })() : ''}
      ${tonight && ri.done ? `<p class="small votes"><span class="avg">${starsText(ri.avg)}</span> last time</p>` : ''}
      ${!ri.done && ri.count ? `<p class="small votes">${ri.mine ? 'You rated it. Waiting for the other vote.' : 'Rated by the other person. Add your vote.'}</p>` : ''}
    </a>`;
}
// The day after a meal (up to 3 days), ask for your rating right on the Meals screen.
function ratePromptHtml() {
  const today = dayKey(new Date()), dismissed = ls.get('rate-dismissed', {});
  for (const [k, v, slot] of planList().filter(([x]) => x < today).sort((a, b) => b[0].localeCompare(a[0]) || (a[2] === 'dinner' ? -1 : 1))) {
    const ago = Math.round((keyToDate(today) - keyToDate(k)) / 86400000); if (ago > 3) break;
    const dk = slot === 'dinner' ? k : `${k}-${slot}`;
    const meal = mealByName(v); if (!meal || dismissed[dk] || ratingInfo(meal.name).mine) continue;
    const when = ago === 1 ? (slot === 'dinner' ? 'last night' : 'yesterday') : 'on ' + fmtDay(keyToDate(k), { weekday: 'long' });
    return `<section class="card rateprompt"><button class="rp-x" data-ratedismiss="${dk}" aria-label="Not now">×</button>
      <p class="rp-q">How was <strong>${esc(meal.main.name)}</strong> ${esc(when)}?</p>
      <div class="menustars">${[1, 2, 3, 4, 5].map((n) => `<button type="button" data-quickrate="${n}" data-meal="${esc(meal.id)}" aria-label="${n} star${n > 1 ? 's' : ''}">★</button>`).join('')}</div></section>`;
  }
  return '';
}
// "Pick a day" on a planned meal with no calendar day yet: the next two weeks, one tap to plan it.
async function pickDayFor(meal) {
  const today = new Date(), days = [];
  for (let i = 0; i < 14; i++) { const d = new Date(today); d.setDate(d.getDate() + i); days.push(d); }
  const dlg = formDialog(`<p class="picktitle">${esc(meal.main.name)}</p><p class="small" style="margin:0 0 10px">Which day?</p>
    <div class="picklist daylist">${days.map((d, i) => { const k = dayKey(d), cur = state.plan[k], other = customText(cur) || (mealByName(cur) ? mealByName(cur).main.name : '');
      return `<button class="pickopt" value="${k}"><strong>${i === 0 ? 'Tonight' : i === 1 ? 'Tomorrow' : esc(fmtDay(d, { weekday: 'long' }))} <span class="small">${esc(fmtDay(d, { month: 'short', day: 'numeric' }))}</span></strong>${other ? `<span class="small">Replaces ${esc(other)}</span>` : ''}</button>`; }).join('')}</div>
    <div class="row-btns"><button class="btn ghost" value="">Cancel</button></div>`);
  $('#formdlg').classList.add('centered');
  const r = await dlg;
  if (!r.value) return;
  setPlan(r.value, meal.rid);
  toast(`Planned for ${r.value === dayKey(today) ? 'tonight' : fmtDay(keyToDate(r.value), { weekday: 'long' })}`);
}
function renderMeals() {
  setTitle('Meals', false); setTab('meals');
  const m = state.model;
  if (!m) { view().innerHTML = '<p class="loading">Loading recipes…</p>'; return; }
  let h = banners();
  if (!state.firebase) h += '<div class="banner">Sharing is not set up yet, so ticks stay on this phone. See SETUP.md.</div>';
  h += ratePromptHtml();
  // Everything planned for today gets a highlighted card at the top (and isn't repeated below).
  // Just one thing: "Tonight" for dinner, "For lunch" etc. Several: a "Today" heading and each card says which meal.
  const todayK = dayKey(new Date()), todays = SLOTS.map(([sl, name]) => [sl, name, slotValue(todayK, sl)]).filter((x) => x[2]);
  const one = todays.length === 1;
  if (!one && todays.length) h += '<p class="section-label todaylabel">Today</p>';
  const shownToday = new Set();
  for (const [sl, name, v] of todays) {
    const label = one ? (sl === 'dinner' ? 'Tonight' : `For ${name.toLowerCase()}`) : name, mm = mealByName(v);
    if (mm && !shownToday.has(mm)) { shownToday.add(mm); h += mealCardHtml(mm, true, label); }
    else if (customText(v)) h += `<a class="card tonight" href="#/calendar"><span class="section-label">${esc(label)}</span><strong>${esc(customText(v))}</strong></a>`;
  }
  const tonight = todays.length ? true : null;
  // Order: meals planned for today or later (soonest first), then unplanned meals, then rated meals.
  const rank = (meal) => { const n = nextPlanned(meal.name); return n ? [0, n] : [ratingInfo(meal.name).done ? 2 : 1, '']; };
  const sorted = [...m.meals, ...(m.past || [])].filter((meal) => !shownToday.has(meal) && isPlannedMeal(meal)).map((meal, i) => ({ meal, i, r: rank(meal) }))
    .sort((a, b) => a.r[0] - b.r[0] || a.r[1].localeCompare(b.r[1]) || a.i - b.i).map((x) => x.meal);
  for (const meal of sorted) h += mealCardHtml(meal, false);
  if (!sorted.length && !tonight) h += '<p class="empty">Nothing planned right now.<br><span class="small">Pick something from Planning → Recipe Book and put it on the calendar.</span></p>';
  view().innerHTML = h;
  fitSides();
  refreshTicks();
}

// Prep list: steps marked with a Prep task, gathered across the main and sides and combined by task
// ("Mince garlic: 2 cloves · Chicken + 3 cloves · Garlic bread"). Ticking a task ticks those steps in each dish.
// ---------------------------------------------------------------- Head start: steps to do well before cooking
// Steps with an "Ahead" time in the spreadsheet (marinate 1 hour, soak skewers 30 min, thaw overnight), longest first.
function aheadMinutes(t) {
  const x = String(t || '').toLowerCase();
  if (/overnight/.test(x)) return 12 * 60;
  const n = parseFloat((x.match(/\d+(\.\d+)?/) || [])[0]); if (isNaN(n)) return 0;
  return /day/.test(x) ? n * 1440 : /h(ou)?r/.test(x) ? n * 60 : n;
}
const aheadLabel = (t) => { const x = String(t).trim(); return /overnight/i.test(x) ? 'The night before' : `${x} before`; };
function headStart(meal) { return meal.dishes.flatMap((d) => d.steps.filter((st) => st.ahead).map((st) => ({ d, st }))).sort((a, b) => aheadMinutes(b.st.ahead) - aheadMinutes(a.st.ahead)); }
function headStartHtml(meal) {
  const list = headStart(meal); if (!list.length) return '';
  const t = state.ticks[meal.id] || {}, allDone = list.every((x) => t[x.st.key]);
  return `<details class="card headstart"${allDone ? '' : ' open'} id="${meal.id}-ahead" data-autofold="ahead" data-meal="${esc(meal.id)}"><summary class="section-label">${ICONS.bell} Head start · do these early</summary>${meal.aheadNotes && meal.aheadNotes.length ? `<p class="aheadsum">${esc(headStartSummary(meal))}</p>` : ''}
    <ul class="list">${list.map(({ d, st }) => itemHtml(meal.id, st.key, `<span class="body"><span class="aheadwhen">${esc(aheadLabel(st.ahead))}</span><strong>${esc(st.text)}</strong>${st.amount ? `<span class="used">${esc(st.amount)}</span>` : ''}<span class="sub"><span class="chip-${d.color} prepdish">${esc(cookTabName(meal, d))}</span> step ${esc(st.num)}</span></span>`, 'aheaditem')).join('')}</ul></details>`;
}
// A short summary of the head start for the Meals card and the top of the Head start card: the spreadsheet's
// "Ahead" rows if there are any ("Marinate the steak · about 1 hour ahead"); otherwise the early step that actually
// waits (marinate / rest / thaw), plus how many other early steps go with it.
const aheadWhen = (t) => (/overnight/i.test(t) ? 'the night before' : t ? `about ${String(t).trim()} ahead` : '');
function headStartSummary(meal) {
  if (meal.aheadNotes && meal.aheadNotes.length) return meal.aheadNotes.map((n) => `${n.text}${n.when ? ` · ${aheadWhen(n.when)}` : ''}`).join('; ');
  const list = headStart(meal); if (!list.length) return '';
  const waits = list.find(({ st }) => /\b(marinat|thaw|soak|chill|brine)/i.test(st.text)) || list.find(({ st }) => stepMinutes(st.text))
    || list.find(({ st }) => /\b(rest|sit)\b(?! of)/i.test(st.text)) || list[0];
  const others = list.length - 1, first = waits.st.text.replace(/\s*\(.*?\)\s*/g, ' ').split(/(?<=\.)\s/)[0].replace(/\s+/g, ' ').replace(/\.$/, '').trim();
  return `${first}${others ? ` (+${others} more step${others > 1 ? 's' : ''})` : ''} · ${aheadWhen(waits.st.ahead)}`;
}
function headStartReminder(meal) { // today's card on Meals: until every early step is ticked
  const t = state.ticks[meal.id] || {}; if (headStart(meal).every((y) => t[y.st.key])) return '';
  const sum = headStartSummary(meal);
  return sum ? `<p class="aheadnote">${ICONS.bell}<span><strong>Head start:</strong> ${esc(sum)}</span></p>` : '';
}
function prepTasks(meal) {
  const tasks = new Map();
  meal.dishes.forEach((d) => d.steps.forEach((st) => {
    if (!st.prep || st.ahead) return; // head-start steps have their own card
    const k = st.prep.trim().toLowerCase();
    if (!tasks.has(k)) tasks.set(k, { name: st.prep.trim(), parts: [] });
    tasks.get(k).parts.push({ dish: d, amount: st.prepAmt, key: st.key, num: st.num });
  }));
  return [...tasks.values()];
}
function prepByDish(t) { // [[dish, [amounts]]], one entry per dish
  const m = new Map();
  for (const x of t.parts) { if (!m.has(x.dish)) m.set(x.dish, []); if (x.amount) m.get(x.dish).push(x.amount); }
  return [...m.entries()];
}
function prepCardHtml(meal) {
  const tasks = prepTasks(meal);
  if (!tasks.length) return '';
  const allPrepped = tasks.every((t) => t.parts.every((x) => (state.ticks[meal.id] || {})[x.key]));
  return `<details class="card prep"${allPrepped ? '' : ' open'} id="${meal.id}-prep" data-autofold="prep" data-meal="${esc(meal.id)}"><summary class="section-label">Prep list · do these first</summary>
    <ul class="list">${tasks.map((t, ti) => `<li class="item preptask" id="${esc(meal.id)}-p${ti}" data-prepkeys="${esc(t.parts.map((x) => x.key).join(','))}" data-scope="${esc(meal.id)}" role="checkbox" aria-checked="false" tabindex="0"><span class="box">${CHECK_SVG}</span>
      <span class="body"><strong>${esc(t.name)}</strong><span class="sub">${prepByDish(t).map(([d, amts]) => `${amts.length ? esc(amts.join(' + ')) + ' · ' : ''}<span class="chip-${d.color} prepdish">${esc(shortName(d))}</span>`).join(' + ')}</span></span>
      <button class="prepgo" data-prepgo="${esc(prepByDish(t).map(([d]) => `${d.anchor}-s${t.parts.find((x) => x.dish === d).num}|${d.type + ': ' + d.name.replace(/;|\|/g, ' ')}`).join(';'))}" aria-label="Go to the step">↓</button></li>`).join('')}</ul></details>`;
}

function renderMeal(id) {
  const meal = mealById(id);
  if (!meal) { if (state.model) location.replace('#/'); else view().innerHTML = '<p class="loading">Loading…</p>'; return; }
  setTitle(meal.title, true); setTab('meals');
  const S = meal.id;
  const hasPlan = !!(meal.planTotal || meal.planFirst || meal.starts.length > 1);
  let h = '<nav class="jump">' + (hasPlan ? '<a href="#" data-jump="' + S + '-plan">Plan</a>' : '') + (prepTasks(meal).length ? '<a href="#" data-jump="' + S + '-prep">Prep</a>' : '') +
    (meal.combined ? '<a href="#" class="jump-all" data-jump="' + S + '-combined">All</a>' : '') +
    meal.dishes.map((d) => `<a href="#" class="chip-${d.color}" data-jump="${d.anchor}">${esc(cookTabName(meal, d))}</a>`).join('') + '</nav>';
  h += banners();
  if (meal.archived) h += '<div class="banner">This recipe is in the Recipe Book. You can cook it, tick it and put it on the calendar; while it’s on the calendar it shows on the Meals list too.</div>';
  { const ri = ratingInfo(meal.name); h += `<div class="listhead">${progressHtml(S, meal.keys)}<a href="#" data-jump="${meal.id}-rate" class="rate-link">${ri.done ? starsText(ri.avg) : ri.mine ? '★ Rated' : '☆ Rate'}</a></div>`; }
  h += nextTimeTopHtml(meal);

  // ---- Plan first
  if (meal.planTotal || meal.planFirst || meal.starts.length > 1) {
    h += `<details class="card plan" id="${S}-plan" data-autofold="plan" data-meal="${esc(S)}"${planStarted(meal) ? '' : ' open'}><summary class="plan-sum">${planHeadHtml(meal)}</summary>`;
    if (meal.planFirst) h += `<div class="startfirst"><strong>${ICONS.warn} Start first</strong>${esc(meal.planFirst)}</div>`;
    if (meal.starts.length > 1) {
      const mainStart = (meal.starts.find((s) => s.dish === meal.main) || {}).min;
      h += '<p class="section-label">Start order</p><ol class="startorder">' + meal.starts.map((s, i) => {
        const early = i === 0 && s.dish !== meal.main && mainStart != null && s.min < mainStart;
        return `<li class="chip-${s.dish.color}" data-jump="${s.dish.anchor}">${early ? '<span class="first-badge">First</span>' : ''}${esc(s.dish.name)}<span class="at">${s.min === 0 ? 'at 0:00' : 'at ' + Math.floor(s.min / 60) + ':' + String(s.min % 60).padStart(2, '0')}</span></li>`;
      }).join('') + '</ol>';
    }
    h += '</details>';
  }

  h += headStartHtml(meal);
  h += prepCardHtml(meal);
  const prepOf = new Map(); prepTasks(meal).forEach((t, ti) => t.parts.forEach((x) => prepOf.set(x.key, `${meal.id}-p${ti}`)));
  headStart(meal).forEach(({ st }) => prepOf.set(st.key, `${meal.id}-ahead`)); // the ↑ on a head-start step goes to that card

  // ---- steps all together, in cooking order (when the meal has a combined order and that view is chosen)
  const combinedView = !!meal.combined && ls.get('combined-view', true);
  if (meal.combined) h += `<nav class="seg viewseg" aria-label="Steps"><a href="#" data-stepview="split" class="${combinedView ? '' : 'on'}">By dish</a><a href="#" data-stepview="combined" class="${combinedView ? 'on' : ''}">Combined</a></nav>`;
  if (combinedView) {
    const ingAnchor = `${S}-combined`; // all the dishes' ingredients, grouped by dish, folded like each dish's own list
    const ings = meal.dishes.filter((x) => x.groups.length).map((x) => x.groups.map((g) => `<h4><span class="stepdish chip-${x.color}">${esc(cookTabName(meal, x))}</span>${g.name ? esc(g.name) : ''}</h4><ul class="list ings">${g.items.map((it) =>
      `<li class="ing">${it.amount ? `<span class="amt">${esc(it.amount)}</span> ` : ''}${esc(it.text)}</li>`).join('')}</ul>`).join('')).join('');
    const allDone = meal.combined.every((x) => (state.ticks[meal.id] || {})[x.s.key]);
    h += `<details class="card combined" id="${S}-combined" data-autofold="all" data-meal="${esc(meal.id)}"${allDone ? '' : ' open'}><summary class="comb-sum"><div class="dish-head comb-head"><p class="section-label">Combined</p><h2>All steps, in order</h2></div></summary>
      <div class="dishbtns"><button class="cookdishbtn" data-opencook="${esc(meal.id)}" data-dish="-1">${ICONS.cook} Cook mode</button></div>
      ${ings ? `<details class="ingfold" data-ings="${ingAnchor}"${openIngs.has(ingAnchor) ? ' open' : ''}><summary><h3>Ingredients</h3></summary>${ings}</details>` : ''}
      ${stepsHeadHtml(meal, meal.combined.map((x) => x.s.key))}<ul class="list">${meal.combined.map((x) => stepItemHtml(meal, x.d, x.s, prepOf, true)).join('')}</ul></details>`;
  }

  // ---- dishes (side by side on wide screens)
  h += `<div class="dishes cols-${meal.dishes.length}">`;
  for (const d of meal.dishes) {
    const pk = dishKey(d);
    h += `<section class="card dish${foldedDishes.has(d.anchor) ? ' folded' : ''}" id="${d.anchor}" data-dishkeys="${esc(d.steps.map((x) => x.key).join(','))}" data-meal="${esc(meal.id)}"><div class="dishphoto" data-photo="${esc(pk)}" data-name="${esc(d.name)}" hidden></div><div class="dish-head foldable c${d.color}" title="Fold or open this dish"><p class="section-label">${d.type}</p><h2>${esc(d.name)}</h2>${d.makes ? `<p class="makes">${esc(d.makes)}</p>` : ''}<div class="dishbtns">${d.steps.length ? `<button class="cookdishbtn" data-opencook="${esc(meal.id)}" data-dish="${meal.dishes.indexOf(d)}">${ICONS.cook} Cook mode</button>` : ''}<button class="photobtn" data-addphoto="${esc(pk)}" hidden>${ICONS.camera} Add photo</button></div></div>`;
    for (const t of d.tips) h += `<div class="tip tip-top"><strong>Tip:</strong> ${esc(t)}</div>`; // read before cooking, not after
    if (d.groups.length) {
      h += `<details class="ingfold" data-ings="${esc(d.anchor)}"${openIngs.has(d.anchor) ? ' open' : ''}><summary><h3>Ingredients</h3></summary>`;
      for (const g of d.groups) {
        if (g.name) h += `<h4>${esc(g.name)}</h4>`;
        h += '<ul class="list ings">' + g.items.map((it) =>
          `<li class="ing">${it.amount ? `<span class="amt">${esc(it.amount)}</span> ` : ''}${esc(it.text)}</li>`).join('') + '</ul>';
      }
      h += '</details>';
    }
    if (d.steps.length && !combinedView) h += stepsHeadHtml(meal, d.steps.map((x) => x.key)) + '<ul class="list">' + d.steps.map((x) => stepItemHtml(meal, d, x, prepOf, false)).join('') + '</ul>';
    h += '</section>';
  }
  h += '</div>';
  h += ratingCardHtml(meal);
  h += nextTimeCardHtml(meal);
  h += `<button class="btn outline-danger" data-clear="${S}" data-label="${esc(meal.title)}">Clear ticks for this meal</button>`;
  view().innerHTML = h;
  fitJumpBar();
  refreshTicks();
}
// Plan card folds once every dish has been started; Prep list folds once everything on it is ticked.
// (Only when that happens, so you can open them again.) Ingredient lists fold by hand and stay folded.
const openIngs = new Set(), lastFoldDone = {}; // ingredient lists start folded (the steps show each amount); remember the ones opened
// "THE PLAN  (clock) About 1 hour", then whatever the Plan total row says beyond the time ("all in one skillet"), if anything
function planHeadHtml(meal) {
  const m = String(meal.planTotal || '').trim().match(/^([^.,;]*)[.,;]?\s*(.*)$/);
  const time = (m && m[1].trim()) || (meal.totalMin != null ? `About ${fmtDuration(meal.totalMin)}` : '');
  let extra = m ? m[2].trim() : ''; if (extra) extra = extra[0].toUpperCase() + extra.slice(1);
  return `<div class="planhead"><p class="section-label">The plan</p>${time ? `<span class="plantime">${ICONS.timer} ${esc(time)}</span>` : ''}</div>${extra ? `<p class="planextra">${esc(extra)}</p>` : ''}`;
}
function planStarted(meal) { const t = state.ticks[meal.id] || {}; return meal.dishes.filter((d) => d.steps.length).every((d) => d.steps.some((x) => t[x.key])); }
function autoFold() {
  document.querySelectorAll('details[data-autofold]').forEach((el) => {
    const meal = mealById(el.dataset.meal); if (!meal) return;
    const done = el.dataset.autofold === 'plan' ? planStarted(meal) : [...el.querySelectorAll('.item')].every((x) => x.classList.contains('done'));
    if (done && lastFoldDone[el.id] === false) el.open = false;
    lastFoldDone[el.id] = done;
  });
  // each dish card folds once all its steps are ticked (and starts folded when they already are)
  document.querySelectorAll('.card.dish[data-dishkeys]').forEach((el) => {
    const t = state.ticks[el.dataset.meal] || {}, keys = el.dataset.dishkeys.split(',').filter(Boolean);
    const done = keys.length > 0 && keys.every((k) => t[k]);
    if (done && lastFoldDone[el.id] !== true) setDishFold(el, true);
    lastFoldDone[el.id] = done;
  });
}
const foldedDishes = new Set(); // dish cards folded by hand or when finished; kept while the app is open
function setDishFold(el, fold) { el.classList.toggle('folded', fold); if (fold) foldedDishes.add(el.id); else foldedDishes.delete(el.id); }
document.addEventListener('toggle', (e) => { const el = e.target; if (el.matches && el.matches('details.ingfold')) { if (el.open) openIngs.add(el.dataset.ings); else openIngs.delete(el.dataset.ings); } }, true);
// "STEPS  2/37 ▬▬▬" on one line: how many of these steps are ticked (kept up to date by refreshTicks)
const stepsHeadHtml = (meal, keys) => `<div class="stepshead"><h3>Steps</h3><span class="stepprog" data-stepscope="${esc(meal.id)}" data-stepkeys="${esc(keys.join(','))}"><span class="sp-n"></span><span class="bar"><b></b></span></span></div>`;
// One step of a recipe. In the combined list each step is tagged with its dish and numbered as in that dish.
function stepItemHtml(meal, d, s, prepOf, tagged) {
  const mins = stepMinutes(s.text);
  return itemHtml(meal.id, s.key,
    `<span class="body">${tagged ? `<span class="stepdish chip-${d.color}">${esc(cookTabName(meal, d))} ${esc(s.num)}</span>` : `<span class="num">${esc(s.num)}.</span>`}${prepOf.has(s.key) ? `<a href="#" class="toprep toprep-ic" data-jump="${esc(prepOf.get(s.key))}" aria-label="In the prep list" title="In the prep list">↑</a>` : ''}${esc(s.text)}${s.amount ? `<br><span class="used">${esc(s.amount)}</span>` : ''}</span>${mins ? `<button class="timerbtn tb-side" data-timer="${mins}" data-step="${esc(s.key)}" data-label="${esc(timerLabel(d, s))}" data-meal="${esc(meal.id)}" aria-label="Start a ${mins} minute timer">${ICONS.timer}<span>${mins}m</span></button>` : ''}`,
    mins ? 'hastimer' : '', `${d.anchor}-s${s.num}`);
}
function fitJumpBar() { // keep every top button on screen: shrink the text a little (down to 12.5px) if they don't fit
  const bar = document.querySelector('.jump'); if (!bar) return;
  bar.style.removeProperty('--jump-font');
  for (let f = 14.5; f >= 12.5 && bar.scrollWidth > bar.clientWidth + 1; f -= 0.5) bar.style.setProperty('--jump-font', f + 'px');
}
addEventListener('resize', fitJumpBar);
function shortName(d) { // short label for the jump bar
  const n = d.name.replace(/\s+with\s+.*$/i, '');
  return n.length > 26 ? n.slice(0, 24) + '…' : n;
}

function renderCalendar() {
  setTitle('Calendar', false); setTab('calendar');
  const start = weekStart(state.calWeek), end = new Date(start); end.setDate(end.getDate() + 6);
  const today = dayKey(new Date());
  const label = state.calWeek === 0 ? 'This week' : state.calWeek === 1 ? 'Next week' : state.calWeek === -1 ? 'Last week' : '';
  let h = banners();
  if (!ls.get('cal-hint-seen', false)) h += `<div class="card hintcard"><p>Tap a day to choose or change its meal. Press and hold a meal to drag it to another day.</p><button class="btn ghost" data-hintok="cal-hint-seen">Got it</button></div>`;
  h += `<div class="weeknav">
    <button class="btn ghost" data-week="-1" aria-label="Previous week">‹</button>
    <div><strong>${esc(fmtDay(start, { month: 'short', day: 'numeric' }))} – ${esc(fmtDay(end, { month: 'short', day: 'numeric' }))}</strong>${label ? `<span class="small">${label}</span>` : '<button class="linkbtn" data-week="0">Back to this week</button>'}</div>
    <button class="btn ghost" data-week="1" aria-label="Next week">›</button></div>`;
  h += `<ul class="days${weekSlide ? ' slide-' + weekSlide : ''}">`; weekSlide = '';
  for (let i = 0; i < 7; i++) {
    const d = new Date(start); d.setDate(d.getDate() + i);
    const k = dayKey(d), custom = customText(state.plan[k]), name = custom ? '' : state.plan[k], meal = name ? mealByName(name) : null;
    const past = k < today;
    const extras = SLOTS.filter(([sl]) => sl !== 'dinner' && slotValue(k, sl)).map(([sl, label]) => { const v = slotValue(k, sl), mm = mealByName(v);
      return `<span class="xslot" role="button" tabindex="0" data-slotpick="${k}|${sl}"><b>${label}</b> ${esc(customText(v) || (mm ? mm.main.name.replace(/\s+with\s+.*$/i, '') : 'No longer in the recipes'))}</span>`; }).join('');
    h += `<li class="day${k === today ? ' today' : ''}${past ? ' past' : ''}${state.plan[k] || extras ? '' : ' noplan'}">
      <div class="date"><span>${esc(fmtDay(d, { weekday: 'short' }))}</span><b>${d.getDate()}</b>${state.plan[k] || extras ? `<button class="addslot" data-addslot="${k}" aria-label="Add breakfast, lunch or dessert">＋</button>` : ''}</div>
      <button class="pick" data-day="${k}">${custom
        ? `<strong>${esc(custom)}</strong>`
        : name
        ? (() => {
          const ri = ratingInfo(name), rec = ri.record;
          const main = meal ? meal.main.name : rec && rec.main ? rec.main : displayMealName(name);
          const sides = meal ? meal.sides.map((x) => x.name) : rec && rec.sides ? rec.sides : [];
          return `${meal && meal.kind !== 'Main' ? kindBadge(meal) : ''}<strong>${esc(main)}</strong>${sides.length ? `<span class="small">with ${sides.map(esc).join(' &amp; ')}</span>` : ''}${ri.done ? `<span class="avg">${starsText(ri.avg)}</span>` : ''}${meal ? '' : '<span class="small">No longer in the recipes</span>'}`;
        })()
        : `<span class="addmeal">${k === today ? '+ Plan tonight' : '+ Add meal'}</span>`}${extras ? `<span class="xslots">${extras}</span>` : ''}</button>
      ${meal && k <= today ? (() => { const ri = ratingInfo(meal.name); return `<button class="ratebtn${ri.mine ? ' voted' : ''}" data-rateday="${k}" aria-label="Rate this meal">${ri.mine ? `★ ${ri.mine}` : '☆ Rate'}</button>`; })() : ''}
      ${state.plan[k] || extras ? '' : `<button class="addslot" data-addslot="${k}" aria-label="Add breakfast, lunch or dessert">＋</button>`}
      ${meal ? `<a class="open" href="#/meal/${meal.id}" aria-label="Open recipe">›</a>` : ''}
    </li>`;
  }
  h += '</ul>';
  view().innerHTML = h;
}
// A plan entry is either a recipe meal name (string) or a custom note ({ text: 'Chipotle + Leftovers' }).
const customText = (v) => (v && typeof v === 'object' && v.text ? String(v.text) : '');
// A meal can sit on only one upcoming day (today or later). Past days don't count, so meals can be repeated.
function plannedElsewhere(mealName, day) {
  const today = dayKey(new Date());
  const key = planKey(mealName);
  return planList().filter(([k, v]) => k !== day && k >= today && planKey(v) === key).map(([k]) => k).sort()[0] || null;
}
async function rateDialog(meal) {
  const ri = ratingInfo(meal.name);
  let dlg = $('#ratedlg');
  if (!dlg) { dlg = document.createElement('dialog'); dlg.id = 'ratedlg'; document.body.appendChild(dlg); }
  const status = ri.done ? `<strong class="avg">${starsText(ri.avg)}</strong> average of ${ri.count} votes`
    : ri.count && !ri.mine ? 'The other vote is in. Add yours to see the average.'
    : ri.mine ? 'Waiting for the other vote.' : 'The average shows once you have both voted.';
  dlg.innerHTML = `<form method="dialog"><p class="picktitle">Rate ${esc(meal.main.name)}</p>
    <div class="stars-input">${[1, 2, 3, 4, 5].map((n) => `<button class="star${n <= ri.mine ? ' on' : ''}" value="${n}" aria-label="${n} star${n > 1 ? 's' : ''}">★</button>`).join('')}</div>
    <p class="small">${status}</p>
    <div class="row-btns">${ri.mine ? '<button class="btn outline-danger" value="clear" style="width:auto">Remove my vote</button>' : ''}<button class="btn ghost" value="">Close</button></div></form>`;
  dlg.returnValue = '';
  const v = await new Promise((resolve) => { dlg.addEventListener('close', () => resolve(dlg.returnValue), { once: true }); dlg.showModal(); });
  if (v === 'clear') setRating(meal, null);
  else if (/^[1-5]$/.test(v)) { setRating(meal, Number(v)); toast(`You gave ${v} star${v > 1 ? 's' : ''}`); }
}
async function pickSlot(day) { // the + on a calendar day: which meal?
  const r = await formDialog(`<p class="picktitle">${esc(fmtDay(keyToDate(day), { weekday: 'long', month: 'long', day: 'numeric' }))}</p>
    <div class="picklist">${SLOTS.map(([sl, label]) => { const v = slotValue(day, sl), mm = mealByName(v);
      return `<button class="pickopt" value="${sl}"><strong>${label}</strong>${v ? `<span class="small">${esc(customText(v) || (mm ? mm.main.name : ''))}</span>` : ''}</button>`; }).join('')}</div>
    <div class="row-btns"><button class="btn ghost" value="">Cancel</button></div>`);
  if (r.value) pickMeal(day, r.value);
}
async function pickMeal(day, slot = 'dinner') {
  const m = state.model; if (!m) return;
  const current = slotValue(day, slot), currentCustom = customText(current);
  let dlg = $('#picker');
  if (!dlg) { dlg = document.createElement('dialog'); dlg.id = 'picker'; document.body.appendChild(dlg); }
  const sgIds = new Set(slot === 'dinner' ? suggestions(day).map((x) => x.rid) : []), many = m.meals.length + (m.past || []).length > 8;
  const pastSorted = [...(m.past || [])].sort((a, b) => (slot === 'dessert' ? (a.kind !== 'Dessert') - (b.kind !== 'Dessert') : (a.kind === 'Dessert') - (b.kind === 'Dessert'))); // desserts first for dessert, last otherwise
  dlg.innerHTML = `<form method="dialog"><p class="picktitle" tabindex="-1" autofocus>${esc(fmtDay(keyToDate(day), { weekday: 'long', month: 'long', day: 'numeric' }))}${slot === 'dinner' ? '' : ` · ${slotName(slot)}`}</p>
    ${many ? '<input type="search" class="dlg-input pksearch" placeholder="Search meals" autocomplete="off" enterkeyhint="search">' : ''}
    <div class="pk-body">${(() => { const sg = slot === 'dinner' ? suggestions(day) : []; return sg.length ? `<p class="section-label pastlabel" style="margin-top:0">Suggested</p><div class="picklist suggest">${sg.map((x) => `<button class="pickopt" value="${esc(x.rid)}">${x.archived ? kindBadge(x) : ''}<strong>${esc(x.main.name)}</strong><span class="small">${esc(lastMadeText(x))}${ratingInfo(x.name).done ? ` · ${starsText(ratingInfo(x.name).avg)}` : ''}</span></button>`).join('')}</div><p class="section-label pastlabel">All meals</p>` : ''; })()}
    <div class="picklist">${[...m.meals].filter((x) => !sgIds.has(x.rid)).sort((a, b) => !!plannedElsewhere(a.rid, day) - !!plannedElsewhere(b.rid, day) || ratingInfo(a.name).done - ratingInfo(b.name).done).map((x) => {
      const other = plannedElsewhere(x.rid, day);
      return `<button class="pickopt${x.rid === planKey(current) ? ' sel' : ''}" value="${esc(x.rid)}"${other ? ' disabled' : ''}><strong>${esc(x.main.name)}</strong>${x.sides.length ? `<span class="small">with ${x.sides.map((s) => esc(s.name)).join(' &amp; ')}</span>` : ''}${ratingInfo(x.name).done ? `<span class="small avg">${starsText(ratingInfo(x.name).avg)}</span>` : ''}${other ? `<span class="small taken">Already planned for ${esc(fmtDay(keyToDate(other)))}</span>` : ''}</button>`;
    }).join('')}${pastSorted.filter((x) => !sgIds.has(x.rid)).length ? `<p class="section-label pastlabel">Recipe Book</p>${pastSorted.filter((x) => !sgIds.has(x.rid)).sort((a, b) => !!plannedElsewhere(a.rid, day) - !!plannedElsewhere(b.rid, day)).map((x) => {
      const other = plannedElsewhere(x.rid, day);
      return `<button class="pickopt past${x.rid === planKey(current) ? ' sel' : ''}" value="${esc(x.rid)}"${other ? ' disabled' : ''}>${kindBadge(x)}<strong>${esc(x.main.name)}</strong>${x.sides.length ? `<span class="small">with ${x.sides.map((s) => esc(s.name)).join(' &amp; ')}</span>` : ''}${ratingInfo(x.name).done ? `<span class="small avg">${starsText(ratingInfo(x.name).avg)}</span>` : ''}${other ? `<span class="small taken">Already planned for ${esc(fmtDay(keyToDate(other)))}</span>` : ''}</button>`;
    }).join('')}` : ''}</div></div>
    <div class="custom">
      <label class="small" for="customtext">Or type your own</label>
      <div class="customrow"><input type="text" id="customtext" maxlength="80" placeholder="e.g. Chipotle + Leftovers" value="${esc(currentCustom)}" autocomplete="off"><button class="btn" value="__custom" id="customsave">Save</button></div>
    </div>
    <div class="row-btns">${current ? '<button class="btn outline-danger" value="__none" style="width:auto">Remove</button>' : ''}<button class="btn ghost" value="">Cancel</button></div></form>`;
  dlg.returnValue = '';
  const input = $('#customtext', dlg), srch = $('.pksearch', dlg);
  if (srch) { // filter the list as you type; Enter mustn't pick the first meal
    srch.addEventListener('keydown', (ev) => { if (ev.key === 'Enter') ev.preventDefault(); });
    srch.addEventListener('input', () => { const q = srch.value.trim().toLowerCase(); dlg.querySelectorAll('.picklist .pickopt').forEach((b) => { b.hidden = !!q && !b.textContent.toLowerCase().includes(q); }); dlg.querySelectorAll('.pastlabel').forEach((l) => { l.hidden = !!q; }); });
  }
  input.addEventListener('keydown', (ev) => { if (ev.key === 'Enter') { ev.preventDefault(); $('#customsave', dlg).click(); } });
  const choice = await new Promise((resolve) => { dlg.addEventListener('close', () => resolve(dlg.returnValue), { once: true }); dlg.showModal(); });
  if (choice === '__custom') {
    const text = input.value.replace(/\s+/g, ' ').trim().slice(0, 80);
    if (text && text !== currentCustom) setPlan(day, { text }, slot);
    return;
  }
  if (choice === '__none') setPlan(day, null, slot);
  else if (choice && choice !== planKey(current) && !plannedElsewhere(choice, day)) setPlan(day, choice, slot);
}

function renderShopping() {
  setTitle('Shopping list', false); setTab('shopping');
  const m = state.model;
  if (!m) { view().innerHTML = '<p class="loading">Loading…</p>'; return; }
  let h = banners();
  h += `<details class="card notes" id="notescard"${Object.keys(state.notes).length ? '' : ' hidden'}${notesOpen ? ' open' : ''}><summary id="notesummary">${notesSummaryHtml()}</summary>
    <ul id="notelist" class="list notelist">${notesListHtml()}</ul></details>`;
  h += `<div class="listhead">${progressHtml(SHOP(), m.shopKeys)}</div>
    <div class="shoptools"><label class="switch hidetoggle"><input type="checkbox" id="hideticked"${ls.get('hide-ticked', false) ? ' checked' : ''}><span>Hide ticked</span></label><button class="btn ghost addinline" id="addinline">＋ Add item</button></div>
    <p class="donemsg" id="donemsg" hidden></p>`;
  if (autoShop()) {
    h += mealChipsHtml();
    h += renderAutoShopping(m);
    h += `<div class="shopfoot"><p class="small staplehint">Tip: press and hold items you usually have at home</p>
      <div class="shopbtns"><button class="btn outline-danger" data-clearauto="1">Clear all ticks</button><button class="btn" id="markall">Mark all bought</button></div></div>`;
    view().innerHTML = h;
    fitShopPage();
    refreshTicks();
    return;
  }
  // "Other" goes first. Its own card holds items added in the app as Other, unless the spreadsheet also has an Other section.
  const other = m.shopping.some((sec) => sec.name === 'Other') ? '' : extrasListHtml('Other');
  if (other) h += `<section class="card"><p class="section-label">Other</p><ul class="list">${other}</ul></section>`;
  const shopRow = (x) => itemHtml(SHOP(), x.key,
    `<span class="body"><strong>${esc(x.item)}</strong>${x.amount ? ` · ${esc(x.amount)}` : ''}${x.usedIn ? `<span class="sub">For: ${esc(x.usedIn)}</span>` : ''}</span>`, '', '', x.item);
  for (const sec of m.shopping) {
    const rows = sec.items.filter((x) => !isStaple(x.item)).map(shopRow).join('') + extrasListHtml(sec.name);
    if (rows) h += `<section class="card"><p class="section-label">${esc(sec.name)}</p><ul class="list">${rows}</ul></section>`;
  }
  const staples = m.shopping.flatMap((sec) => sec.items).filter((x) => isStaple(x.item));
  const stapleExtras = extrasListHtml(null, true);
  if (staples.length || stapleExtras) h += `<details class="card staples"${staplesOpen ? ' open' : ''}><summary class="section-label">Usually have (${staples.length + (stapleExtras.match(/data-key=/g) || []).length})</summary>
    <p class="small">Things you usually have at home. They don't count toward the list. Tick one if you're buying it this time.</p><ul class="list">${staples.map(shopRow).join('')}${stapleExtras}</ul></details>`;
  h += '<p class="small staplehint">Tip: press and hold items you usually have at home</p>';
  if (!m.shopping.length && !other) h += '<p class="empty">The shopping list is empty. Tap + to add an item.</p>';
  if (m.shopping.length || other) h += `<button class="btn outline-danger" data-clear="${SHOP()}" data-label="the shopping list">Clear shopping ticks</button>`;
  view().innerHTML = h;
  refreshTicks();
}

function renderSettings() {
  setTitle('Settings', false); setTab('settings');
  const info = state.recipeInfo || {};
  let h = banners();
  // one small card for the two everyday settings, so the whole page fits on a phone screen
  h += `<section class="card settop"><label class="switch"><span>Keep screen on while cooking</span><input type="checkbox" id="wake" ${ls.get('wake', true) ? 'checked' : ''}></label>
    <div class="updrow"><button class="btn ghost" type="button" id="update">Check for updates</button><span class="ver">${esc(window.APP_VERSION || 'dev')}</span></div>
    ${info.error ? `<p class="err">${esc(info.error)}</p>` : ''}</section>`;
  const cur = ls.get('color-theme', 'teal');
  h += `<section class="card"><p class="section-label">Color</p><div class="swatches">${THEMES.map(([id, name, sw]) => `<button class="swatch${id === cur ? ' on' : ''}" data-colortheme="${id}" aria-pressed="${id === cur}"><i style="background:${sw}"></i>${name}</button>`).join('')}</div></section>`;
  // Account: rarely needed, so folded; the summary shows who you are
  h += `<details class="card account"><summary class="section-label">Account<span class="acct-who">${esc(state.myName || (state.user && state.user.email) || '')}</span></summary>
    ${!state.firebase ? '<p>Not set up: everything stays on this phone.</p>' : state.user ? `<div class="signedin"><span>${esc(state.user.email)}</span><button class="btn ghost" id="signout">Sign out</button></div>` : ''}
    <form id="nameform"><label class="field">Your name <span class="small">(on your notes, items and ideas; saves by itself)</span>
    <input type="text" id="myname" class="dlg-input" maxlength="40" placeholder="${esc((state.user && state.user.email) || 'Your name')}" value="${esc(state.myName)}" autocomplete="off" enterkeyhint="done"></label></form></details>`;
  if (state.firebase && state.user) {
    const row = (what, n) => { const x = syncInfo[what]; return `<li>${esc(what)}: ${x ? (x.ok ? `${n} ${x.server ? '✓' : '(from this phone, not yet confirmed)'}` : `<span class="err">error: ${esc(x.error)}</span>`) : 'not loaded yet'}</li>`; };
    h += `<details class="card synccheck"><summary class="section-label">Diagnostics</summary><ul class="small">
      ${row('notes', Object.keys(state.notes).length + ' notes')}${row('added items', Object.keys(state.extras).length + ' items')}
      ${row('calendar', Object.keys(state.plan).length + ' days')}${row('ratings', Object.keys(state.ratings).length + ' rated')}
      ${row('names', Object.keys(state.names).length + ' names')}${row('cooking', Object.keys(state.session.timers || {}).length + ' timers')}${row('staples', Object.keys(state.staples).length + ' staples')}${row('bought', Object.keys(state.bought).length + ' meals')}${row('ideas', Object.keys(state.ideas).length + ' ideas')}${row('times made', Object.keys(state.madeAdjust).length + ' changed')}${row('next-time notes', Object.keys(state.nextTime).length + ' notes')}<li>shopping list: ${autoShop() ? 'automatic, from the recipes' : 'from the Shopping tab (not every dish has shopping details yet)'}</li></ul>
      <p class="small">If this phone is missing things the other phone shows, reload everything from Firebase. Close any other copy of the app (e.g. a Chrome tab on the same address) first.</p>
      <button class="btn ghost" type="button" id="resetcache">Reload all data from Firebase</button></details>`;
  }
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
  setTimeout(autoFold, 0);
  if (cook) markWheel(cook.step);
  document.querySelectorAll('.item.preptask').forEach((el) => { // a prep task is done when all its steps are
    const t = state.ticks[el.dataset.scope] || {};
    const on = el.dataset.prepkeys.split(',').every((k) => t[k]);
    el.classList.toggle('done', on); el.setAttribute('aria-checked', on);
  });
  // shopping list only: ticked items move to the bottom of their section, unticked ones keep their order
  document.querySelectorAll('ul.list').forEach((ul) => {
    const items = [...ul.children].filter((li) => li.dataset.key);
    if (!items.length || !isShopScope(items[0].dataset.scope)) return;
    items.forEach((li, i) => { if (li.dataset.i == null) li.dataset.i = i; });
    const want = [...items].sort((a, b) => a.classList.contains('done') - b.classList.contains('done') || a.dataset.i - b.dataset.i);
    if (want.some((li, i) => li !== items[i])) want.forEach((li) => ul.appendChild(li));
  });
  document.querySelectorAll('[data-stepkeys]').forEach((el) => { // per-list step counts under "Steps"
    const t = state.ticks[el.dataset.stepscope] || {}, keys = el.dataset.stepkeys.split(',').filter(Boolean), done = keys.filter((k) => t[k]).length;
    el.querySelector('.sp-n').textContent = `${done}/${keys.length}`; el.querySelector('b').style.width = keys.length ? (100 * done / keys.length) + '%' : 0;
    el.classList.toggle('complete', keys.length > 0 && done === keys.length);
  });
  document.querySelectorAll('[data-progress]').forEach((el) => {
    const scope = el.dataset.progress, m = state.model; if (!m) return;
    let keys = isShopScope(scope) ? shopCountKeys() : (mealById(scope) || {}).keys || [];
    let done = countDone(scope, keys), total = keys.length;
    if (isShopScope(scope) && autoShop()) [done, total] = autoCounts();
    el.querySelector('b').style.width = total ? (100 * done / total) + '%' : 0;
    el.querySelector('.pct').textContent = isShopScope(scope) ? (total - done ? `${total - done} to buy` : 'All bought') : `${done} of ${total} steps`;
    if (isShopScope(scope)) showDoneMessage(total > 0 && done === total);
  });
  // shopping list: optionally hide ticked items (this phone only), and any card left with nothing showing
  const hide = !!$('#hideticked') && ls.get('hide-ticked', false);
  document.body.classList.toggle('hide-ticked', hide);
  if ($('#hideticked')) document.querySelectorAll('main section.card').forEach((card) => {
    const items = card.querySelectorAll('.item[data-key], .item.gline');
    card.classList.toggle('all-ticked', items.length > 0 && [...items].every((x) => x.classList.contains('done')));
  });
}

// Shown under the shopping counter when every item (not notes) is ticked. A new one is picked each time.
// shopping items that count toward the list (staples don't)
function shopCountKeys() {
  const m = state.model; if (!m) return [];
  return [...m.shopping.flatMap((sec) => sec.items).filter((x) => !isStaple(x.item)).map((x) => x.key),
    ...Object.entries(state.extras).filter(([, x]) => !isStaple(x.name)).map(([id]) => extraKey(id))];
}
const DONE_MESSAGES = [
  "It's in the bag 🛍️", 'Groceries: handled 🥕', "Shopping done. Dinner's on its way 🍽️", 'Nailed it! 🎉', 'Every last thing. Nice work 🙌',
  "That's a wrap 🛒", 'Fridge, prepare yourself 😄', 'List conquered 💪', 'All set. Time to cook 🧑‍🍳', 'Look at you go ✨',
  'Cart: full. Mission: complete 🚀', 'Pantry restocked like a pro 🏆', 'Shopping champion of the week 🥇', 'The fridge thanks you 🧊',
  'Done and dusted 🧹', 'Groceries secured. Snacks earned 🍪', 'Flawless grocery run 👌', 'Smooth shopping, smooth sailing ⛵',
  'Everything ticked. High five ✋', 'Bags packed, dinner sorted 🥡',
];
// keep the same message while you're on the screen (ticks redraw the list); pick a new one on each visit
let doneMsg = '';
function showDoneMessage(complete) {
  const el = $('#donemsg'); if (!el) return;
  if (!complete) { doneMsg = ''; el.hidden = true; el.textContent = ''; return; }
  if (!doneMsg) doneMsg = DONE_MESSAGES[Math.floor(Math.random() * DONE_MESSAGES.length)];
  if (el.textContent !== doneMsg) el.textContent = doneMsg;
  el.hidden = false;
}

// ---------------------------------------------------------------- events
document.addEventListener('click', async (e) => {
  const jump = e.target.closest('[data-jump]');
  if (jump) {
    e.preventDefault();
    if (jump.classList.contains('jump-all') && !document.getElementById(jump.dataset.jump)) { // in By dish view: switch to Combined, then go there
      ls.set('combined-view', true); const y = scrollY; route(); scrollTo(0, y);
      const again = document.querySelector('.jump-all'); if (again) setTimeout(() => again.click(), 30);
      return;
    }
    const target = document.getElementById(jump.dataset.jump) || document.getElementById(jump.dataset.jump.replace(/-s\d+$/, ''));
    if (target) {
      for (let f = target.closest('details'); f; f = f.parentElement && f.parentElement.closest('details')) f.open = true; // open any folded card it's in
      { const dc = target.closest('.card.dish.folded'); if (dc) setDishFold(dc, false); }
      // a dish column that scrolls on its own (wide screens): scroll inside it; otherwise scroll the page
      // All, and the first dish in By dish view: bring the By dish / Combined switch into view too
      const seg = document.querySelector('.viewseg');
      const toSeg = seg && (jump.classList.contains('jump-all') || (!ls.get('combined-view', true) && jump === jump.parentElement.querySelector('[class*="chip-"]')));
      const col = target.closest('.dishes .card.dish');
      if (toSeg) {
        scrollTo({ top: seg.getBoundingClientRect().top + scrollY - 118, behavior: 'smooth' });
        if (col) col.scrollTo({ top: 0, behavior: 'smooth' });
      } else if (col && col.scrollHeight > col.clientHeight + 2) {
        col.scrollTo({ top: col.scrollTop + target.getBoundingClientRect().top - col.getBoundingClientRect().top - col.clientHeight / 3, behavior: 'smooth' });
        const ct = col.getBoundingClientRect().top; if (ct < 56 || ct > innerHeight / 2) scrollTo({ top: ct + scrollY - 130, behavior: 'smooth' });
      } else {
        scrollTo({ top: target.getBoundingClientRect().top + scrollY - 130, behavior: 'smooth' });
      }
      if (target.classList.contains('item')) { target.classList.remove('flash'); void target.offsetWidth; target.classList.add('flash'); }
    }
    return;
  }
  const dh = e.target.closest('.card.dish > .dish-head.foldable');
  if (dh && !e.target.closest('button, a')) { const card = dh.parentElement; setDishFold(card, !card.classList.contains('folded')); return; }
  const wk = e.target.closest('[data-week]');
  if (wk) { state.calWeek = wk.dataset.week === '0' ? 0 : state.calWeek + Number(wk.dataset.week); route(); return; }
  const star = e.target.closest('[data-rate]');
  if (star) {
    const meal = mealById(star.dataset.meal);
    if (meal) { const n = Number(star.dataset.rate); setRating(meal, ratingInfo(meal.name).mine === n ? null : n); }
    return;
  }
  if (e.target.closest('#addbtn, #addinline')) { addMenu(); return; }
  if (held) { held = false; e.preventDefault(); return; } // a press-and-hold just opened a menu (don't also open the link)
  const tb = e.target.closest('[data-timer]');
  if (tb) { timerTap(tb.dataset.label, +tb.dataset.timer, tb.dataset.meal, tb.dataset.step); return; }
  const stp = e.target.closest('[data-stoptimer]');
  if (stp) { writeSession('timers', stp.dataset.stoptimer, null); return; }
  const oc = e.target.closest('[data-opencook]');
  if (oc) { openCook(oc.dataset.opencook, +(oc.dataset.dish || 0)); return; }
  const cw = e.target.closest('.cw-step:not(.cw-d0)');
  if (cw) { wheelTo(+cw.dataset.cwi, true); return; }
  const ca = e.target.closest('[data-cook]');
  if (ca) { cookAction(ca.dataset.cook); return; }
  if (e.target.id === 'resetcache') { resetLocalData(); return; }
  const qr = e.target.closest('[data-quickrate]');
  if (qr) { const meal = mealById(qr.dataset.meal); if (meal) { const n = +qr.dataset.quickrate; setRating(meal, n); toast(`You gave ${n} star${n > 1 ? 's' : ''}`); } return; }
  const rd = e.target.closest('[data-ratedismiss]');
  if (rd) { const d = ls.get('rate-dismissed', {}); d[rd.dataset.ratedismiss] = 1; ls.set('rate-dismissed', d); route(); return; }
  const tc = e.target.closest('[data-tonightcook]');
  if (tc) { e.preventDefault(); const id = tc.dataset.tonightcook, tm = mealById(id), i = tm && tm.combined && ls.get('combined-view', true) ? -1 : +tc.dataset.dish; addEventListener('hashchange', () => openCook(id, i), { once: true }); location.hash = '#/meal/' + id; return; }
  const pd = e.target.closest('[data-pickday]');
  if (pd) { e.preventDefault(); const meal = mealById(pd.dataset.pickday); if (meal) pickDayFor(meal); return; }
  const hk = e.target.closest('[data-hintok]');
  if (hk) { ls.set(hk.dataset.hintok, true); route(); return; }
  const rp = e.target.closest('[data-replacephoto]');
  if (rp) { pickPhoto(rp.dataset.replacephoto); return; }
  const xp = e.target.closest('[data-removephoto]');
  if (xp) { removePhoto(xp.dataset.removephoto); return; }
  const bm = e.target.closest('[data-bookmenu]');
  if (bm) { const meal = mealById(bm.dataset.meal); if (meal) madeMenu(meal, meal.dishes[+bm.dataset.dish]); return; }
  const ap = e.target.closest('[data-addphoto]');
  if (ap) { pickPhoto(ap.dataset.addphoto); return; }
  const ph = e.target.closest('.dishphoto[data-photo]');
  if (ph) { photoMenu(ph.dataset.photo, ph.dataset.name); return; }
  const ct = e.target.closest('[data-colortheme]');
  if (ct) { ls.set('color-theme', ct.dataset.colortheme); applyColor(); route(); return; }
  if (e.target.id === 'copyideas') { copyIdeas(); return; }
  const enx = e.target.closest('[data-editnext]');
  if (enx) { const meal = mealById(enx.dataset.editnext); if (meal) editNextTime(meal); return; }
  if (e.target.id === 'addidea') { addIdea(); return; }
  if (e.target.id === 'unsendall') { // every sent idea back to the Ideas list (anyone's), with Undo
    const n = ideaList(true).filter(([, x]) => !ideaRecipe(x)); if (!n.length) return;
    writeIdeas(Object.fromEntries(n.map(([id, x]) => [id, { ...x, sent: null }])));
    undoToast(`${n.length} idea${n.length === 1 ? '' : 's'} moved back`, () => writeIdeas(Object.fromEntries(n))); return; }
  if (e.target.id === 'clearsent') { const n = ideaList(true).filter(([, x]) => !ideaRecipe(x)); if (n.length) { writeIdeas(Object.fromEntries(n.map(([id]) => [id, null]))); undoToast(`${n.length} sent idea${n.length === 1 ? '' : 's'} cleared`, () => writeIdeas(Object.fromEntries(n))); } return; }
  const ei = e.target.closest('[data-editidea]');
  if (ei) { editIdea(ei.dataset.editidea); return; }
  const di = e.target.closest('[data-delidea]');
  if (di) { const id = di.dataset.delidea, x = state.ideas[id]; if (x) { writeIdeas({ [id]: null }); undoToast('Idea deleted', () => writeIdeas({ [id]: x })); } return; }
  const en = e.target.closest('[data-editnote]');
  if (en) { editNote(en.dataset.editnote); return; }
  const ex = e.target.closest('[data-editextra]');
  if (ex) { editItem(ex.dataset.editextra); return; }
  const dx = e.target.closest('[data-delextra]');
  if (dx) { const id = dx.dataset.delextra, x = state.extras[id]; if (x) { writeExtras({ [id]: null }); undoToast(`${x.name} removed`, () => writeExtras({ [id]: x })); } return; }
  const dn = e.target.closest('[data-delnote]');
  if (dn) { const id = dn.dataset.delnote, x = state.notes[id]; if (x) { writeNotes({ [id]: null }); undoToast('Note deleted', () => writeNotes({ [id]: x })); } return; }
  const bsrt = e.target.closest('[data-booksort]');
  if (bsrt) { ls.set('book-sort', bsrt.dataset.booksort); route(); return; }
  const sv = e.target.closest('[data-stepview]');
  if (sv) { e.preventDefault(); ls.set('combined-view', sv.dataset.stepview === 'combined'); route(); return; }
  const pf = e.target.closest('[data-pastfilter]');
  if (pf) { ls.set('past-filter', pf.dataset.pastfilter); route(); return; }
  const rm = e.target.closest('[data-ratemeal]');
  if (rm) { const meal = mealById(rm.dataset.ratemeal); if (meal) rateDialog(meal); return; }
  const rb = e.target.closest('[data-rateday]');
  if (rb) { const meal = mealByName(state.plan[rb.dataset.rateday]); if (meal) rateDialog(meal); return; }
  const sp = e.target.closest('[data-slotpick]');
  if (sp) { if (!suppressClick) { const [d, sl] = sp.dataset.slotpick.split('|'); pickMeal(d, sl); } return; }
  const as = e.target.closest('[data-addslot]');
  if (as) { pickSlot(as.dataset.addslot); return; }
  const pick = e.target.closest('.pick[data-day]');
  if (pick) { if (!suppressClick) pickMeal(pick.dataset.day); return; }
  if (e.target.closest('.delnote')) return;
  const pg = e.target.closest('[data-prepgo]');
  if (pg) { // jump to the step; ask which dish when the task is used by more than one
    const opts = pg.dataset.prepgo.split(';').map((x) => x.split('|'));
    let go = opts[0][0];
    if (opts.length > 1) {
      const r = await formDialog(`<p class="picktitle">Go to which dish?</p><div class="picklist">${opts.map(([id, name]) => `<button class="pickopt" value="${esc(id)}"><strong>${esc(name)}</strong></button>`).join('')}</div><div class="row-btns"><button class="btn ghost" value="">Cancel</button></div>`);
      if (!r.value) return; go = r.value;
    }
    const a = document.createElement('a'); a.dataset.jump = go; a.hidden = true; document.body.appendChild(a); a.click(); a.remove();
    return;
  }
  const pt = e.target.closest('.item.preptask');
  if (pt) { const on = !pt.classList.contains('done'); pt.dataset.prepkeys.split(',').forEach((k) => setTick(pt.dataset.scope, k, on)); return; }
  const gl = e.target.closest('.item.gline');
  if (gl) { toggleLine(gl.dataset.line); return; }
  if (e.target.id === 'markall') { markAllBought(); return; }
  const msh = e.target.closest('[data-mealshopped]');
  if (msh) { toggleMealShopped(msh.dataset.mealshopped); return; }
  const ca2 = e.target.closest('[data-clearauto]');
  if (ca2) { // clear straight away, with Undo
    const beforeB = JSON.parse(JSON.stringify(state.bought)), beforeT = { ...(state.ticks[SHOP()] || {}) };
    const saveBought = () => { if (state.firebase) setDoc(doc(state.firebase.db, 'shopping', 'bought'), { b: state.bought }).catch((err) => toast('Could not save: ' + (err.code || err.message))); else ls.set('local-bought', state.bought); };
    state.bought = {}; saveBought(); clearScope(SHOP()); route();
    undoToast('All ticks cleared', () => { state.bought = beforeB; saveBought(); clearScope(SHOP(), beforeT); route(); });
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
  if (e.target.id === 'update') { // short messages: Updating… → Up to Date! (or Already Up to Date)
    toast('Updating…', 20000);
    const before = state.rowsSig;
    await loadRecipes();
    const recipesFailed = !!(state.recipeInfo && state.recipeInfo.error), recipesChanged = state.rowsSig !== before;
    if (location.hash.startsWith('#/settings')) route();
    try {
      const r = await fetch('sw.js', { cache: 'no-store' });
      const v = ((await r.text()).match(/VERSION = '([^']+)'/) || [])[1];
      if (!v) throw new Error('no version found');
      if (v === window.APP_VERSION) {
        toast(recipesFailed ? 'Couldn’t update. Check your connection.' : recipesChanged ? 'Up to Date!' : 'Already Up to Date', 3000);
        return;
      }
      const reg = await navigator.serviceWorker.getRegistration();
      if (reg) {
        await reg.update();
        const sw = reg.installing || reg.waiting;
        if (sw && sw.state !== 'activated') await new Promise((res) => { sw.addEventListener('statechange', () => sw.state === 'activated' && res()); setTimeout(res, 15000); });
      }
      ls.set('just-updated', true); // the reloaded app says "Up to Date!"
      location.reload();
    } catch { toast('Couldn’t update. Check your connection.', 3500); }
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
  if ((e.key === ' ' || e.key === 'Enter') && e.target.matches('.item[data-key], .item.gline')) { e.preventDefault(); e.target.click(); }
});
document.addEventListener('toggle', (e) => { if (e.target.matches && e.target.matches('details.ideasent')) ideasSentOpen = e.target.open; }, true);
function saveName() { // the Your name box saves itself when you leave it
  const name = $('#myname').value.replace(/\s+/g, ' ').trim().slice(0, 40);
  if (name === state.myName) return;
  state.myName = name;
    if (state.user) { if (name) state.names[myEmail()] = name; else delete state.names[myEmail()]; }
    renderNotes(); renderExtras();
    if (state.firebase && state.user) setDoc(doc(state.firebase.db, 'profiles', state.user.uid), { name, email: state.user.email }, { merge: true }).catch((err) => toast('Could not save: ' + (err.code || err.message)));
    else ls.set('local-name', name);
    toast(name ? `Saved. Your notes and items will show "${name}"` : 'Saved. Your notes and items will show your email');
  const who = $('.acct-who'); if (who) who.textContent = name || (state.user && state.user.email) || '';
}
document.addEventListener('change', (e) => { if (e.target.id === 'myname') saveName(); });
document.addEventListener('submit', async (e) => {
  if (e.target.id === 'nameform') {
    e.preventDefault(); $('#myname').blur(); // saving happens on change (below)
    return;
  }
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
});
document.addEventListener('change', (e) => {
  if (e.target.id === 'wake') { ls.set('wake', e.target.checked); updateWakeLock(); }
  if (e.target.id === 'hideticked') { ls.set('hide-ticked', e.target.checked); refreshTicks(); }
});

// keep screen on while a meal is open
let wakeLock = null;
async function updateWakeLock() {
  const myTimer = Object.values(state.session.timers || {}).some((t) => (t.by || '').toLowerCase() === myEmail() && t.ends > Date.now());
  const want = ls.get('wake', true) && (location.hash.startsWith('#/meal/') || myTimer) && document.visibilityState === 'visible';
  try {
    if (want && !wakeLock && 'wakeLock' in navigator) { wakeLock = await navigator.wakeLock.request('screen'); wakeLock.addEventListener('release', () => { wakeLock = null; }); }
    else if (!want && wakeLock) { await wakeLock.release(); wakeLock = null; }
  } catch { wakeLock = null; }
}
document.addEventListener('visibilitychange', () => { updateWakeLock(); if (document.visibilityState === 'visible') loadRecipes(); });

// ---------------------------------------------------------------- calendar drag and drop
// Press and hold a planned day (~0.35 s), then drag it onto another day. Works with touch and mouse.
let drag = null, suppressClick = false;
// Swipe the calendar left/right to go to the next/previous week (not while dragging a meal).
let weekSlide = '', calSwipe = null;
document.addEventListener('touchstart', (e) => {
  calSwipe = e.touches.length === 1 && /^#\/calendar/.test(location.hash) && !e.target.closest('dialog') ? { x: e.touches[0].clientX, y: e.touches[0].clientY, t: Date.now() } : null;
}, { passive: true });
document.addEventListener('touchend', (e) => {
  if (!calSwipe || (drag && drag.active)) { calSwipe = null; return; }
  const t = e.changedTouches[0], dx = t.clientX - calSwipe.x, dy = t.clientY - calSwipe.y, quick = Date.now() - calSwipe.t < 700; calSwipe = null;
  if (!quick || Math.abs(dx) < 70 || Math.abs(dy) > Math.abs(dx) * 0.6) return;
  state.calWeek += dx < 0 ? 1 : -1; weekSlide = dx < 0 ? 'next' : 'prev';
  suppressClick = true; setTimeout(() => { suppressClick = false; }, 350);
  route();
}, { passive: true });
function dragCleanup() {
  if (!drag) return;
  clearTimeout(drag.timer); clearTimeout(drag.weekTimer);
  if (drag.ghost) drag.ghost.remove();
  if (drag.held) drag.held.remove();
  document.querySelectorAll('.day.drop-target, .day.dragging, [data-week].drop-target').forEach((el) => el.classList.remove('drop-target', 'dragging'));
  document.body.classList.remove('is-dragging');
  drag = null;
}
function dragStart() {
  const src = document.querySelector(`.pick[data-day="${drag.from}"]`);
  if (!src) { dragCleanup(); return; }
  drag.active = true;
  const r = src.getBoundingClientRect();
  const g = src.cloneNode(true); g.className = 'drag-ghost'; g.style.width = r.width + 'px';
  drag.dx = drag.x - r.left; drag.dy = drag.y - r.top;
  document.body.appendChild(g); drag.ghost = g; dragMoveGhost();
  src.closest('.day').classList.add('dragging'); document.body.classList.add('is-dragging');
  // The finger's touch stays tied to this element. Keep a hidden copy of it in the page for the whole drag,
  // so redrawing the calendar (changing week, or an update from the other phone) doesn't drop the drag.
  const holder = document.createElement('div'); holder.className = 'drag-holder';
  const placeholder = src.cloneNode(true); src.replaceWith(placeholder);
  holder.appendChild(src); document.body.appendChild(holder); drag.held = holder;
  if (navigator.vibrate) navigator.vibrate(25);
}
function dragMoveGhost() { drag.ghost.style.transform = `translate(${drag.x - drag.dx}px, ${drag.y - drag.dy}px)`; }
function dragTargetAt(x, y) {
  const el = document.elementFromPoint(x, y);
  return el ? el.closest('.day, .weeknav [data-week]') : null;
}
document.addEventListener('pointerdown', (e) => {
  const pick = e.target.closest('.day .pick');
  if (!pick || !state.plan[pick.dataset.day] || (e.pointerType === 'mouse' && e.button !== 0)) return;
  dragCleanup();
  drag = { from: pick.dataset.day, x: e.clientX, y: e.clientY, x0: e.clientX, y0: e.clientY, active: false };
  drag.timer = setTimeout(dragStart, 350);
});
document.addEventListener('pointermove', (e) => {
  if (!drag) return;
  drag.x = e.clientX; drag.y = e.clientY;
  if (!drag.active) { if (Math.hypot(drag.x - drag.x0, drag.y - drag.y0) > 10) dragCleanup(); return; } // it's a scroll
  dragMoveGhost();
  const t = dragTargetAt(drag.x, drag.y);
  document.querySelectorAll('.drop-target').forEach((el) => el !== t && el.classList.remove('drop-target'));
  if (t && !t.classList.contains('drop-target')) {
    t.classList.add('drop-target');
    clearTimeout(drag.weekTimer);
    if (t.dataset.week) drag.weekTimer = setTimeout(() => { // hover on ‹ › to change week
      drag.paged = true;
      state.calWeek += Number(t.dataset.week); route(); document.body.classList.add('is-dragging');
      const from = document.querySelector(`.day .pick[data-day="${drag.from}"]`); if (from) from.closest('.day').classList.add('dragging');
    }, drag.paged ? 900 : 600);
  } else if (!t) clearTimeout(drag.weekTimer);
  // scroll the page when dragging near the top or bottom edge
  const edge = 90;
  if (drag.y < edge + 60) scrollBy(0, -12); else if (drag.y > innerHeight - edge - 70) scrollBy(0, 12);
});
document.addEventListener('pointerup', () => {
  if (!drag) return;
  if (drag.active) {
    const t = dragTargetAt(drag.x, drag.y);
    const to = t && t.classList.contains('day') ? t.querySelector('.pick').dataset.day : null;
    const from = drag.from;
    suppressClick = true; setTimeout(() => { suppressClick = false; }, 400);
    dragCleanup();
    if (to && to !== from) { movePlan(from, to); toast('Moved'); }
  } else dragCleanup();
});
document.addEventListener('pointercancel', () => { if (drag && !drag.active) dragCleanup(); });
// once a drag has started, stop the page from scrolling under the finger
document.addEventListener('touchmove', (e) => { if (drag && drag.active) e.preventDefault(); }, { passive: false });
document.addEventListener('contextmenu', (e) => { if (e.target.closest('.day .pick')) e.preventDefault(); });

// ---------------------------------------------------------------- router
let lastRoute = '', navCount = 0;
function route() {
  if (state.firebase && !state.authReady) { view().innerHTML = '<p class="loading">Loading…</p>'; return; }
  if (state.firebase && !state.user) { renderLogin(); return; }
  $('#tabs').hidden = false;
  const hash = location.hash.replace(/^#\/?/, '');
  const [page, arg] = hash.split('/');
  const keepScroll = hash === lastRoute;
  if (!keepScroll) doneMsg = ''; // a new visit to a screen gets a new cheer
  const y = scrollY;
  $('#addbtn').hidden = true; // adding is the "+ Add item" button on the list itself
  view().style.minHeight = ''; // only the shopping page stretches to fill the screen
  if (cook && !(page === 'meal' && decodeURIComponent(arg || '') === cook.meal)) { cook = null; renderCook(); }
  if (page === 'meal') renderMeal(decodeURIComponent(arg || ''));
  else if (page === 'shopping') renderShopping();
  else if (page === 'calendar') renderCalendar();
  else if (page === 'past' || page === 'planning') renderPast(); // Recipe Book (Planning tab)
  else if (page === 'ideas') renderPlanning();
  else if (page === 'settings') renderSettings();
  else renderMeals();
  paintPhotos();
  if (keepScroll) scrollTo(0, y); else scrollTo(0, 0);
  if (hash !== lastRoute) navCount++;
  lastRoute = hash;
  updateWakeLock();
}
addEventListener('hashchange', route);

// ---------------------------------------------------------------- start
function start() {
  applyColor();
  if (ls.get('just-updated', false)) { ls.set('just-updated', false); setTimeout(() => toast('Up to Date!', 3000), 300); }
  try { localStorage.removeItem('sheet-url'); localStorage.removeItem('clock-timer'); } catch { /* old settings, no longer used */ }
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
    state.plan = ls.get('local-plan', {});
    state.planX = ls.get('local-planx', {});
    state.ratings = ls.get('local-ratings', {});
    state.notes = ls.get('local-notes', {});
    state.myName = ls.get('local-name', '');
    state.extras = ls.get('local-extras', {});
    state.session = ls.get('local-session', { dinner: {}, timers: {} });
    state.staples = ls.get('local-staples', {});
    state.bought = ls.get('local-bought', {});
    state.ideas = ls.get('local-ideas', {});
    state.nextTime = ls.get('local-nexttime', {});
    state.madeAdjust = ls.get('local-adjust', {});
    state.photos = ls.get('local-photos', {});
    state.thumbs = ls.get('local-thumbs', {}); thumbsLoaded = true;
    state.sync = 'local'; renderSync();
  }
  route();
  loadRecipes();
  if ('serviceWorker' in navigator && location.protocol !== 'file:') {
    navigator.serviceWorker.register('sw.js').catch((e) => console.warn('SW failed', e));
  }
}
start();
