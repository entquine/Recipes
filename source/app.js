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
    meal.kind = meal.main.type; // 'Main' or 'Side' (past entries are one dish each)
    const mins = meal.plan.map((p) => p.min).filter((m) => m != null);
    meal.totalMin = mins.length ? Math.max(...mins) : null;
    meal.keys = dishes.flatMap((d) => d.steps.map((s) => s.key));
    computeStarts(meal);
    out.push(meal);
  }
  out.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }));
  return out;
}
function buildShopping(shopping) {
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
  return { shopping: shop, shopKeys: shop.flatMap((s) => s.items.map((x) => x.key)) };
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
  plan: {},                // 'YYYY-MM-DD' -> meal name (shared calendar)
  ratings: {},             // meal id -> { name, main, sides, votes: { userId: 1-5 } }
  notes: {},               // shopping notes: id -> { text, by, at }
  myName: '',              // this person's name (Settings)
  names: {},               // email (lower-case) -> name, from Firestore profiles/*; used to label notes and added items
  extras: {},              // items added to the shopping list in the app: id -> { name, amount, email, at }
  calWeek: 0,              // calendar page: 0 = this week
};
const CFG = window.APP_CONFIG || {};
const fbConfig = CFG.firebase && CFG.firebase.apiKey && CFG.firebase.projectId ? CFG.firebase : null;

// ---------------------------------------------------------------- ticks store
let unsubTicks = null, unsubPlan = null, unsubRatings = null, unsubNotes = null, unsubProfile = null, unsubExtras = null, lastSnapPending = false;

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
  unsubPlan = onSnapshot(doc(db, 'calendar', 'plan'), (snap) => {
    state.plan = { ...((snap.exists() && snap.data().d) || {}) };
    if (/^#?\/?(calendar|past)?$/.test(location.hash)) route(); // re-draw Calendar or Meals list
  }, () => {});
  unsubRatings = onSnapshot(collection(db, 'ratings'), (snap) => {
    const r = {}; snap.forEach((d) => { r[d.id] = d.data(); });
    state.ratings = r;
    if (/^#?\/?(calendar|past|meal\/.*)?$/.test(location.hash)) route();
  }, () => {});
  unsubNotes = onSnapshot(doc(db, 'notes', 'shopping'), (snap) => {
    state.notes = { ...((snap.exists() && snap.data().n) || {}) };
    renderNotes();
  }, () => {});
  // everyone's names, so notes and added items show the latest name for each email
  unsubProfile = onSnapshot(collection(db, 'profiles'), (snap) => {
    const names = {};
    snap.forEach((d) => {
      const x = d.data();
      if (x.email && x.name) names[String(x.email).toLowerCase()] = x.name;
      if (d.id === state.user.uid) state.myName = x.name || '';
    });
    state.names = names;
    const inp = $('#myname'); if (inp && document.activeElement !== inp) inp.value = state.myName;
    renderNotes(); renderExtras();
  }, () => {});
  unsubExtras = onSnapshot(doc(db, 'shopping', 'extra'), (snap) => {
    state.extras = { ...((snap.exists() && snap.data().i) || {}) };
    renderExtras();
  }, () => {});
}
function stopFirestore() { for (const u of [unsubTicks, unsubPlan, unsubRatings, unsubNotes, unsubProfile, unsubExtras]) if (u) u(); unsubTicks = unsubPlan = unsubRatings = unsubNotes = unsubProfile = unsubExtras = null; }

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
const fmtWhen = (at) => (at ? new Date(at).toLocaleString(undefined, { weekday: 'short', hour: 'numeric', minute: '2-digit' }) : '');
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
  return notes.map(([id, x]) => `<li class="note"><div><span class="notetext">${esc(x.text)}</span>
    <span class="small">${esc(nameFor(x.email, x.by))}${x.at ? ` · ${esc(fmtWhen(x.at))}` : ''}</span></div>
    <button class="delnote" data-delnote="${esc(id)}" aria-label="Delete note">×</button></li>`).join('');
}
function renderNotes() {
  const ul = $('#notelist'); if (!ul) return;
  ul.innerHTML = notesListHtml();
  $('#notescard').hidden = !Object.keys(state.notes).length;
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
function extrasListHtml() {
  const items = Object.entries(state.extras).sort((a, b) => (a[1].at || 0) - (b[1].at || 0));
  if (!items.length) return '';
  return items.map(([id, x]) => itemHtml('shopping', extraKey(id),
    `<span class="body"><strong>${esc(x.name)}</strong>${x.amount ? ` · ${esc(x.amount)}` : ''}<span class="sub">Added by ${esc(nameFor(x.email))}</span></span>${
      (x.email || '').toLowerCase() === myEmail()
        ? `<button class="delnote" data-editextra="${esc(id)}" aria-label="Edit ${esc(x.name)}">✎</button>`
        : `<button class="delnote" data-delextra="${esc(id)}" aria-label="Remove ${esc(x.name)}">×</button>`}`, 'extra')).join('');
}
function renderExtras() {
  const ul = $('#extralist'); if (!ul) return;
  ul.innerHTML = extrasListHtml();
  $('#extrascard').hidden = !Object.keys(state.extras).length;
  refreshTicks();
}

// ---------------------------------------------------------------- "+" in the top bar (shopping list)
function formDialog(html) { // shows a small form dialog; resolves with the submit button's value and the form
  let dlg = $('#formdlg');
  if (!dlg) { dlg = document.createElement('dialog'); dlg.id = 'formdlg'; document.body.appendChild(dlg); }
  dlg.innerHTML = `<form method="dialog">${html}</form>`;
  dlg.returnValue = '';
  // Enter in a text box saves (otherwise the first button, Cancel or Remove, would be used)
  dlg.querySelectorAll('input').forEach((inp) => inp.addEventListener('keydown', (ev) => {
    if (ev.key === 'Enter') { ev.preventDefault(); const save = dlg.querySelector('button[value="save"]'); if (save) save.click(); }
  }));
  return new Promise((resolve) => { dlg.addEventListener('close', () => resolve({ value: dlg.returnValue, form: dlg.querySelector('form') }), { once: true }); dlg.showModal(); });
}
const clean = (v, max) => String(v || '').replace(/\s+/g, ' ').trim().slice(0, max);
async function addMenu() {
  const full = myNoteCount() >= NOTE_LIMIT;
  const r = await formDialog(`<p class="picktitle">Add to the shopping list</p>
    <div class="picklist">
      <button class="pickopt" value="item"><strong>Item</strong><span class="small">Something to buy, with an amount</span></button>
      <button class="pickopt" value="note"${full ? ' disabled' : ''}><strong>Note</strong><span class="small">${full ? `You have ${NOTE_LIMIT} notes. Delete one to add another.` : `A note for both of you (up to ${NOTE_LIMIT} each)`}</span></button>
    </div><div class="row-btns"><button class="btn ghost" value="">Cancel</button></div>`);
  if (r.value === 'item') editItem(null);
  else if (r.value === 'note') {
    const n = await formDialog(`<p class="picktitle">New note</p>
      <input class="dlg-input" name="text" maxlength="200" placeholder="e.g. Coupon in the glovebox" autocomplete="off">
      <div class="row-btns"><button class="btn ghost" value="">Cancel</button><button class="btn" value="save">Add note</button></div>`);
    const text = clean(n.form.text.value, 200);
    if (n.value === 'save' && text && addNote(text)) toast('Note added');
  }
}
async function editItem(id) { // id null = new item; otherwise edit your own item
  const x = id ? state.extras[id] : null;
  const r = await formDialog(`<p class="picktitle">${x ? 'Edit item' : 'New item'}</p>
    <label class="field">Item<input class="dlg-input" name="name" maxlength="80" placeholder="e.g. Paper towels" value="${esc(x ? x.name : '')}" autocomplete="off"></label>
    <label class="field">Amount <span class="small">(optional)</span><input class="dlg-input" name="amount" maxlength="40" placeholder="e.g. 1 pack" value="${esc(x ? x.amount : '')}" autocomplete="off"></label>
    <div class="row-btns">${x ? '<button class="btn outline-danger" value="delete" style="width:auto;margin-right:auto">Remove</button>' : ''}<button class="btn ghost" value="">Cancel</button><button class="btn" value="save">${x ? 'Save' : 'Add item'}</button></div>`);
  if (r.value === 'delete') { writeExtras({ [id]: null }); toast('Item removed'); return; }
  if (r.value !== 'save') return;
  const name = clean(r.form.name.value, 80), amount = clean(r.form.amount.value, 40);
  if (!name) { toast('Type the item name'); return; }
  if (x) writeExtras({ [id]: { ...x, name, amount } });
  else writeExtras({ [Date.now().toString(36) + Math.random().toString(36).slice(2, 6)]: { name, amount, email: myEmail(), at: Date.now() } });
  toast(x ? 'Item updated' : 'Item added');
}

// ---------------------------------------------------------------- ratings
// One record per meal in Firestore (ratings/<meal id>). It keeps the meal's names too,
// so the calendar can still show the rating after the recipe is removed from the sheet.
const myId = () => (state.user ? state.user.uid : 'local');
const votesNeeded = () => (state.firebase ? 2 : 1);
const ratingId = (mealName) => 'm-' + slug(mealName);
function ratingInfo(mealName) {
  const r = state.ratings[ratingId(mealName)];
  const votes = Object.values((r && r.votes) || {}).filter((v) => typeof v === 'number');
  const mine = r && r.votes ? r.votes[myId()] || 0 : 0;
  const done = votes.length >= votesNeeded();
  const avg = votes.length ? votes.reduce((a, b) => a + b, 0) / votes.length : 0;
  return { record: r || null, count: votes.length, mine, done, avg };
}
const starsText = (avg) => `★ ${avg.toFixed(1)} / 5.0`;
function setRating(meal, stars) {
  const id = ratingId(meal.name);
  const r = state.ratings[id] || (state.ratings[id] = { votes: {} });
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
const mealByName = (name) => allMeals().find((x) => x.name === name) || null;
const mealById = (id) => allMeals().find((x) => x.id === id) || null;
function setPlan(day, mealName) { // mealName: recipe meal name, { text } for a custom entry, or null to remove
  writePlan({ [day]: mealName || null });
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
  return Object.keys(state.plan).filter((k) => k >= today && state.plan[k] === mealName).sort()[0] || null;
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

function mealsSwitch(active) {
  return `<nav class="seg" aria-label="Which meals">
    <a href="#/" class="${active === 'current' ? 'on' : ''}">Current</a>
    <a href="#/past" class="${active === 'past' ? 'on' : ''}">Past</a></nav>`;
}
const kindBadge = (meal) => `<span class="kind chip-${meal.kind === 'Side' ? 1 : 0}">${meal.kind === 'Side' ? 'Side' : 'Main'}</span>`;
function plannedChip(meal) {
  const n = nextPlanned(meal.name);
  return n ? `<span class="planned">📅 ${n === dayKey(new Date()) ? 'Tonight' : esc(fmtDay(keyToDate(n), { weekday: 'short', month: 'short', day: 'numeric' }))}</span>` : '';
}
function renderPast() {
  setTitle(APP_NAME, false); setTab('meals');
  const m = state.model;
  if (!m) { view().innerHTML = '<p class="loading">Loading recipes…</p>'; return; }
  let h = banners() + mealsSwitch('past');
  const filter = ls.get('past-filter', 'all'); // all | Main | Side
  const all = m.past || [];
  const count = (k) => all.filter((x) => x.kind === k).length;
  if (all.length) h += `<nav class="filters" aria-label="Show">${[['all', 'All', all.length], ['Main', 'Mains', count('Main')], ['Side', 'Sides', count('Side')]]
    .map(([v, label, n]) => `<button class="filter${filter === v ? ' on' : ''}" data-pastfilter="${v}">${label} <span>${n}</span></button>`).join('')}</nav>`;
  const past = all.filter((x) => filter === 'all' || x.kind === filter).sort((a, b) => {
    const ra = ratingInfo(a.name), rb = ratingInfo(b.name);
    return (rb.done ? rb.avg : -1) - (ra.done ? ra.avg : -1) || a.name.localeCompare(b.name, undefined, { numeric: true });
  });
  if (all.length && !past.length) h += `<p class="empty">No past ${filter === 'Side' ? 'sides' : 'mains'} yet.</p>`;
  if (!all.length) h += '<p class="empty">No past recipes yet.<br><span class="small">To archive a meal, ask Claude, for example "archive the pot pie". It moves to this list with its rating, and you can still open it or plan it again.</span></p>';
  for (const meal of past) {
    const ri = ratingInfo(meal.name);
    const sides = meal.sides.map((d) => esc(d.name)).join(' &amp; ');
    const status = ri.done ? `<span class="avg">${starsText(ri.avg)}</span>`
      : `<span class="small">${ri.mine ? 'Waiting for the other vote' : ri.count ? 'The other vote is in' : 'Not rated'}</span>`;
    h += `<div class="card meal-card past-card">
      <a class="cardlink" href="#/meal/${meal.id}">
        <div class="eyebrow">${kindBadge(meal)}${plannedChip(meal)}</div>
        <h2>${esc(meal.main.name)}</h2>
        ${sides ? `<p class="sides">with ${sides}</p>` : ''}
      </a>
      <div class="meta">${status}<button class="ratebtn${ri.mine ? ' voted' : ''}" data-ratemeal="${meal.id}" aria-label="Rate ${esc(meal.main.name)}">${ri.mine ? `★ ${ri.mine}` : '☆ Rate'}</button></div>
    </div>`;
  }
  view().innerHTML = h;
}

function renderMeals() {
  setTitle(APP_NAME, false); setTab('meals');
  const m = state.model;
  if (!m) { view().innerHTML = '<p class="loading">Loading recipes…</p>'; return; }
  let h = banners() + mealsSwitch('current');
  if (!state.firebase) h += '<div class="banner">Sharing is not set up yet, so ticks stay on this phone. See SETUP.md.</div>';
  const tonightCustom = customText(state.plan[dayKey(new Date())]);
  if (tonightCustom) h += `<a class="card tonight" href="#/calendar"><span class="section-label">Tonight</span><strong>${esc(tonightCustom)}</strong></a>`;
  const tonight = mealByName(state.plan[dayKey(new Date())]);
  if (tonight) h += `<a class="card tonight" href="#/meal/${tonight.id}"><span class="section-label">Tonight</span><strong>${esc(tonight.main.name)}</strong>${tonight.sides.length ? `<span class="small">with ${tonight.sides.map((d) => esc(d.name)).join(' &amp; ')}</span>` : ''}</a>`;
  // Order: meals planned for today or later (soonest first), then unplanned meals, then rated meals.
  const rank = (meal) => { const n = nextPlanned(meal.name); return n ? [0, n] : [ratingInfo(meal.name).done ? 2 : 1, '']; };
  const sorted = m.meals.map((meal, i) => ({ meal, i, r: rank(meal) }))
    .sort((a, b) => a.r[0] - b.r[0] || a.r[1].localeCompare(b.r[1]) || a.i - b.i).map((x) => x.meal);
  for (const meal of sorted) {
    const ri = ratingInfo(meal.name);
    const sides = meal.sides.map((d) => esc(d.name)).join(' &amp; ');
    const time = meal.planTotal ? esc(meal.planTotal.split(/[.,;]/)[0]) : meal.totalMin != null ? `≈ ${fmtDuration(meal.totalMin)}` : '';
    h += `<a class="card meal-card${ri.done ? ' rated' : ''}" href="#/meal/${meal.id}">
      <div class="eyebrow">${esc(meal.title)}${plannedChip(meal)}</div>
      <h2>${esc(meal.main.name)}</h2>
      ${sides ? `<p class="sides">with ${sides}</p>` : ''}
      <div class="meta">${ri.done ? `<span class="avg">${starsText(ri.avg)}</span>` : `${time ? `<span class="time">⏱ ${time}</span>` : ''}${progressHtml(meal.id, meal.keys)}`}</div>
      ${!ri.done && ri.count ? `<p class="small votes">${ri.mine ? 'You rated it. Waiting for the other vote.' : 'Rated by the other person. Add your vote.'}</p>` : ''}
    </a>`;
  }
  view().innerHTML = h || '<p class="empty">No meals yet.</p>';
  refreshTicks();
}

function renderMeal(id) {
  const meal = mealById(id);
  if (!meal) { if (state.model) location.replace('#/'); else view().innerHTML = '<p class="loading">Loading…</p>'; return; }
  setTitle(meal.title, true); setTab('meals');
  const S = meal.id;
  const hasPlan = !!(meal.planTotal || meal.planFirst || meal.starts.length > 1);
  let h = '<nav class="jump">' + (hasPlan ? '<a href="#" data-jump="' + S + '-plan">Plan</a>' : '') +
    meal.dishes.map((d) => `<a href="#" class="chip-${d.color}" data-jump="${d.anchor}">${esc(shortName(d))}</a>`).join('') + '</nav>';
  h += banners();
  if (meal.archived) h += '<div class="banner">This is a past recipe. You can still cook it, tick it and plan it. To bring it back to the current list, ask Claude to un-archive it.</div>';
  { const ri = ratingInfo(meal.name); h += `<div class="listhead">${progressHtml(S, meal.keys)}<a href="#" data-jump="${meal.id}-rate" class="rate-link">${ri.done ? starsText(ri.avg) : ri.mine ? '★ Rated' : '☆ Rate'}</a></div>`; }

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
  h += ratingCardHtml(meal);
  h += `<button class="btn outline-danger" data-clear="${S}" data-label="${esc(meal.title)}">Clear ticks for this meal</button>`;
  view().innerHTML = h;
  refreshTicks();
}
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
  h += `<div class="weeknav">
    <button class="btn ghost" data-week="-1" aria-label="Previous week">‹</button>
    <div><strong>${esc(fmtDay(start, { month: 'short', day: 'numeric' }))} – ${esc(fmtDay(end, { month: 'short', day: 'numeric' }))}</strong>${label ? `<span class="small">${label}</span>` : '<button class="linkbtn" data-week="0">Back to this week</button>'}</div>
    <button class="btn ghost" data-week="1" aria-label="Next week">›</button></div>`;
  h += '<ul class="days">';
  for (let i = 0; i < 7; i++) {
    const d = new Date(start); d.setDate(d.getDate() + i);
    const k = dayKey(d), custom = customText(state.plan[k]), name = custom ? '' : state.plan[k], meal = name ? mealByName(name) : null;
    const past = k < today;
    h += `<li class="day${k === today ? ' today' : ''}${past ? ' past' : ''}">
      <div class="date"><span>${esc(fmtDay(d, { weekday: 'short' }))}</span><b>${d.getDate()}</b></div>
      <button class="pick" data-day="${k}">${custom
        ? `<strong>${esc(custom)}</strong>`
        : name
        ? (() => {
          const ri = ratingInfo(name), rec = ri.record;
          const main = meal ? meal.main.name : rec && rec.main ? rec.main : displayMealName(name);
          const sides = meal ? meal.sides.map((x) => x.name) : rec && rec.sides ? rec.sides : [];
          return `${meal && meal.kind === 'Side' ? kindBadge(meal) : ''}<strong>${esc(main)}</strong>${sides.length ? `<span class="small">with ${sides.map(esc).join(' &amp; ')}</span>` : ''}${ri.done ? `<span class="avg">${starsText(ri.avg)}</span>` : ''}${meal ? '' : '<span class="small">No longer in the recipes</span>'}`;
        })()
        : `<span class="addmeal">${k === today ? '+ Plan tonight' : '+ Add meal'}</span>`}</button>
      ${meal && k <= today ? (() => { const ri = ratingInfo(meal.name); return `<button class="ratebtn${ri.mine ? ' voted' : ''}" data-rateday="${k}" aria-label="Rate this meal">${ri.mine ? `★ ${ri.mine}` : '☆ Rate'}</button>`; })() : ''}
      ${meal ? `<a class="open" href="#/meal/${meal.id}" aria-label="Open recipe">›</a>` : ''}
    </li>`;
  }
  h += '</ul><p class="small" style="text-align:center">Tap a day to choose or change its meal. Press and hold a meal to drag it to another day. Both phones see the same calendar.</p>';
  view().innerHTML = h;
}
// A plan entry is either a recipe meal name (string) or a custom note ({ text: 'Chipotle + Leftovers' }).
const customText = (v) => (v && typeof v === 'object' && v.text ? String(v.text) : '');
// A meal can sit on only one upcoming day (today or later). Past days don't count, so meals can be repeated.
function plannedElsewhere(mealName, day) {
  const today = dayKey(new Date());
  return Object.keys(state.plan).filter((k) => k !== day && k >= today && state.plan[k] === mealName).sort()[0] || null;
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
async function pickMeal(day) {
  const m = state.model; if (!m) return;
  const current = state.plan[day], currentCustom = customText(current);
  let dlg = $('#picker');
  if (!dlg) { dlg = document.createElement('dialog'); dlg.id = 'picker'; document.body.appendChild(dlg); }
  dlg.innerHTML = `<form method="dialog"><p class="picktitle">${esc(fmtDay(keyToDate(day), { weekday: 'long', month: 'long', day: 'numeric' }))}</p>
    <div class="picklist">${[...m.meals].sort((a, b) => ratingInfo(a.name).done - ratingInfo(b.name).done).map((x) => {
      const other = plannedElsewhere(x.name, day);
      return `<button class="pickopt${x.name === current ? ' sel' : ''}" value="${esc(x.name)}"${other ? ' disabled' : ''}><strong>${esc(x.main.name)}</strong>${x.sides.length ? `<span class="small">with ${x.sides.map((s) => esc(s.name)).join(' &amp; ')}</span>` : ''}${ratingInfo(x.name).done ? `<span class="small avg">${starsText(ratingInfo(x.name).avg)}</span>` : ''}${other ? `<span class="small taken">Already planned for ${esc(fmtDay(keyToDate(other)))}</span>` : ''}</button>`;
    }).join('')}${(m.past || []).length ? `<p class="section-label pastlabel">Past recipes</p>${m.past.map((x) => {
      const other = plannedElsewhere(x.name, day);
      return `<button class="pickopt past${x.name === current ? ' sel' : ''}" value="${esc(x.name)}"${other ? ' disabled' : ''}>${kindBadge(x)}<strong>${esc(x.main.name)}</strong>${x.sides.length ? `<span class="small">with ${x.sides.map((s) => esc(s.name)).join(' &amp; ')}</span>` : ''}${ratingInfo(x.name).done ? `<span class="small avg">${starsText(ratingInfo(x.name).avg)}</span>` : ''}${other ? `<span class="small taken">Already planned for ${esc(fmtDay(keyToDate(other)))}</span>` : ''}</button>`;
    }).join('')}` : ''}</div>
    <div class="custom">
      <label class="small" for="customtext">Or type your own</label>
      <div class="customrow"><input type="text" id="customtext" maxlength="80" placeholder="e.g. Chipotle + Leftovers" value="${esc(currentCustom)}" autocomplete="off"><button class="btn" value="__custom" id="customsave">Save</button></div>
    </div>
    <div class="row-btns">${current ? '<button class="btn outline-danger" value="__none" style="width:auto">Remove</button>' : ''}<button class="btn ghost" value="">Cancel</button></div></form>`;
  dlg.returnValue = '';
  const input = $('#customtext', dlg);
  input.addEventListener('keydown', (ev) => { if (ev.key === 'Enter') { ev.preventDefault(); $('#customsave', dlg).click(); } });
  const choice = await new Promise((resolve) => { dlg.addEventListener('close', () => resolve(dlg.returnValue), { once: true }); dlg.showModal(); });
  if (choice === '__custom') {
    const text = input.value.replace(/\s+/g, ' ').trim().slice(0, 80);
    if (text && text !== currentCustom) setPlan(day, { text });
    return;
  }
  if (choice === '__none') setPlan(day, null);
  else if (choice && choice !== current && !plannedElsewhere(choice, day)) setPlan(day, choice);
}

function renderShopping() {
  setTitle('Shopping list', false); setTab('shopping');
  const m = state.model;
  if (!m) { view().innerHTML = '<p class="loading">Loading…</p>'; return; }
  let h = banners();
  h += `<section class="card notes" id="notescard"${Object.keys(state.notes).length ? '' : ' hidden'}><p class="section-label">Notes</p>
    <ul id="notelist" class="notelist">${notesListHtml()}</ul></section>`;
  h += `<div class="listhead">${progressHtml('shopping', m.shopKeys)}</div>`;
  h += `<section class="card extras" id="extrascard"${Object.keys(state.extras).length ? '' : ' hidden'}><p class="section-label">Added by us</p>
    <ul class="list" id="extralist">${extrasListHtml()}</ul></section>`;
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
  let h = banners();
  h += '<section class="card"><p class="section-label">Sharing</p>';
  if (!state.firebase) h += '<p>Not set up yet. Ticks are saved on this phone only. Follow SETUP.md steps 2–5 to share ticks between phones.</p>';
  else if (state.user) h += `<p>Signed in as <strong>${esc(state.user.email)}</strong>.</p><p class="small">Status: ${esc($('#sync').textContent)}</p><button class="btn ghost" id="signout">Sign out</button>`;
  h += `<form id="nameform"><label class="field">Your name <span class="small">(shown on your notes and added items)</span>
    <div class="customrow"><input type="text" id="myname" maxlength="40" placeholder="${esc((state.user && state.user.email) || 'Your name')}" value="${esc(state.myName)}" autocomplete="off"><button class="btn" type="submit">Save</button></div></label></form>`;
  h += '</section>';

  h += `<section class="card"><p class="section-label">While cooking</p>
    <label class="switch"><span>Keep screen on in a meal</span><input type="checkbox" id="wake" ${ls.get('wake', true) ? 'checked' : ''}></label></section>`;
  h += `<section class="card"><p class="section-label">Updates</p>
    <p class="small">Updates app, loads latest recipes</p>
    ${info.error ? `<p class="err">${esc(info.error)}</p>` : ''}
    <button class="btn ghost" type="button" id="update">Check for updates</button></section>`;
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
  // shopping list only: ticked items move to the bottom of their section, unticked ones keep their order
  document.querySelectorAll('ul.list').forEach((ul) => {
    const items = [...ul.children].filter((li) => li.dataset.key);
    if (!items.length || items[0].dataset.scope !== 'shopping') return;
    items.forEach((li, i) => { if (li.dataset.i == null) li.dataset.i = i; });
    const want = [...items].sort((a, b) => a.classList.contains('done') - b.classList.contains('done') || a.dataset.i - b.dataset.i);
    if (want.some((li, i) => li !== items[i])) want.forEach((li) => ul.appendChild(li));
  });
  document.querySelectorAll('[data-progress]').forEach((el) => {
    const scope = el.dataset.progress, m = state.model; if (!m) return;
    const keys = scope === 'shopping' ? [...m.shopKeys, ...extraKeys()] : (mealById(scope) || {}).keys || [];
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
  const wk = e.target.closest('[data-week]');
  if (wk) { state.calWeek = wk.dataset.week === '0' ? 0 : state.calWeek + Number(wk.dataset.week); route(); return; }
  const star = e.target.closest('[data-rate]');
  if (star) {
    const meal = mealById(star.dataset.meal);
    if (meal) { const n = Number(star.dataset.rate); setRating(meal, ratingInfo(meal.name).mine === n ? null : n); }
    return;
  }
  if (e.target.closest('#addbtn')) { addMenu(); return; }
  const ex = e.target.closest('[data-editextra]');
  if (ex) { editItem(ex.dataset.editextra); return; }
  const dx = e.target.closest('[data-delextra]');
  if (dx) { const x = state.extras[dx.dataset.delextra]; if (x && await confirmDialog(`Remove "${x.name}" from the shopping list for both of you?`, 'Remove')) writeExtras({ [dx.dataset.delextra]: null }); return; }
  const dn = e.target.closest('[data-delnote]');
  if (dn) { if (await confirmDialog('Delete this note for both of you?', 'Delete')) writeNotes({ [dn.dataset.delnote]: null }); return; }
  const pf = e.target.closest('[data-pastfilter]');
  if (pf) { ls.set('past-filter', pf.dataset.pastfilter); route(); return; }
  const rm = e.target.closest('[data-ratemeal]');
  if (rm) { const meal = mealById(rm.dataset.ratemeal); if (meal) rateDialog(meal); return; }
  const rb = e.target.closest('[data-rateday]');
  if (rb) { const meal = mealByName(state.plan[rb.dataset.rateday]); if (meal) rateDialog(meal); return; }
  const pick = e.target.closest('.pick[data-day]');
  if (pick) { if (!suppressClick) pickMeal(pick.dataset.day); return; }
  if (e.target.closest('.delnote')) return;
  const item = e.target.closest('.item[data-key]');
  if (item) { setTick(item.dataset.scope, item.dataset.key, !item.classList.contains('done')); return; }
  const clr = e.target.closest('[data-clear]');
  if (clr) {
    if (await confirmDialog(`Clear all ticks for ${clr.dataset.label}? This clears them on both phones.`)) { clearScope(clr.dataset.clear); toast('Ticks cleared'); }
    return;
  }
  if (e.target.closest('#back')) { if (navCount > 1) history.back(); else location.hash = '#/'; return; }
  if (e.target.id === 'signout') { stopFirestore(); await signOut(state.firebase.auth); return; }
  if (e.target.id === 'update') {
    toast('Checking for updates…');
    const before = state.rowsSig;
    await loadRecipes();
    const recipeMsg = state.recipeInfo && state.recipeInfo.error ? 'Could not load recipes' : state.rowsSig !== before ? 'Recipes updated' : 'Recipes up to date';
    if (location.hash.startsWith('#/settings')) route();
    try {
      const r = await fetch('sw.js', { cache: 'no-store' });
      const v = ((await r.text()).match(/VERSION = '([^']+)'/) || [])[1];
      if (!v) throw new Error('no version found');
      if (v === window.APP_VERSION) { toast(`${recipeMsg} · app is the latest version (${v})`, 3500); return; }
      const reg = await navigator.serviceWorker.getRegistration();
      if (reg) {
        await reg.update();
        const sw = reg.installing || reg.waiting;
        if (sw && sw.state !== 'activated') await new Promise((res) => { sw.addEventListener('statechange', () => sw.state === 'activated' && res()); setTimeout(res, 15000); });
      }
      toast(`${recipeMsg} · updating app to ${v}…`); setTimeout(() => location.reload(), 900);
    } catch (err) { toast(`${recipeMsg} · could not check the app: ${err.message}`, 3500); }
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
  if (e.target.id === 'nameform') {
    e.preventDefault();
    const name = $('#myname').value.replace(/\s+/g, ' ').trim().slice(0, 40);
    state.myName = name;
    if (state.user) { if (name) state.names[myEmail()] = name; else delete state.names[myEmail()]; }
    renderNotes(); renderExtras();
    if (state.firebase && state.user) setDoc(doc(state.firebase.db, 'profiles', state.user.uid), { name, email: state.user.email }, { merge: true }).catch((err) => toast('Could not save: ' + (err.code || err.message)));
    else ls.set('local-name', name);
    toast(name ? `Saved. Your notes and items will show "${name}"` : 'Saved. Your notes and items will show your email');
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
document.addEventListener('visibilitychange', () => { updateWakeLock(); if (document.visibilityState === 'visible') loadRecipes(); });

// ---------------------------------------------------------------- calendar drag and drop
// Press and hold a planned day (~0.35 s), then drag it onto another day. Works with touch and mouse.
let drag = null, suppressClick = false;
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
  const y = scrollY;
  $('#addbtn').hidden = page !== 'shopping';
  if (page === 'meal') renderMeal(decodeURIComponent(arg || ''));
  else if (page === 'shopping') renderShopping();
  else if (page === 'calendar') renderCalendar();
  else if (page === 'past') renderPast();
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
  try { localStorage.removeItem('sheet-url'); } catch { /* old setting, no longer used */ }
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
    state.ratings = ls.get('local-ratings', {});
    state.notes = ls.get('local-notes', {});
    state.myName = ls.get('local-name', '');
    state.extras = ls.get('local-extras', {});
    state.sync = 'local'; renderSync();
  }
  route();
  loadRecipes();
  if ('serviceWorker' in navigator && location.protocol !== 'file:') {
    navigator.serviceWorker.register('sw.js').catch((e) => console.warn('SW failed', e));
  }
}
start();
