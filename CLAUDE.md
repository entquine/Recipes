# What's Cooking with Brian: recipe app

A small installable web app (PWA) for two people cooking dinner. It is hosted on GitHub Pages from the `main` branch of this repository, so **anything pushed to `main` is live within about 2 minutes**. The owner is not a developer: explain things in plain language and don't ask them to run commands.

## Adding or changing recipes (the usual task)

All recipes and the shopping list live in **`recipes.xlsx`** in the repository root. The app loads this file directly, so updating the recipes means editing this file, checking it, and pushing it. No rebuild is needed.

1. Edit `recipes.xlsx` with openpyxl (`pip install openpyxl` if needed). **Add to the existing rows. Never drop existing meals unless the owner asks for that specific meal to be removed.**
2. Run `python3 source/check_recipes.py`. Fix every ERROR. Read the warnings.
3. Commit with a message naming the meals added or changed, then push to `main`.
4. Tell the owner the change is live in about 2 minutes, and that the phones pick it up the next time the app is opened.

If the owner gives you a new spreadsheet made elsewhere, **merge its rows into `recipes.xlsx`** instead of replacing the file, unless they say to replace everything.

### Tab "Recipes" (one row per item; columns in this order)
| Column | Meaning |
|---|---|
| Recipe | Dish name, e.g. `Garlic Butter Ribeye Steak Bites` |
| Type | `Main` or `Side`. Each meal has exactly one Main and usually one or two Sides |
| Meal | Groups dishes into a meal, e.g. `4 Steak Bites`. The leading number sets the order in the app and is hidden. New meals take the next number |
| Section | `Makes`, `Plan total`, `Plan first`, `Plan`, `Ingredient`, `Step`, `Tip`, or `Made` |
| Group | Optional ingredient sub-group, e.g. `Teriyaki sauce` |
| Order | Step number, or the order of Plan rows |
| Amount | Ingredient rows: quantity (`1¼ lb`). Step rows: the "amount used" in that step (`2 chicken breasts`). Plan rows: time from the start, `0:27`, or `Before` |
| Text | The ingredient, step, tip or plan text |
| ID | The dish's permanent ID, e.g. `d-lemon-herb-chicken-cutlets`. **Every row of a dish carries the same ID.** See "Dish IDs" below |
| Shop item | **Ingredient rows only.** What to buy, as it appears on a shopping list, e.g. `Garlic`, `Chicken breast`, `Olive oil`. The **same spelling everywhere**, so amounts combine across dishes. `(none)` for things never bought (water) |
| Shop qty | **Ingredient rows only.** How much to buy for this dish, as a plain number: `2`, `0.5`, `1.25`. Blank = "as needed" (e.g. salt) |
| Shop unit | **Ingredient rows only.** The unit for Shop qty, in shopping terms: `lb`, `oz`, `cup`, `tbsp`, `tsp`, `clove`, `head`, `bunch`, `can`, `package`, `jar`, `each`. Use **one unit per Shop item** across all dishes (e.g. garlic is always `clove`) |
| Shop section | **Ingredient rows only.** `Meat`, `Produce`, `Dairy`, `Frozen`, `Bakery`, `Pantry` or `Other`. One section per Shop item |
| Timer | **Step rows only, optional.** A short name for the step's timer, e.g. `Cook cutlets`, `Roast potatoes`. Only on steps whose text has a duration ("3 minutes") |
| Prep task | **Step rows only, optional.** Marks a step that's pure preparation and can be done before cooking starts (chopping, mincing, zesting, measuring a sauce). A short, generic task name, e.g. `Mince garlic`, `Chop parsley`, `Zest and juice lemon`, `Whisk teriyaki sauce`. Use the **same wording** across dishes so the app combines them |
| Prep amount | **Step rows only, optional.** How much this step preps, e.g. `2 cloves`, `1 tbsp`, `1 lemon`. Leave blank if it doesn't apply |
| Hide from Recipe Book | **Optional.** `yes` on a dish's rows keeps that dish out of the app's Recipe Book list and its counts. It still shows in its meal on the Planned list (and on the calendar, shopping list, etc.). Meant for simple sides like "rice, as directed on the package" that are worth offering as a side but not worth listing as a recipe. The checker warns if it's set on a Main |

- Plan rows (`Plan total`, `Plan first`, `Plan`) go on the **Main** dish's rows only.
- `Plan first` says which dish to start first. This matters when a side takes longer than the main.
- Plan text can reference steps like `(Carrot steps 2-6)`, `(Main step 3)` or `(Side step 1)`. The word before "step" must be `Main`, `Side`, or a word from the dish's name. The app uses these to build the "Start order" list.
- Keep amounts and Plan times as **text** (e.g. `0:27`, not a spreadsheet time value).
- Ratings and calendar entries are keyed on the **main dish's ID**, so renaming a dish or meal is safe as long as the ID stays. Ticks are tied to names and text: renaming a Meal resets its ticks, and editing a step's text resets that step's tick.

### Dish IDs (permanent, never change them)
- Each dish has one ID, in the **ID** column (last column) of the Recipes and Archive tabs, repeated on every row of that dish.
- Format: `d-` plus the dish name in lower case with dashes, set **once** when the dish is first added, e.g. `d-garlic-green-beans`. If that ID is already taken (now or in the past), add `-2`, `-3`, and so on.
- **Never change, reuse or swap an ID**, even when a dish or meal is renamed, edited, moved between tabs, split or re-paired. A meal's rating lives under its Main dish's ID, so it follows the Main into the archive automatically. Sides have their own IDs and start unrated once archived on their own.
- A new version of an existing dish that should keep the old ratings keeps the old ID. A genuinely different recipe with the same name gets a new ID (e.g. the Recipe Book `Roasted Broccoli` in Archive is `d-roasted-broccoli-2`).
- The checker reports missing, mixed or duplicate IDs.

### Tab "Archive" (past recipes): one entry per dish
Same columns as **Recipes**, but each entry is a **single dish**: one Main, or one Side on its own. The entry's **Meal** value is its key, shown in the app under **Meals → Recipe Book** with a Main or Side label (from **Type**). (In the app, the Recipes tab is the **Planned** list; the Recipe Book shows every recipe, current and archived, each dish (main or side) on its own line, grouped as Planned / Made before / Want to try, with how many times each has been made.) Archived entries keep their ratings and calendar history, open as full recipes, and can be planned again from the calendar.
- **Archive a current meal** ("archive the pot pie") splits it into one entry per dish:
  - The **Main** keeps its **ID** (that's what carries the ratings and calendar history over) and, by convention, its original **Meal** name, leading number included (e.g. `2 Chicken Pot Pie`). Keep its Makes / Plan total / Plan first / Plan rows, but **rewrite the Plan rows to drop the sides**: remove any `Side step` references and any steps, times or "start first" notes that only concern a side. Renumber the Plan `Order` if rows are removed.
  - Each **Side** becomes its own entry with **Meal = the dish's own name** (e.g. `Roasted Broccoli`) and no number, keeping its own **ID**. A Side keeps its Makes, Ingredient, Step and Tip rows, and has **no Plan rows**.
  - Then remove the meal's ingredients from the Shopping tab (or trim the amount and "Used in" text for items that other current meals still use).
- **Un-archive a dish**: move its rows back to **Recipes**, IDs unchanged, and re-pair it with its meal. A Main gets its original Meal name back (it never lost it). A Side takes the Meal name of the main it's served with, and that Main's Plan rows are updated to include the side again (`Side step` references, timing, "start first" if the side takes longer). Add the ingredients back to Shopping.
- A Meal name must be in only one of the two tabs. The checker reports an error if it's in both.
- The checker enforces: each Archive entry holds exactly one dish; a Side has no Plan rows. It warns if an archived Main's Plan text still mentions a side step.
- The leading number in a Meal name only orders the **Current** list. A dish added straight to Archive doesn't need a number.
- Nothing is archived automatically. Only archive when the owner asks.

**Step timers:** a step whose text has a duration (e.g. "Roast 20-25 minutes", "Simmer 1 hour") gets a timer button in the app, using the first number. Write durations as digits with "minutes"/"hours" when a step involves waiting. Give every such step a short **Timer** name (what's happening, not the step number), e.g. `Cook cutlets`, `Rest steak`, `Simmer sauce`. Without one, the timer is called "<dish>: step N".

### Times made (Section `Made`)
The app shows how many times each recipe has been made: every past day it was on the app's calendar counts automatically, **plus** an optional `Made` row for times **before** it was tracked on the calendar.
- One `Made` row per dish at most, on that dish's own rows (Recipe, Type, Meal, ID as usual). **Amount** = the number of times made before the calendar, as a whole number (`4`). **Text** = optional rough "last made", e.g. `Oct 2026` (shown only until the calendar has a date). Other columns blank.
- Put it on the dish that was made: the Main for a meal (its sides share the meal's calendar days and, unless they have their own `Made` row, the Main's count), or an archived Side entry for a side made on its own.
- In the Recipe Book, a meal on the Planned list moves from **Planned** to **Made before** once it has been made (unless it's on the calendar for an upcoming day).
- Never count days that are on the app's calendar; the app adds those. If unsure whether a time was on the calendar, ask the owner.
- When the owner says how many times something was made, add or update its `Made` row. Don't invent numbers: ask.
- The owner can also nudge the shown number in the app (press and hold a recipe in the Recipe Book). That change is stored in the app as a difference and stays on top of whatever the `Made` row says, so if you change a `Made` row for a dish the owner has adjusted, tell them to check the number shown and correct it in that menu if needed.

### Recipe ideas from the app
The app's **Planning** tab collects recipe ideas (links or plain text) from both people. Its "Copy for recipe Claude" button makes a message starting "Recipe ideas from the app's Planning tab". When the owner pastes one: add each idea as a new meal (next Meal number, new Dish IDs, a Main with one or two fitting sides, every column filled including Shop, Timer and Prep task). For a link, fetch the page and follow that recipe; if it can't be read, say so. For a vague idea, propose a recipe and confirm before adding it. Then check and push as usual, and list the meals added.

### Prep list
Steps with a **Prep task** appear in a "Prep list" at the top of the meal, gathered across the main and its sides and combined by task name, e.g. "Mince garlic: 2 cloves · Chicken + 3 cloves · Garlic bread". Ticking a prep task ticks those steps in each dish. Only mark steps that can really be done up front: not cooking steps, and not prep that has to happen mid-cook (e.g. "slice the rested steak").

### The shopping list is built from the recipes
The app makes the shopping list itself from the **Shop** columns: every current meal (Recipes tab), plus any dish planned on the calendar for today or later (including archived ones), combined by **Shop item** and summed per **Shop unit**. Each line says which meals it's for. Ticking a line marks it bought for those meals; once a meal has been cooked (its calendar day has passed), its items come back for next time.
- **Every Ingredient row needs Shop item, Shop section, and (unless "as needed") Shop qty + Shop unit**, in Recipes and Archive. The checker lists dishes still missing them.
- Shop qty is how much to **buy for that dish**, rounded to what you'd buy where it matters (e.g. a recipe using "2 tbsp lemon juice" → `Lemons`, `1`, `each`).
- Keep names generic and consistent: `Garlic` (not "garlic cloves, minced"), `Chicken breast`, `Butter`. Check the existing Shop items before adding a new name.
- Until every current and planned dish has shopping details, the app keeps showing the **Shopping** tab instead.

### Tab "Shopping" (old: only used until the Shop columns are complete)
Once every dish has Shop columns, the app ignores this tab; stop updating it then. Until then: columns: `Item`, `Amount`, `Section` (`Meat`, `Produce`, `Dairy`, `Pantry`), `Used in`, `List`. One row per item. When adding meals, add their ingredients here too, merging with items that are already listed (update the amount and the "Used in" text).

**`List` = the list's date, the same on every row** (e.g. `2026-11-09`). It decides when the shopping list starts fresh:
- **Replacing the list for a new week** (a new shop): set a **new, later date** on every row. In the app, everything starts unticked again (even items with the same name and amount as last week), items added in the app that were ticked are removed as bought, unticked ones carry over, and last week's ticks are deleted.
- **Fixing or adding to the current list mid-week**: keep the **same** date, so existing ticks stay.
- A second new list on the same day: add `-2` (e.g. `2026-11-09-2`). Dates must only ever go forward.
- The checker reports a missing, mixed or malformed `List`.

## Changing the app itself
- Readable source: `source/app.js`, `source/sw.template.js`. Root `styles.css`, `index.html`, `manifest.webmanifest` and `config.js` are edited directly.
- Build: `cd source && npm install && node build.mjs`. This writes the bundled `app.js` and `sw.js` in the root and updates `source/version.json`. Commit all three, plus the source changes. The version shown in Settings counts up (`wcwb-v-001`, `wcwb-v-002`, …): each build goes one past the last committed `source/version.json` when anything changed, so build once more after any merge or rebase and commit the result. Don't commit `node_modules`.
- `config.js` holds the Firebase web config (public by design; security comes from the Firestore rules in `source/firestore.rules`).
- Shared data lives in Firestore: `ticks/*` (`ticks/m-<meal>` per meal, `ticks/shopping-<List date>` for the current shopping list, including notes and added items), `calendar/plan` (days hold a main dish ID or `{text}`), `ratings/<main dish ID>`, `notes/shopping` (shopping-list notes), `shopping/extra` (items added to the shopping list in the app: name, amount, section, author email), `profiles/<uid>` (each person's name and email). Also `shopping/bought` (for the automatic list: per meal ID, which shop items are bought and since when), `cooking/session` (running step timers with who started them) and `shopping/staples` (shopping items marked "we usually have this", keyed by item name; they're shown in a folded card and don't count toward the list). `history/adjust` holds, per dish ID, a +/- change the owner made to the shown "times made" (press and hold a Recipe Book line); it's added on top of the calendar + spreadsheet count. `ideas/list` holds recipe ideas from the app's **Planning** tab (text or links, author email, and when they were sent to Claude). Notes and added items store the author's email; the app shows the matching name from `profiles`, so names update everywhere when changed. (`config/app` held an old Google Sheet link and is no longer read.)
- Recipes come only from `recipes.xlsx` in this repository. There is deliberately no other recipe source (the old Google Sheet option was removed).
- **Free-plan check for anything that syncs.** The project is on Firebase's free Spark plan (50,000 reads, 20,000 writes, 20,000 deletes per day; 1 GB stored). Whenever a change adds or changes Firestore data or listeners, estimate reads per app open and per change for two people on a heavy day, and tell the owner how far that is from the limits. Prefer one document (with a map inside) over many documents for small shared lists, since each document a listener loads costs a read on every app open. As of the automatic shopping list, cooking timers, staples, Planning ideas and times-made adjustments, a heavy day is about 1,700 reads and 260 writes, and each rated meal adds about 1 read per app open.
- After pushing app changes, tell the owner the new version code (shown at the bottom of Settings, from `sw.js`) and to tap **Settings → Check for updates**.
