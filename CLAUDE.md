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
| Section | `Makes`, `Plan total`, `Plan first`, `Plan`, `Ingredient`, `Step`, or `Tip` |
| Group | Optional ingredient sub-group, e.g. `Teriyaki sauce` |
| Order | Step number, or the order of Plan rows |
| Amount | Ingredient rows: quantity (`1¼ lb`). Step rows: the "amount used" in that step (`2 chicken breasts`). Plan rows: time from the start, `0:27`, or `Before` |
| Text | The ingredient, step, tip or plan text |

- Plan rows (`Plan total`, `Plan first`, `Plan`) go on the **Main** dish's rows only.
- `Plan first` says which dish to start first. This matters when a side takes longer than the main.
- Plan text can reference steps like `(Carrot steps 2-6)`, `(Main step 3)` or `(Side step 1)`. The word before "step" must be `Main`, `Side`, or a word from the dish's name. The app uses these to build the "Start order" list.
- Keep amounts and Plan times as **text** (e.g. `0:27`, not a spreadsheet time value).
- Ticks, calendar entries and ratings are tied to names and text. Renaming a Meal resets its ticks, and editing a step's text resets that step's tick. Ratings and calendar entries are keyed on the Meal name, so **don't rename an existing Meal** unless asked.

### Tab "Archive" (past recipes)
Same columns as **Recipes**. Meals here appear in the app under **Meals → Past**: they keep their ratings and calendar history, can still be opened and cooked, and can be planned again from the calendar.
- **Archive a meal** ("archive the pot pie"): move all of that meal's rows from Recipes to Archive, unchanged. Keep the exact Meal name, because ratings and calendar history are tied to it. Then remove its ingredients from the Shopping tab (or trim the amount and "Used in" text for items that other current meals still use).
- **Un-archive a meal**: move its rows back to Recipes, and add its ingredients back to Shopping.
- A Meal must be in only one of the two tabs. The checker reports an error if it's in both.
- Meals are never archived automatically. Only archive when the owner asks.

### Tab "Shopping"
Columns: `Item`, `Amount`, `Section` (`Meat`, `Produce`, `Dairy`, `Pantry`), `Used in`. One row per item. When adding meals, add their ingredients here too, merging with items that are already listed (update the amount and the "Used in" text).

## Changing the app itself
- Readable source: `source/app.js`, `source/sw.template.js`. Root `styles.css`, `index.html`, `manifest.webmanifest` and `config.js` are edited directly.
- Build: `cd source && npm install && node build.mjs`. This writes the bundled `app.js` and `sw.js` in the root. Commit both, plus the source changes. Don't commit `node_modules`.
- `config.js` holds the Firebase web config (public by design; security comes from the Firestore rules in `source/firestore.rules`).
- Shared data lives in Firestore: `ticks/*`, `calendar/plan`, `ratings/*`. (`config/app` held an old Google Sheet link and is no longer read.)
- Recipes come only from `recipes.xlsx` in this repository. There is deliberately no other recipe source (the old Google Sheet option was removed).
- After pushing app changes, tell the owner the new version code (shown at the bottom of Settings, from `sw.js`) and to tap **Settings → Check for app update**.
