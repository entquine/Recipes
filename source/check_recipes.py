"""Check recipes.xlsx before publishing it.

Usage (from the repository root):  python3 source/check_recipes.py [path/to/recipes.xlsx]
Prints a summary of every meal and exits with an error if anything would break the app.
"""
import re
import sys

import openpyxl

SECTIONS = {"makes", "plan total", "plan first", "plan", "ingredient", "step", "tip", "made", "made with", "ahead", "source"}
RECIPE_COLS = ["Recipe", "Type", "Meal", "Section", "Group", "Order", "Amount", "Text", "ID"]
SHOP_COLS = ["Item", "Amount", "Section", "Used in", "List"]
ALL_IDS = {}  # dish ID -> (tab, meal, dish), across both tabs
SHOP_COLS_NEW = ["Shop item", "Shop qty", "Shop unit", "Shop section", "Timer", "Prep task", "Prep amount", "Hide from Recipe Book", "Combined order", "Ahead", "Count up"]
SHOP_SECTIONS = {"Meat", "Produce", "Dairy", "Frozen", "Bakery", "Pantry", "Other"}
SHOP_SEEN = {}     # shop item (lower case) -> (section, unit set, where)
SHOP_MISSING = []  # dishes with ingredients that have no shopping details yet


def same_dish_check(wb, names, warnings):
    """The same dish served with more than one meal (same Recipe name and Type) is one recipe in the app, so its copies must match."""
    copies = {}
    for tab in ("recipes", "archive"):
        if tab not in names:
            continue
        ws = wb[names[tab]]
        head = [str(x or "").strip().lower() for x in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
        col = {h: i for i, h in enumerate(head)}
        keep = [h for h in ("section", "group", "order", "amount", "text", "shop item", "shop qty", "shop unit", "shop section", "timer", "prep task", "prep amount", "ahead", "count up") if h in col]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            sec = str(row[col["section"]] or "").strip().lower()
            if sec not in ("makes", "ingredient", "step", "tip"):
                continue
            key = (str(row[0]).strip().lower(), str(row[col["type"]] or "").strip().lower())
            meal = str(row[col["meal"]] or "").strip()
            copies.setdefault(key, {}).setdefault(meal, []).append(tuple("" if row[col[h]] is None else str(row[col[h]]).strip() for h in keep))
    for (name, _), by_meal in copies.items():
        if len(by_meal) > 1:
            versions = {tuple(sorted(v)) for v in by_meal.values()}
            if len(versions) > 1:
                warnings.append(f"{name.title()!r} is in {len(by_meal)} meals ({', '.join(by_meal)}) but the copies differ. The app shows it as one recipe, so keep every copy identical (or rename one if it's really a different recipe)")


def main(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    errors, warnings = [], []
    names = {n.strip().lower(): n for n in wb.sheetnames}
    if "recipes" not in names:
        print('ERROR: no tab named "Recipes"'); sys.exit(1)
    meals = check_meals(wb[names["recipes"]], "Recipes", errors, warnings)
    if "archive" in names:
        print("Archive (past recipes):")
        past = check_meals(wb[names["archive"]], "Archive", errors, warnings)
        for both in sorted(set(meals) & set(past)):
            errors.append(f"Meal {both!r} is in both Recipes and Archive. Move its rows to one tab only")
    else:
        warnings.append('No "Archive" tab (past recipes). Add one with the same columns as Recipes when archiving a meal')
    same_dish_check(wb, names, warnings)
    shopping_check(wb, names, errors, warnings)
    shop_report(warnings)
    for w in warnings:
        print("warning:", w)
    for e in errors:
        print("ERROR:", e)
    print(f"{len(meals)} current meals checked: {'OK' if not errors else str(len(errors)) + ' error(s)'}")
    sys.exit(1 if errors else 0)


def check_meals(ws, tab, errors, warnings):
    header = [str(c.value or "").strip() for c in ws[1]]
    if header[: len(RECIPE_COLS)] != RECIPE_COLS:
        errors.append(f"{tab} columns must be {RECIPE_COLS}, found {header}")
    extra = [h for h in header[len(RECIPE_COLS):] if h]
    if extra and extra != SHOP_COLS_NEW[: len(extra)]:
        errors.append(f"{tab}: columns after ID must be {SHOP_COLS_NEW}, found {extra}")
    meals = {}
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        recipe, typ, meal, section, group, order, amount, text, rid, s_item, s_qty, s_unit, s_sec, timer, prep, prep_amt, hide, combo, ahead, countup = (list(row) + [None] * 20)[:20]
        if not any(row):
            continue
        where = f"{tab} row {i}"
        if not recipe or not section:
            errors.append(f"{where}: Recipe and Section are required"); continue
        if str(section).strip().lower() not in SECTIONS:
            errors.append(f"{where}: unknown Section {section!r}")
        if str(typ or "").strip() not in ("Main", "Side", "Dessert", "Breakfast"):
            errors.append(f"{where}: Type must be Main, Side, Breakfast or Dessert, found {typ!r}")
        m = meals.setdefault(str(meal or recipe).strip(), {})
        d = m.setdefault(str(recipe).strip(), {"type": str(typ or "").strip(), "steps": 0, "ings": 0, "plan": 0, "planrows": 0, "sideref": False, "ids": set()})
        rid = str(rid or "").strip()
        if not rid:
            errors.append(f"{where}: missing ID (every row of a dish carries the dish's ID)")
        elif not re.fullmatch(r"d-[a-z0-9-]+", rid):
            errors.append(f"{where}: ID {rid!r} should look like d-dish-name (lower-case letters, digits, dashes)")
        else:
            d["ids"].add(rid)
        s = str(section).strip().lower()
        if s == "step":
            d["steps"] += 1
            if order is None:
                warnings.append(f"{where}: Step without an Order number")
        elif s == "ingredient":
            d["ings"] += 1
        elif s in ("plan total", "plan first"):
            d["planrows"] += 1
        if s in ("plan", "plan first") and re.search(r"\bSides? steps?\b", str(text or "")):
            d["sideref"] = True
        if s == "plan":
            d["plan"] += 1
            if amount is not None and not re.fullmatch(r"Before|\d{1,2}:\d{2}(:00)?", str(amount).strip()):
                warnings.append(f"{where}: Plan time should look like 0:27 or Before, found {amount!r}")
        if s == "made with":
            if not re.match(r"^d-[a-z0-9-]+$", str(text or "").strip()):
                errors.append(f"{where}: Made with needs the old main dish's ID in Text (e.g. d-lemon-herb-chicken-cutlets), found {text!r}")
            if amount not in (None, "") and not re.match(r"^\d{4}-\d{2}-\d{2}$", str(amount).strip()):
                errors.append(f"{where}: Made with Amount should be the date it was split off, as YYYY-MM-DD, found {amount!r}")
        if s == "source" and not re.match(r"^https?://", str(text or "").strip()):
            warnings.append(f"{where}: Source should be the recipe's web address (starting with https://), found {text!r}")
        if s == "made":
            d["made"] = d.get("made", 0) + 1
            if d["made"] > 1:
                errors.append(f"{where}: more than one Made row for {recipe!r}; use one row with the total")
            if not re.fullmatch(r"\d{1,3}", str(amount if amount is not None else "").strip()):
                errors.append(f"{where}: Made rows need Amount = how many times it was made before the calendar, as a whole number, found {amount!r}")
        elif not text:
            warnings.append(f"{where}: empty Text")
        if timer not in (None, ""):
            if s != "step":
                errors.append(f"{where}: Timer is only used on Step rows")
            elif not re.search(r"\d+(\s*(-|–|to)\s*\d+)?\s*(minutes?|mins?|hours?|hrs?)\b", str(text or ""), re.I):
                errors.append(f"{where}: has a Timer name but the step text has no duration like '3 minutes'")
            elif len(str(timer)) > 40:
                warnings.append(f"{where}: Timer name is long; keep it under 40 characters, e.g. 'Cook cutlets'")
        if prep not in (None, "") or prep_amt not in (None, ""):
            if s != "step":
                errors.append(f"{where}: Prep task / Prep amount are only used on Step rows")
            elif prep in (None, ""):
                errors.append(f"{where}: has a Prep amount but no Prep task")
            elif len(str(prep)) > 40:
                warnings.append(f"{where}: Prep task is long; keep it short like 'Mince garlic'")
        if countup not in (None, ""):
            if s != "step":
                errors.append(f"{where}: Count up is only used on Step rows")
            elif not re.search(r"\d+\s*(min|hour|hr|day)", str(countup), re.I):
                errors.append(f"{where}: Count up should be the longest time to count to, e.g. '4 hours' or '45 min'; found {countup!r}")
        if ahead not in (None, ""):
            if s != "step":
                errors.append(f"{where}: Ahead is only used on Step rows")
            elif not re.search(r"\d|overnight", str(ahead), re.I):
                warnings.append(f"{where}: Ahead should say how long before cooking, e.g. '30 min', '1 hour', 'overnight'; found {ahead!r}")
        if combo not in (None, ""):
            if s != "step":
                errors.append(f"{where}: Combined order is only used on Step rows")
            else:
                try:
                    float(combo)
                    m.setdefault("_combo", []).append((float(combo), where))
                except (TypeError, ValueError):
                    errors.append(f"{where}: Combined order must be a number, found {combo!r}")
        elif s == "step":
            m.setdefault("_nocombo", []).append(where)
        if hide not in (None, ""):
            hv = str(hide).strip().lower()
            if hv not in ("yes", "no", "y", "n", "true", "false"):
                warnings.append(f"{where}: Hide from Recipe Book should be yes or blank, found {hide!r}")
            elif hv in ("yes", "y", "true") and str(typ or "").strip() == "Main":
                warnings.append(f"{where}: Hide from Recipe Book is set on a Main ({recipe!r}); it's meant for simple sides")
        if s == "ingredient":
            check_shop(where, tab, str(recipe).strip(), s_item, s_qty, s_unit, s_sec, errors)
        elif any(v not in (None, "") for v in (s_item, s_qty, s_unit, s_sec)):
            warnings.append(f"{where}: shopping columns are only used on Ingredient rows")
    for meal, dishes in list(meals.items()):
        combos, missing = dishes.pop("_combo", []), dishes.pop("_nocombo", [])
        if combos:
            if missing:
                warnings.append(f"Meal {meal!r}: has a Combined order but {len(missing)} step(s) don't (first: {missing[0]}); give every step a number")
            nums = [n for n, _ in combos]
            if len(nums) != len(set(nums)):
                warnings.append(f"Meal {meal!r}: two steps share the same Combined order number")
        mains = [n for n, d in dishes.items() if d["type"] == "Main"]
        sides = [n for n, d in dishes.items() if d["type"] == "Side"]
        desserts = [n for n, d in dishes.items() if d["type"] == "Dessert"]
        breakfasts = [n for n, d in dishes.items() if d["type"] == "Breakfast"]
        if tab == "Archive":
            # Archive: one entry per dish. An entry is a single Main or a single Side.
            if len(dishes) != 1:
                errors.append(f"Archive entry {meal!r}: must hold exactly one dish (one entry per dish), found {len(dishes)}")
            for n in sides:
                if dishes[n]["plan"] or dishes[n]["planrows"]:
                    errors.append(f"Archive entry {meal!r}: a Side can't have Plan rows")
            for n in mains + desserts + breakfasts:
                if dishes[n]["sideref"]:
                    warnings.append(f"Archive entry {meal!r}: Plan text still mentions a side step; rewrite it without the side")
        else:
            if len(mains) + len(breakfasts) != 1:
                errors.append(f"Meal {meal!r}: needs exactly one Main (or, for a breakfast meal, one Breakfast), found {len(mains)} Main and {len(breakfasts)} Breakfast")
            if not 0 <= len(sides) <= 2:
                warnings.append(f"Meal {meal!r}: usually one or two sides, found {len(sides)}")
        if tab == "Recipes" and not re.match(r"^\d+ ", meal):
            warnings.append(f"Meal {meal!r}: start the name with a number (e.g. '5 Salmon Night') to set its place in the list")
        for n, d in dishes.items():
            if len(d["ids"]) > 1:
                errors.append(f"{tab} dish {n!r}: rows have different IDs {sorted(d['ids'])}. Use one ID for all of the dish's rows")
            for i in d["ids"]:
                other = ALL_IDS.get(i)
                if other and other != (tab, meal, n):
                    errors.append(f"ID {i!r} is used by two dishes: {other[2]!r} ({other[0]}) and {n!r} ({tab}). IDs must be unique and never reused")
                ALL_IDS[i] = (tab, meal, n)
        print(f"- {meal}: " + "; ".join(f"{n} ({d['type']}, {next(iter(d['ids']), 'no ID')})" for n, d in dishes.items()))
    return meals


def check_shop(where, tab, dish, item, qty, unit, sec, errors):
    item = str(item or "").strip()
    if not item:
        SHOP_MISSING.append((tab, dish))
        return
    if item.lower() in ("(none)", "none"):
        return
    if qty not in (None, ""):
        try:
            q = float(str(qty).replace(",", "."))
            if q <= 0:
                raise ValueError
        except ValueError:
            errors.append(f"{where}: Shop qty must be a plain number like 2, 0.5 or 1.25 (or blank for 'as needed'), found {qty!r}")
        if not str(unit or "").strip():
            errors.append(f"{where}: Shop unit is required when there's a Shop qty (use 'each' for things you count)")
    sec = str(sec or "").strip()
    if sec not in SHOP_SECTIONS:
        errors.append(f"{where}: Shop section must be one of {sorted(SHOP_SECTIONS)}, found {sec!r}")
    seen = SHOP_SEEN.setdefault(item.lower(), {"section": sec, "units": set(), "where": where, "name": item})
    if seen["section"] != sec:
        errors.append(f"{where}: Shop item {item!r} is in section {sec!r} here but {seen['section']!r} at {seen['where']}. Use one section per item")
    if seen["name"] != item:
        errors.append(f"{where}: Shop item {item!r} is spelled {seen['name']!r} at {seen['where']}. Use exactly the same name so they combine")
    if str(unit or "").strip():
        seen["units"].add(str(unit).strip())


def shop_report(warnings):
    for k, v in SHOP_SEEN.items():
        if len(v["units"]) > 1:
            warnings.append(f"Shop item {v['name']!r} uses several units {sorted(v['units'])}; amounts only add up within the same unit")
    if SHOP_MISSING:
        dishes = sorted({f"{d} ({t})" for t, d in SHOP_MISSING})
        print(f"Shopping details missing on {len(SHOP_MISSING)} ingredient rows in {len(dishes)} dishes: the app keeps using the Shopping tab until every current or planned dish has them.")
        for d in dishes[:40]:
            print("  -", d)
    else:
        print("Shopping details: complete. The app builds the shopping list from the recipes.")


def shopping_check(wb, names, errors, warnings):
    if "shopping" in names:
        sh = wb[names["shopping"]]
        sheader = [str(c.value or "").strip() for c in sh[1]]
        if sheader[: len(SHOP_COLS)] != SHOP_COLS:
            errors.append(f"Shopping columns must be {SHOP_COLS}, found {sheader}")
        lists = {str(r[4]).strip() for r in sh.iter_rows(min_row=2, values_only=True) if r and r[0] and len(r) > 4 and r[4] is not None}
        blank = sum(1 for r in sh.iter_rows(min_row=2, values_only=True) if r and r[0] and (len(r) < 5 or r[4] is None or not str(r[4]).strip()))
        if blank:
            errors.append(f"Shopping: {blank} item(s) have no List value. Every row needs the list's date")
        if len(lists) > 1:
            errors.append(f"Shopping: all rows must have the same List value, found {sorted(lists)}")
        for l in lists:
            if not re.fullmatch(r"\d{4}-\d{2}-\d{2}(-\d+)?", l):
                errors.append(f"Shopping: List {l!r} should be a date like 2026-11-09 (add -2 for a second list on the same day)")
        print(f"Shopping list: {sum(1 for r in sh.iter_rows(min_row=2, values_only=True) if r and r[0])} items, list {', '.join(sorted(lists)) or '(none)'}")
    else:
        warnings.append('No "Shopping" tab')


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "recipes.xlsx")
