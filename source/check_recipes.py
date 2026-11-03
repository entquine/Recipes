"""Check recipes.xlsx before publishing it.

Usage (from the repository root):  python3 source/check_recipes.py [path/to/recipes.xlsx]
Prints a summary of every meal and exits with an error if anything would break the app.
"""
import re
import sys

import openpyxl

SECTIONS = {"makes", "plan total", "plan first", "plan", "ingredient", "step", "tip"}
RECIPE_COLS = ["Recipe", "Type", "Meal", "Section", "Group", "Order", "Amount", "Text"]
SHOP_COLS = ["Item", "Amount", "Section", "Used in"]


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
    shopping_check(wb, names, errors, warnings)
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
    meals = {}
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        recipe, typ, meal, section, group, order, amount, text = (list(row) + [None] * 8)[:8]
        if not any(row):
            continue
        where = f"{tab} row {i}"
        if not recipe or not section:
            errors.append(f"{where}: Recipe and Section are required"); continue
        if str(section).strip().lower() not in SECTIONS:
            errors.append(f"{where}: unknown Section {section!r}")
        if str(typ or "").strip() not in ("Main", "Side"):
            errors.append(f"{where}: Type must be Main or Side, found {typ!r}")
        m = meals.setdefault(str(meal or recipe).strip(), {})
        d = m.setdefault(str(recipe).strip(), {"type": str(typ or "").strip(), "steps": 0, "ings": 0, "plan": 0, "planrows": 0, "sideref": False})
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
        if not text:
            warnings.append(f"{where}: empty Text")
    for meal, dishes in meals.items():
        mains = [n for n, d in dishes.items() if d["type"] == "Main"]
        sides = [n for n, d in dishes.items() if d["type"] == "Side"]
        if tab == "Archive":
            # Archive: one entry per dish. An entry is a single Main or a single Side.
            if len(dishes) != 1:
                errors.append(f"Archive entry {meal!r}: must hold exactly one dish (one entry per dish), found {len(dishes)}")
            for n in sides:
                if dishes[n]["plan"] or dishes[n]["planrows"]:
                    errors.append(f"Archive entry {meal!r}: a Side can't have Plan rows")
            for n in mains:
                if dishes[n]["sideref"]:
                    warnings.append(f"Archive entry {meal!r}: Plan text still mentions a side step; rewrite it without the side")
        else:
            if len(mains) != 1:
                errors.append(f"Meal {meal!r}: needs exactly one Main, found {len(mains)}")
            if not 0 <= len(sides) <= 2:
                warnings.append(f"Meal {meal!r}: usually one or two sides, found {len(sides)}")
        if tab == "Recipes" and not re.match(r"^\d+ ", meal):
            warnings.append(f"Meal {meal!r}: start the name with a number (e.g. '5 Salmon Night') to set its place in the list")
        print(f"- {meal}: " + "; ".join(f"{n} ({d['type']}, {d['ings']} ingredients, {d['steps']} steps)" for n, d in dishes.items()))
    return meals


def shopping_check(wb, names, errors, warnings):
    if "shopping" in names:
        sh = wb[names["shopping"]]
        sheader = [str(c.value or "").strip() for c in sh[1]]
        if sheader[: len(SHOP_COLS)] != SHOP_COLS:
            errors.append(f"Shopping columns must be {SHOP_COLS}, found {sheader}")
        print(f"Shopping list: {sum(1 for r in sh.iter_rows(min_row=2, values_only=True) if r and r[0])} items")
    else:
        warnings.append('No "Shopping" tab')


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "recipes.xlsx")
