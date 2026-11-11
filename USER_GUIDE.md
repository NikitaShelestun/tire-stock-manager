# Tire Stock Manager — User Guide

A command-line tool for managing tire inventory in a warehouse.

---

## Table of Contents

- [What this program does](#what-this-program-does)
- [How to start](#how-to-start)
- [Main menu](#main-menu)
- [Search by size](#search-by-size)
- [Search by brand](#search-by-brand)
- [Add a new tire](#add-a-new-tire)
- [Sell tires](#sell-tires)
- [Data file structure](#data-file-structure)
- [Tips](#tips)
- [Troubleshooting](#troubleshooting)

---

## What this program does

This program helps you manage tires in a warehouse. You can:

- Find tires by size or brand
- See where tires are stored (sector and floor)
- Add new tires to the database
- Remove tires when they are sold

---

## How to start

1. Place `tires_data.xlsx` in the same folder as the program.
2. Double-click `TireManager.exe`.
3. The program starts and loads the inventory.

---

## Main menu

When the program starts, you will see the main menu

Type the number and press **Enter**.

---

## Search by size

1. Press `1` in the main menu.
2. Type the size in the format `205/55/16`.
3. Press **Enter**.
4. The program shows all tires matching this size, including location and quantity.


Press **Enter** to return to the main menu.

> Type `0` to cancel and go back.

---

## Search by brand

1. Press `2` in the main menu.
2. Type the brand name (case-insensitive).
3. Press **Enter**.

The program performs a partial match — typing `mich` will find `Michelin`.

Press **Enter** to return to the main menu.

---

## Add a new tire

1. Press `4` in the main menu.
2. Enter the requested fields one by one:

| Field | Example | Notes |
|---|---|---|
| Brand | `Michelin` | Required |
| Width | `205` | Required |
| Profile | `55` | Required |
| Diameter | `16` | Required |
| Season | `Winter` | Summer / Winter / AllSeason |
| Condition | `New` | New / Used |
| Sector | `A` | Warehouse sector |
| Floor | `2` | Floor within sector |
| Quantity | `8` | Number of pieces |

3. The program assigns a new ID and saves the data automatically.

---

## Sell tires

1. Press `5` in the main menu.
2. Enter the **ID** of the tire (use option `3` to see all IDs).
3. Enter the quantity to sell:
   - Leave empty and press **Enter** → the whole position is removed.
   - Enter a number → the quantity is reduced. If the remaining quantity is zero, the position is removed.

The program updates the Excel file automatically.

---

## Data file structure

The Excel file `tires_data.xlsx` must contain these columns:

| Column | Type | Description |
|---|---|---|
| `ID` | number | Auto-assigned |
| `BRAND` | text | Tire brand |
| `WIDTH` | number | Tire width |
| `PROFILE` | number | Tire profile |
| `DIAMETER` | number | Rim diameter |
| `SEASON` | text | Summer / Winter / AllSeason |
| `CONDITION` | text | New / Used |
| `SECTOR` | text | Warehouse sector (A, B, C…) |
| `FLOOR` | number | Floor within sector |
| `QUANTITY` | number | Pieces in stock |

**Example data:**

| ID | BRAND | SIZE | SEASON | CONDITION | SECTOR | FLOOR | QTY |
|---|---|---|---|---|---|---|---|
| 1 | Michelin | 205/55/R16 | Winter | New | A | 2 | 8 |
| 2 | Bridgestone | 205/55/R16 | Winter | Used | B | 4 | 12 |

---

## Tips

- Use option `3` (Show all tires) to see every ID.
- Brand search is case-insensitive: `MICHELIN`, `michelin`, `Michelin` all work.
- If you type something invalid, the program asks again.
- Type `0` during search to return to the main menu.
- The program saves changes automatically — no manual save needed.

---

## Troubleshooting

### The program does not start

- Make sure `tires_data.xlsx` is in the same folder as the program.
- Make sure the Excel file has all required columns (see [Data file structure](#data-file-structure)).

### A tire is not found

- Check the size format: `205/55/16`, not `205-55-16`.
- Try searching by brand instead.

### The Excel file is corrupted or numbers look wrong

- Run `fix_excel.py` to normalize numeric columns.
- Then restart the program.
