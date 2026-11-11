# Tire Stock Manager

A CLI application for managing tire inventory in a warehouse. Built to solve a real problem: a 1500+ tire set warehouse with no structured search, where finding a specific size took too long.

## Problem

The warehouse was organized by shelves and floors, but there was no system to find a tire by size or brand. Workers had to physically search through rows of tires. Misplacement was common.

## Solution

A Python CLI tool that:
- Stores 1500+ tire sets in an Excel database (no external DBMS needed)
- Searches by size (`205/55/16`) or brand
- Shows exact location: sector, floor, quantity
- Supports adding new tires and removing sold ones
- Caches parsed data with pickle + hash validation to reduce startup time on repeat runs

## Features

- **Search by size** — enter `205/55/16`, get all matching positions with location
- **Search by brand** — case-insensitive partial match
- **Full inventory view** — all tires sorted by sector, floor, quantity
- **Add / sell operations** — directly in Excel file
- **Fast startup** — progress bar with real loading stages, pickle cache, hash validation
- **Data repair script** — `fix_excel.py` fixes corrupted numeric columns (float/string → int)

## Tech Stack

- Python
- pandas, openpyxl — Excel read/write
- pickle + hashlib-style file hash — caching layer
- threading — animated progress bar during startup
