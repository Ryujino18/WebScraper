# MeITY Startup Hub — India State Scraper

A command-line tool that fetches **Startup** or **Incubator** listings for **any Indian state** from the [MeITY Startup Hub API](https://msh.meity.gov.in) and exports rich contact data to a formatted Excel file.

## Features

- 🔍 **Interactive** — prompts you to choose **Startups** or **Incubators** and then the **state**
- 🇮🇳 Supports all **28 states + 8 UTs** (with custom input support for unlisted names)
- 📧 Extracts full contact info: Email, Phone, Address, Social Links
- ⚡ Concurrent fetching (threading) for speed
- 📊 Exports to a styled Excel file (`<state>_<type>.xlsx`)

## Prerequisites

- Python 3.x
- `pip` (Python package manager)

## Installation

```bash
git clone https://github.com/YOUR_USERNAME/webscrambler.git
cd webscrambler
pip install -r requirements.txt
```

## Usage

```bash
python meity_scraper.py
```

The script will ask two questions:

```
What do you want to fetch?
  [1] Startups
  [2] Incubators

Enter 1 or 2 (default: 1): 1

Enter state name (e.g. Karnataka): Maharashtra
```

Then it will:
1. Fetch **all** listings from the MeITY API
2. Filter for the chosen **state**
3. Enrich each entry with detailed contact info
4. Save results to e.g. **`maharashtra_startups.xlsx`**

## Example Output Files

| State | Type | Output File |
|---|---|---|
| Karnataka | Startups | `karnataka_startups.xlsx` |
| Maharashtra | Incubators | `maharashtra_incubators.xlsx` |
| Tamil Nadu | Startups | `tamil_nadu_startups.xlsx` |

## Project Structure

```
webscrambler/
├── meity_scraper.py              # Main interactive scraper
├── scrape_karnataka_incubators.py  # Legacy Karnataka-only script
├── requirements.txt
├── .gitignore
└── README.md
```

> **Note:** Output `.xlsx` files are excluded from Git via `.gitignore`.
