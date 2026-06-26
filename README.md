# leapy

A Google Maps lead generation tool that searches for businesses, extracts contact information, and scrapes emails and social media links from their websites.

## What It Does

Searches Google Maps for businesses in any location and category. Extracts name, phone number, address, website, rating, and reviews. Then visits each business website to find email addresses and social media profiles. Saves everything to CSV, Excel, and JSON files. Supports running multiple browsers at the same time to scrape faster.

## Features

- Search by category and location
- Automatic GPS coordinate lookup from location names
- Scrape emails from business websites
- Find social media profiles (Facebook, Instagram, Twitter, LinkedIn, YouTube, TikTok)
- Run multiple browsers in parallel for faster results
- Duplicate detection across multiple runs and browsers
- First occurrence of each lead is always kept, duplicates are skipped
- Split output into separate folders per category or region
- Each split folder has its own duplicate detection
- Auto-save progress every 30 leads
- Config file support to save and reuse your settings
- Export to CSV, Excel, and JSON

## Requirements

- Python 3.7 or higher
- Google Chrome or Chromium browser
- ChromeDriver (matching your Chrome version)

## Installation

### Clone the repository

```bash
git clone https://github.com/qurixion/leapy.git
cd leapy
```

### Install Python packages

```bash
pip install selenium webdriver-manager beautifulsoup4 requests pandas openpyxl lxml
```

### Install Chrome

**Fedora:**
```bash
sudo dnf install chromium chromium-headless chromedriver
```

**Ubuntu/Debian:**
```bash
sudo apt install chromium-browser chromium-chromedriver
```

**Windows:**
Download and install from https://google.com/chrome

**macOS:**
```bash
brew install --cask google-chrome
```
## Usage

### Interactive Mode

Run without any arguments and follow the prompts:

```bash
python leapy.py
```

The script will show a menu:

```
  1. Start new run
  2. Load settings from config file
  3. Generate a sample config file
```

If you choose option 1 it will ask you for:
1. Search mode (GPS or text-based)
2. Number of parallel browsers
3. Categories to search
4. Regions to search
5. Number of leads per search
6. Working directory for output files
7. Output filename
8. Which formats to save (CSV, Excel, JSON)
9. Whether to split into separate folders per category or region
10. Whether to save settings to a config file for next time

### Config File Mode

Generate a sample config file:

```bash
python leapy.py --template
```

Edit `config.txt` with your settings then run:

```bash
python leapy.py --config config.txt
```

### Command Line Mode

```bash
python leapy.py -c "restaurant" -r "Paris France" -w ./output -n paris_leads
```

---

## Command Line Flags

| Flag | Description | Example |
|------|-------------|---------|
| `--config` | Load settings from config file | `--config config.txt` |
| `--template` | Generate a sample config file | `--template` |
| `-c` | Categories to search | `-c "bar,cafe,restaurant"` |
| `-r` | Regions to search | `-r "Paris France,London UK"` |
| `-w` | Working directory for output files | `-w /home/user/leads` |
| `-n` | Output filename without extension | `-n paris_leads` |
| `-l` | Leads per search (default 20) | `-l 50` |
| `-b` | Number of parallel browsers | `-b 2` |
| `--no-gps` | Use text search instead of GPS | `--no-gps` |
| `--no-csv` | Skip CSV output | `--no-csv` |
| `--no-xlsx` | Skip Excel output | `--no-xlsx` |
| `--no-json` | Skip JSON output | `--no-json` |
| `--split-category` | Save separate folder per category | `--split-category` |
| `--split-region` | Save separate folder per region | `--split-region` |

---

## Config File Format

Generate a template with `python leapy.py --template` or create `config.txt` manually:

```
# leapy Config File
# Lines starting with # are comments

# ── SEARCH ──────────────────────────────────
categories = bar, cafe, restaurant
regions = Paris France, Lyon France
leads_per_search = 20
use_gps = true
num_browsers = 1

# ── WORKING DIRECTORY ────────────────────────
working_directory = /home/user/leads

# ── OUTPUT ──────────────────────────────────
output_name = leads
save_csv  = true
save_xlsx = true
save_json = true

# ── SPLIT FILES ──────────────────────────────
split_by_category = false
split_by_region   = false
```

---

## How Location Search Works

When you enter a location name, the script automatically finds its GPS coordinates using OpenStreetMap for free with no API key needed:

```
Paris France   →  48.8566, 2.3522
Tokyo Japan    →  35.6828, 139.7594
New York NY    →  40.7128, -74.0060
London UK      →  51.5074, -0.1278
```

The script then searches Google Maps at those exact coordinates. This gives accurate results because Google shows businesses actually located there, not businesses that just have the city name in their business name.

Use `--no-gps` flag for faster but less accurate text-based search.

---

## Output Files

### Combined File

A combined file containing all leads is always saved in the working directory root:

```
working_directory/
├── leads.csv
├── leads.xlsx
└── leads.json
```

### Split by Category

When split by category is enabled, each category gets its own subfolder with its own file:

```
working_directory/
├── leads.csv              (combined, always created)
├── bar/
│   ├── leads_bar.csv
│   ├── leads_bar.xlsx
│   └── leads_bar.json
├── cafe/
│   ├── leads_cafe.csv
│   └── leads_cafe.xlsx
└── restaurant/
    └── leads_restaurant.csv
```

### Split by Region

When split by region is enabled, each region gets its own subfolder:

```
working_directory/
├── leads.csv              (combined, always created)
├── Paris_France/
│   └── leads_Paris_France.csv
└── London_UK/
    └── leads_London_UK.csv
```

### Split by Both

When both are enabled, each category and region combination gets its own subfolder:

```
working_directory/
├── leads.csv              (combined, always created)
├── bar_Paris_France/
│   └── leads_bar_Paris_France.csv
├── cafe_Paris_France/
│   └── leads_cafe_Paris_France.csv
├── bar_London_UK/
│   └── leads_bar_London_UK.csv
└── cafe_London_UK/
    └── leads_cafe_London_UK.csv
```

If the output file already exists from a previous run, new leads are appended without duplicates.

---

## Duplicate Detection

The script has three layers of duplicate detection:

**Layer 1: During scraping**
All browsers share one duplicate checker. The first time a business is found it is kept. Any subsequent browser that finds the same business will skip it. This works across all parallel browsers.

**Layer 2: Against existing files**
When you run the script again with the same output file, leads already in that file are skipped. This means you can run the script multiple times and never get duplicates.

**Layer 3: Per split file**
Each split folder has its own duplicate detection separate from the others. So the same business can appear in the bar folder and the cafe folder if it fits both categories, but it will never appear twice within the same folder.

---

## Data Collected

For each business the script collects:

| Field | Description |
|-------|-------------|
| name | Business name |
| phone | Phone number |
| address | Full address |
| website | Website URL |
| rating | Google rating |
| reviews | Number of reviews |
| category | Business type |
| emails | Emails found on website |
| facebook | Facebook page URL |
| instagram | Instagram profile URL |
| twitter | Twitter/X profile URL |
| linkedin | LinkedIn page URL |
| youtube | YouTube channel URL |
| tiktok | TikTok profile URL |
| search_region | Region that was searched |
| search_category | Category that was searched |
| scraped_at | Date and time of scraping |

---

## Parallel Browsers

The script supports running multiple Chrome browsers at the same time to speed up scraping. Each browser handles a different search so they work in parallel.

```bash
python leapy.py -r "Paris,Lyon,Marseille" -b 3 -w ./output
```

**RAM requirements:**
- 1 browser: ~1GB RAM
- 2 browsers: ~2GB RAM
- 3 browsers: ~3GB RAM

All browsers share one duplicate detection system so the same business is never collected twice even when multiple browsers are running at the same time.

---

## Using Text Files

You can put categories and regions in text files with one item per line:

**categories.txt**
```
restaurant
cafe
bar
hotel
gym
```

**regions.txt**
```
Paris France
London UK
Berlin Germany
Tokyo Japan
```

Then use them with:
```bash
python leapy.py -c categories.txt -r regions.txt -w ./output
```

---

## Examples

Collect 50 restaurants in London:
```bash
python leapy.py -c "restaurant" -r "London UK" -w ./output -n london -l 50
```

Multiple categories and regions:
```bash
python leapy.py -c "bar,cafe" -r "Paris France,Lyon France" -w ./output -n france
```

Use 2 browsers for faster scraping:
```bash
python leapy.py -c "gym" -r "New York NY" -w ./output -b 2
```

Split into separate folders per category:
```bash
python leapy.py -c "bar,cafe,restaurant" -r "Paris France" -w ./output --split-category
```

Split into separate folders per region:
```bash
python leapy.py -c "restaurant" -r "Paris France,London UK,Berlin Germany" -w ./output --split-region
```

Split by both category and region:
```bash
python leapy.py -c "bar,cafe" -r "Paris France,London UK" -w ./output --split-category --split-region
```

Run from config file:
```bash
python leapy.py --config config.txt
```

Text search without GPS:
```bash
python leapy.py -r "Tokyo" -w ./output --no-gps
```

Save only CSV, skip Excel and JSON:
```bash
python leapy.py -r "Madrid Spain" -w ./output --no-xlsx --no-json
```

---

## Project Structure

```
leapy/
├── leapy.py    (main script)
├── config.txt             (your config file, generated with --template)
├── README.md              (this file)
└── leapy.log   (error log, created when you run the script)
```

---

## Troubleshooting

**Browser fails to start:**
Make sure Chrome or Chromium is installed. On Windows see the Windows Setup section above.

**Geocoding fails:**
Add the country name to make it more specific. Use `Paris France` instead of just `Paris`.

**No results found:**
Try a different category or a larger city. Some areas have fewer businesses on Google Maps.

**Script interrupted:**
Check the working directory for partial results. The script saves progress every 30 leads so you should not lose much data.

**Duplicate chromedriver error on Windows:**
Make sure you only have one version of ChromeDriver installed and that it matches your Chrome version.

**Excel save fails:**
Install openpyxl:
```bash
pip install openpyxl
```

**Split folders not created:**
Make sure you enabled split by category or split by region during setup or in the config file.

---

## Notes

- The script adds delays between requests to avoid being blocked by Google
- GPS geocoding uses free OpenStreetMap service with a 1 second delay between requests as required
- All errors are logged to `leapy.log` in the folder where you run the script
- The combined file is always created regardless of split settings

## License

MIT License
