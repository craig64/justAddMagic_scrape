#!/usr/bin/env python3
"""
Just Add Magic - Recipe Scraper

Outputs:
 - JustAddMagic_Cookbook.docx   (alphabetical cookbook; image URLs shown)
 - JustAddMagic_Recipes.json    (raw structured data)
 - JustAddMagic_MagicalSpices.docx (index of magical spices)

Usage:
    python jam_scraper.py [--delay 1.0] [--start-index 0]

Notes:
 - Be polite: the script sleeps between requests (default delay 1.0s).
 - If interrupted, re-run with --start-index to resume working through the recipe list.
"""

import re
import time
import json
import argparse
from pathlib import Path
from typing import List, Dict, Optional
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt
from tqdm import tqdm

BASE_CATEGORY_URL = "https://justaddmagic.fandom.com/wiki/Category:Recipes"
HEADERS = {
    "User-Agent": "JustAddMagic-Scraper/1.0 (+https://example.com) - for personal use"
}

# --------------------------
# Utilities
# --------------------------

def get_soup(url: str, session: requests.Session, timeout=20) -> Optional[BeautifulSoup]:
    try:
        r = session.get(url, headers=HEADERS, timeout=timeout)
        r.raise_for_status()
        return BeautifulSoup(r.text, "html.parser")
    except Exception as e:
        print(f"[ERROR] Failed to fetch {url}: {e}")
        return None

def extract_recipe_links_from_category(soup: BeautifulSoup) -> List[str]:
    """
    Fandom category pages usually put member links in 'div.category-page__members' or '#mw-pages'.
    This function finds all recipe page links on a category page soup.
    """
    links = []
    # try category-page__members first
    members = soup.select("div.category-page__members a.category-page__member-link")
    if members:
        for a in members:
            href = a.get("href")
            if href:
                links.append(requests.compat.urljoin("https://justaddmagic.fandom.com", href))
        return links

    # fallback: look for anchors under #mw-pages
    mw_pages = soup.select("#mw-pages a")
    for a in mw_pages:
        href = a.get("href")
        if href and "/wiki/" in href:
            full = requests.compat.urljoin("https://justaddmagic.fandom.com", href)
            links.append(full)
    return links

def find_next_page_link(soup: BeautifulSoup) -> Optional[str]:
    # fandom category pagination uses .category-page__pagination .category-page__pagination-next
    nxt = soup.select_one("a.category-page__pagination-next")
    if nxt and nxt.get("href"):
        return requests.compat.urljoin("https://justaddmagic.fandom.com", nxt["href"])
    # fallback: look for link text 'next page' or 'next'
    for a in soup.find_all("a"):
        if a.text.strip().lower() in ("next page", "next"):
            if a.get("href"):
                return requests.compat.urljoin("https://justaddmagic.fandom.com", a["href"])
    return None

def get_all_category_recipe_urls(session: requests.Session, delay: float = 1.0) -> List[str]:
    """
    Crawl category pages (handles simple pagination) and returns unique recipe urls.
    """
    urls = []
    next_url = BASE_CATEGORY_URL
    while next_url:
        print(f"[INFO] Fetching category page: {next_url}")
        soup = get_soup(next_url, session)
        if not soup:
            break
        page_links = extract_recipe_links_from_category(soup)
        print(f"[INFO] Found {len(page_links)} recipe links on this page.")
        for u in page_links:
            if u not in urls:
                urls.append(u)
        next_url = find_next_page_link(soup)
        if next_url:
            print(f"[INFO] Next category page -> {next_url}")
            time.sleep(delay)
    print(f"[INFO] Total recipe links collected: {len(urls)}")
    return urls

# --------------------------
# Recipe page parsing
# --------------------------

def get_title(soup: BeautifulSoup) -> str:
    # Many fandom pages use <h1 id="firstHeading"> or .page-header__title
    t = soup.select_one("#firstHeading") or soup.select_one(".page-header__title")
    if t:
        return t.get_text(strip=True)
    # fallback to title tag
    if soup.title:
        return soup.title.get_text(strip=True)
    return "Unknown Title"

def get_main_image_url(soup: BeautifulSoup) -> Optional[str]:
    # try common fandom selectors: figure.pi-image img, .wikia-article or infobox
    candidates = [
        soup.select_one("figure.pi-image img"),
        soup.select_one(".wds-media img"),
        soup.select_one(".portable-infobox img"),
        soup.select_one("table.infobox img"),
    ]
    for c in candidates:
        if c and c.get("src"):
            src = c.get("src")
            # some images use data-src
            if src.startswith("//"):
                src = "https:" + src
            return src
    # fallback: look for og:image meta
    meta = soup.find("meta", property="og:image")
    if meta and meta.get("content"):
        return meta["content"]
    return None

def get_section_contents(soup: BeautifulSoup, heading_names: List[str]) -> List[str]:
    """
    Find the first heading that matches one of heading_names (case-insensitive),
    then return the text content of the following sibling tags until the next heading of same level.
    Returns lines (list) preserving list items if present.
    """
    # headings can be h2/h3; search for them
    heading = None
    for htag in soup.find_all(re.compile("^h[1-6]$")):
        txt = htag.get_text(separator=" ", strip=True).strip().lower()
        for name in heading_names:
            if name.lower() in txt:
                heading = htag
                break
        if heading:
            break
    if not heading:
        # fallback: search for bold/strong label "Ingredients" in paragraphs
        for strong in soup.find_all(("b", "strong")):
            if any(name.lower() in strong.get_text(strip=True).lower() for name in heading_names):
                # try to get parent
                heading = strong.parent
                break
    if not heading:
        return []
    # collect siblings until next heading of the same level or higher
    results = []
    for sib in heading.find_next_siblings():
        if sib.name and re.match("^h[1-6]$", sib.name):
            break
        # if list, extract li items
        if sib.name in ("ul", "ol"):
            for li in sib.find_all("li"):
                txt = " ".join(li.stripped_strings)
                if txt:
                    results.append(txt)
        else:
            txt = " ".join(sib.stripped_strings)
            if txt:
                results.append(txt)
    return results

MAGICAL_SPICE_PATTERN = re.compile(r"\b([A-Z][a-z]+)\s+([a-z][a-z0-9'\- ]+)\b")
# Explanation: looks for "CapitalizedFamily" followed by lowercase base spice text (e.g. "Taurian salt", "Cedonian cinnamon")

def extract_magical_spices_from_ingredient_line(line: str) -> List[Dict]:
    """
    From a single ingredient line, extract all Family + base_spice occurrences.
    Returns list of dicts: {'family':..., 'magical':..., 'base':...}
    """
    out = []
    # search for multiple occurrences in the line
    for m in MAGICAL_SPICE_PATTERN.finditer(line):
        family = m.group(1).strip()
        base = m.group(2).strip()
        magical = f"{family} {base}"
        # normalize base: remove trailing commas/parentheses etc
        base = re.sub(r"[,\.\(\)]", "", base).strip()
        out.append({
            "family": family,
            "magical": magical,
            "base": base
        })
    return out

def parse_recipe_page(url: str, session: requests.Session, delay: float = 0.5) -> Optional[Dict]:
    soup = get_soup(url, session)
    if not soup:
        return None
    title = get_title(soup)
    image_url = get_main_image_url(soup)
    # Ingredients: check for common headings
    ingredients = get_section_contents(soup, ["Ingredients", "Ingredient"])
    # Directions / Method / Instructions / Preparation
    directions = get_section_contents(soup, ["Directions", "Method", "Instructions", "Directions / Method", "Preparation", "Directions / Instructions"])
    # riddle / notes sections
    riddle = get_section_contents(soup, ["Riddle", "Riddles", "Notes", "Notes / Trivia", "Effects", "Effects / Riddle"])
    # If page provides a recipe template table, attempt to parse it too (fallback)
    if not ingredients:
        # look for lists or paragraphs that could be labeled recipe-ingredients (fallback heuristics)
        for heading_text in ("Ingredients",):
            pass  # we already attempted; leave blank
    # Extract magical spices found inside ingredient lines
    magical_spices = []
    for line in ingredients:
        found = extract_magical_spices_from_ingredient_line(line)
        for f in found:
            magical_spices.append(f)
    # canonicalize
    recipe = {
        "title": title,
        "url": url,
        "image_url": image_url,
        "ingredients": ingredients,
        "directions": directions,
        "riddle": riddle,
        "magical_spices": magical_spices
    }
    time.sleep(delay)
    return recipe

# --------------------------
# Output builders
# --------------------------

def save_json(data, path: Path):
    with path.open("w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)

def build_cookbook_docx(recipes: List[Dict], path: Path):
    doc = Document()
    # default font
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    doc.add_heading("Just Add Magic — Cookbook", level=1)
    doc.add_paragraph("Source: https://justaddmagic.fandom.com/wiki/Category:Recipes")
    doc.add_paragraph("Note: Image URLs are included below each recipe. To embed images run the alternate script that downloads them.")
    # TOC placeholder - user can generate actual TOC in Word/Google Docs from headings
    doc.add_paragraph("Table of Contents (HEADINGS are included below; open this file in Word and 'Update Table' to generate a clickable TOC).")
    doc.add_paragraph("")  # blank line

    # Sort recipes alphabetically by title
    recipes_sorted = sorted(recipes, key=lambda r: (r.get("title") or "").lower())

    for r in recipes_sorted:
        # Use heading so Word can detect and create TOC
        doc.add_heading(r.get("title", "Untitled Recipe"), level=1)
        doc.add_paragraph(f"Source: {r.get('url')}")
        if r.get("image_url"):
            doc.add_paragraph(f"Image URL: {r.get('image_url')}")
        if r.get("magical_spices"):
            ms_text = ", ".join(sorted({m['magical'] for m in r['magical_spices']}))
            doc.add_paragraph(f"Magical spices found: {ms_text}")
        if r.get("ingredients"):
            doc.add_paragraph("Ingredients:")
            for ing in r["ingredients"]:
                p = doc.add_paragraph(f"• {ing}")
                p.style = doc.styles['Normal']
        if r.get("directions"):
            doc.add_paragraph("Directions:")
            for i, step in enumerate(r["directions"], 1):
                p = doc.add_paragraph(f"{i}. {step}")
                p.style = doc.styles['Normal']
        if r.get("riddle"):
            doc.add_paragraph("Notes / Riddle / Effects:")
            for line in r["riddle"]:
                doc.add_paragraph(line)
        # page break between recipes
        doc.add_page_break()

    doc.save(str(path))
    print(f"[INFO] Cookbook saved to {path}")

def build_spices_index_docx(spice_index: Dict[str, Dict], path: Path):
    """
    spice_index: keyed by magical spice name (e.g. "Taurian salt"), value:
      { 'family':..., 'base':..., 'first_appearance':..., 'recipes': [list] }
    """
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)

    doc.add_heading("Magical Spices — Index", level=1)
    doc.add_paragraph("Generated from all recipes scraped from: https://justaddmagic.fandom.com/wiki/Category:Recipes")
    doc.add_paragraph("Columns: Family | Magical Spice | Base Spice | First Appearance | All Recipes")

    # Build a table with header
    table = doc.add_table(rows=1, cols=5)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Family"
    hdr_cells[1].text = "Magical Spice"
    hdr_cells[2].text = "Base Spice"
    hdr_cells[3].text = "First Appearance"
    hdr_cells[4].text = "Recipes (comma separated)"

    # sort by family then base
    items = sorted(spice_index.items(), key=lambda kv: (kv[1]['family'], kv[1]['base']))
    for magical, info in items:
        row = table.add_row().cells
        row[0].text = info.get('family', '')
        row[1].text = magical
        row[2].text = info.get('base', '')
        row[3].text = info.get('first_appearance', '')
        row[4].text = ", ".join(info.get('recipes', []))

    doc.save(str(path))
    print(f"[INFO] Spice index saved to {path}")

# --------------------------
# Main
# --------------------------

def main():
    parser = argparse.ArgumentParser(description="Scrape JustAddMagic fandom recipes and produce a cookbook + spice index.")
    parser.add_argument("--delay", type=float, default=1.0, help="Delay between requests (seconds). Be polite.")
    parser.add_argument("--start-index", type=int, default=0, help="If resuming, the index in the recipe URL list to start from (0-based).")
    parser.add_argument("--max", type=int, default=0, help="Optional: maximum number of recipes to scrape (0 means all).")
    args = parser.parse_args()

    out_dir = Path.cwd() / "jam_output"
    out_dir.mkdir(exist_ok=True)

    session = requests.Session()

    # 1) Gather all recipe URLs from category
    all_urls = get_all_category_recipe_urls(session, delay=args.delay)
    if args.max and args.max > 0:
        all_urls = all_urls[:args.max]

    recipes_data = []
    spice_index = {}  # key=magical spice string -> info dict

    # Resume semantics: if recipes.json exists, load it and resume
    recipes_json_path = out_dir / "JustAddMagic_Recipes.json"
    if recipes_json_path.exists():
        print("[INFO] Existing recipes JSON found; loading to resume.")
        with recipes_json_path.open("r", encoding="utf-8") as fh:
            recipes_data = json.load(fh)
    start_idx = args.start_index if args.start_index >= 0 else len(recipes_data)
    print(f"[INFO] Starting from index {start_idx} (of {len(all_urls)} total URLs).")

    # If recipes_data already has items, we will skip URLs already processed that match
    processed_urls = set(r.get("url") for r in recipes_data)

    # iterate through recipe URLs
    for idx in tqdm(range(start_idx, len(all_urls)), desc="recipes"):
        url = all_urls[idx]
        if url in processed_urls:
            print(f"[INFO] Skipping already-processed URL: {url}")
            continue
        print(f"[INFO] Scraping recipe {idx+1}/{len(all_urls)}: {url}")
        recipe = parse_recipe_page(url, session, delay=args.delay)
        if not recipe:
            print(f"[WARN] Skipping {url} due to parse error.")
            continue

        recipes_data.append(recipe)
        processed_urls.add(url)

        # update spice_index
        for ms in recipe.get("magical_spices", []):
            magical = ms["magical"]
            family = ms["family"]
            base = ms["base"]
            if magical not in spice_index:
                spice_index[magical] = {
                    "family": family,
                    "base": base,
                    "first_appearance": recipe.get("title", ""),
                    "recipes": [recipe.get("title", "")]
                }
            else:
                spice_index[magical]["recipes"].append(recipe.get("title", ""))

        # periodically save progress (every recipe)
        save_json(recipes_data, recipes_json_path)

    # Deduplicate recipe lists in spice_index and sort
    for info in spice_index.values():
        info["recipes"] = sorted(list(set(info.get("recipes", []))))

    # Save final JSON
    save_json(recipes_data, recipes_json_path)
    print(f"[INFO] Saved recipes JSON to {recipes_json_path}")

    # Build DOCX cookbook and spice index docx
    cookbook_path = out_dir / "JustAddMagic_Cookbook.docx"
    spices_docx_path = out_dir / "JustAddMagic_MagicalSpices.docx"
    build_cookbook_docx(recipes_data, cookbook_path)
    build_spices_index_docx(spice_index, spices_docx_path)

    print("[DONE] All files written to:", out_dir)

if __name__ == "__main__":
    main()
