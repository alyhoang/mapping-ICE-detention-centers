#!/usr/bin/env python
"""
ICE Detention Data Scraper -> AGOL

Automated pipeline that:
  1. Scrapes ice.gov for the latest detention stats Excel file
  2. Downloads it
  3. Formats it into MAIN FORMAT (with derived totals, full address; no phone numbers)
  4. Geocodes each facility using the Census Bureau's free batch geocoder
  5. Exports an ArcGIS-ready CSV (with latitude/longitude columns)
  6. Publishes a feature layer to ArcGIS Online
  7. Adds/replaces that layer in an existing Web Map

Scheduling is handled externally via GitHub Actions (.github/workflows/scheduler.yml).

Credits:
    Steps 1-2: https://github.com/lockdown-systems/icewatch/blob/main/src/icewatch/ice_detention_scraper.py

Usage:
    python ice_scraper.py                  # full run (scrape -> format -> geocode -> publish)
    python ice_scraper.py --dry-run        # local only; skips all AGOL calls
    python ice_scraper.py --test-agol      # one full pass including AGOL, then exit
    python ice_scraper.py --output-dir ./output

Required environment variables:
    AGOL_USERNAME     - your ArcGIS Online username
    AGOL_PASSWORD     - your ArcGIS Online password
    AGOL_MAP_ITEM_ID  - item ID of the Web Map to update
    AGOL_FOLDER       - content folder to publish into (optional; default: root)
"""

# library ──────────────────────────────────────────────────────────
import argparse
import io
import json
import logging
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin, urlparse
import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from requests.exceptions import RequestException
from arcgis.gis import GIS
from arcgis.features import FeatureLayer


# Logging setup ─────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
log = logging.getLogger(__name__)


# CONSTANTS ─────────────────────────────────────────────────────────────

ICE_PAGE_URL = "https://www.ice.gov/detain/detention-management"
FALLBACK_URL = "https://www.ice.gov/doclib/detention/FY26_detentionStats_04092026.xlsx"
SEARCH_KEYWORDS = ["detention", "statistics", "FY26", "YTD", "xlsx", "detentionStats", "FY"]

BROWSER_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
}

EXCEL_HEADERS = {
    **BROWSER_HEADERS,
    "Accept": (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,"
        "application/vnd.ms-excel,application/octet-stream"
    ),
}

OUTPUT_COLUMNS = [
    "Name", "Address", "City", "State", "Zip",
    "Type Detailed", "FY25 ALOS", "FY26 ALOS",
    "notes", "last_updated",
    "Total Female Detainment", "Total Male Detainment",
    "Total Detainment", "Total Non-Criminal",
    "Full Address",
]

ARCGIS_COLUMNS = OUTPUT_COLUMNS


# HELPER FUNCTIONS ──────────────────────────────────────────────────────────
def is_valid_date(year: int, month: int, day: int) -> bool:
    try:
        dt = datetime(year=year, month=month, day=day)
        return dt > datetime(year=2025, month=1, day=1)
    except ValueError:
        return False


def _coerce_float(val) -> float:
    try:
        return float(val) if val is not None and str(val).strip() not in ("", "nan") else 0.0
    except (ValueError, TypeError):
        return 0.0


def _zip_str(val) -> str:
    s = str(val).replace(".0", "").strip()
    if s in ("", "nan", "None"):
        return ""
    return s.zfill(5)


def _find_sheet(wb_sheets: list[str]) -> str:
    for s in wb_sheets:
        if "facilit" in s.lower():
            return s
    raise ValueError(f"No 'Facilities' sheet found. Available: {wb_sheets}")


# STEP 1 — Scrape stable link for the latest detention stats Excel link
# ─────────────────────────────────────────────────────────────────────────
def find_detention_stats_link(base_url: str = ICE_PAGE_URL) -> str | None:
    log.info(f"Scraping: {base_url}")
    try:
        resp = requests.get(base_url, headers=BROWSER_HEADERS, timeout=30)
        resp.raise_for_status()

        soup = BeautifulSoup(resp.content, "html.parser")
        found = []

        for link in soup.find_all("a", href=True):
            href = str(link.get("href", "")).lower()
            text = link.get_text().lower()
            score = sum(1 for kw in SEARCH_KEYWORDS if kw.lower() in href or kw.lower() in text)
            if score:
                full_url = urljoin(base_url, str(link["href"]))
                found.append({
                    "full_url":  full_url,
                    "link_text": link.get_text().strip(),
                    "score":     score,
                })
                log.info(f"  Found potential link: {link.get_text().strip()} -> {full_url}")

        if not found:
            log.warning("No relevant links found on page")
            return None

        found.sort(key=lambda x: (x["score"], ".xlsx" in x["full_url"].lower()), reverse=True)
        best = found[0]
        log.info(f"Selected: {best['link_text']!r} -> {best['full_url']}")
        return best["full_url"]

    except RequestException as e:
        log.error(f"Scrape failed: {e}")
        return None
    except Exception as e:
        log.error(f"Unexpected error while scraping: {e}")
        return None


# STEP 2 — Download the Excel file
# ─────────────────────────────────────────────────────────────────────────
def _extract_date(url: str) -> str | None:
    try:
        filename = os.path.basename(urlparse(url).path)
        date_patterns = [
            r"(?P<month>\d{2})(?P<day>\d{2})(?P<year>\d{4})\.xlsx",
            r"(?P<year>\d{4})(?P<month>\d{2})(?P<day>\d{2})\.xlsx",
            r"(?P<month>\d{2})(?P<day>\d{2})(?P<year>\d{2})\.xlsx",
        ]
        for pattern in date_patterns:
            if m := re.search(pattern, filename):
                year, month, day = int(m.group("year")), int(m.group("month")), int(m.group("day"))
                if year < 100:
                    year += 2000
                if is_valid_date(year, month, day):
                    return f"{year}-{month:02}-{day:02}"
        log.warning(f"Could not extract date from filename: {filename}")
        return None
    except Exception as e:
        log.error(f"Error extracting date: {e}")
        return None


def download_xlsx(url: str, output_dir: str) -> tuple[str | None, str | None]:
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    source_date = _extract_date(url)
    original_filename = os.path.basename(urlparse(url).path)

    filename = (
        original_filename
        if original_filename and original_filename.endswith(".xlsx")
        else f"ice_detention_stats_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )
    filepath = os.path.join(output_dir, filename)

    log.info(f"Downloading: {url}")
    try:
        resp = requests.get(url, headers=EXCEL_HEADERS, stream=True, timeout=60)
        resp.raise_for_status()
        with open(filepath, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        log.info(f"Saved to: {filepath}  ({os.path.getsize(filepath) / 1024:.1f} KB)")
        return filepath, source_date
    except RequestException as e:
        log.error(f"Download failed: {e}")
        return None, None
    except Exception as e:
        log.error(f"Unexpected error: {e}")
        return None, None


# STEP 3 — Parse the ICE source Excel into MAIN FORMAT
# ─────────────────────────────────────────────────────────────────────────
def parse_source_excel(filepath: str, source_date: str | None) -> pd.DataFrame | None:
    log.info(f"Parsing source file: {filepath}")
    try:
        xl = pd.ExcelFile(filepath)
        sheet = _find_sheet(xl.sheet_names)
        log.info(f"Using sheet: {sheet!r}")

        df_raw = None
        for hrow in range(6, 15):
            try:
                candidate = pd.read_excel(filepath, sheet_name=sheet, header=hrow)
                if "Name" in candidate.columns:
                    df_raw = candidate
                    log.info(f"Found header row at index {hrow}")
                    break
            except Exception:
                continue

        if df_raw is None:
            log.error("Could not locate header row with 'Name' column")
            return None

        df_raw = df_raw.dropna(how="all").reset_index(drop=True)
        cols = df_raw.columns.tolist()
        log.info(f"Source columns: {cols}")

        def find_col(*candidates) -> str | None:
            for c in candidates:
                if c in cols:
                    return c
                for actual in cols:
                    if str(actual).strip().lower() == c.lower():
                        return actual
            return None

        col_name      = find_col("Name")
        col_address   = find_col("Address")
        col_city      = find_col("City")
        col_state     = find_col("State")
        col_zip       = find_col("Zip", "ZIP", "Zip Code")
        col_type      = find_col("Type Detailed", "Type", "Facility Type", "Detainee Type")
        col_fy25_alos = find_col("FY25 ALOS", "FY 25 ALOS", "ALOS FY25")
        col_fy26_alos = find_col("FY26 ALOS", "FY 26 ALOS", "ALOS FY26", "ALOS")
        col_male_crim = find_col("Male Crim", "M Crim")
        col_male_nc   = find_col("Male Non-Crim", "M Non-Crim", "Male NonCrim")
        col_fem_crim  = find_col("Female Crim", "F Crim")
        col_fem_nc    = find_col("Female Non-Crim", "F Non-Crim", "Female NonCrim")

        missing = [n for n, c in [
            ("Name", col_name), ("Address", col_address),
            ("City", col_city), ("State", col_state), ("Zip", col_zip),
        ] if c is None]
        if missing:
            log.error(f"Missing required columns: {missing}")
            return None

        rows = []
        update_date = source_date or datetime.now().strftime("%Y-%m-%d")

        for _, r in df_raw.iterrows():
            name = str(r[col_name]).strip() if col_name and r[col_name] is not None else ""
            if not name or name.lower() in ("nan", "name", "total", "subtotal"):
                continue

            addr  = str(r[col_address]).strip() if col_address else ""
            city  = str(r[col_city]).strip()    if col_city    else ""
            state = str(r[col_state]).strip()   if col_state   else ""
            zip_  = _zip_str(r[col_zip])        if col_zip     else ""

            ftype     = str(r[col_type]).strip()             if col_type     else ""
            fy25_alos = _coerce_float(r[col_fy25_alos])      if col_fy25_alos else None
            fy26_alos = _coerce_float(r[col_fy26_alos])      if col_fy26_alos else None

            male_crim = _coerce_float(r[col_male_crim]) if col_male_crim else 0.0
            male_nc   = _coerce_float(r[col_male_nc])   if col_male_nc   else 0.0
            fem_crim  = _coerce_float(r[col_fem_crim])  if col_fem_crim  else 0.0
            fem_nc    = _coerce_float(r[col_fem_nc])    if col_fem_nc    else 0.0

            total_female   = round(fem_crim + fem_nc)
            total_male     = round(male_crim + male_nc)
            total_det      = round(total_female + total_male)
            total_non_crim = round(male_nc + fem_nc)

            addr  = addr  if addr.lower()  not in ("nan", "none") else ""
            city  = city  if city.lower()  not in ("nan", "none") else ""
            state = state if state.lower() not in ("nan", "none") else ""

            full_address = ", ".join(p for p in [addr, city, state, zip_] if p)

            rows.append({
                "Name":                    name,
                "Address":                 addr,
                "City":                    city,
                "State":                   state,
                "Zip":                     zip_,
                "Type Detailed":           ftype if ftype.lower() not in ("nan", "") else "",
                "FY25 ALOS":               fy25_alos,
                "FY26 ALOS":               fy26_alos,
                "notes":                   None,
                "last_updated":            update_date,
                "Total Female Detainment": total_female,
                "Total Male Detainment":   total_male,
                "Total Detainment":        total_det,
                "Total Non-Criminal":      total_non_crim,
                "Full Address":            full_address,
            })

        df_out = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
        log.info(f"Parsed {len(df_out)} facility rows")
        return df_out

    except Exception as e:
        log.error(f"Failed to parse source Excel: {e}", exc_info=True)
        return None


# STEP 4 — Write MAIN FORMAT .xlsx
# ─────────────────────────────────────────────────────────────────────────
def write_main_format(df: pd.DataFrame, out_path: str) -> bool:
    log.info(f"Writing MAIN FORMAT to: {out_path}")
    try:
        df[OUTPUT_COLUMNS].to_excel(out_path, index=False, sheet_name="ICE_Detention_Data")
        log.info(f"MAIN FORMAT saved: {out_path}")
        return True
    except Exception as e:
        log.error(f"Failed to write MAIN FORMAT: {e}", exc_info=True)
        return False


# STEP 5a — Address cleaning helpers
# ─────────────────────────────────────────────────────────────────────────
def clean_address(street: str, city: str) -> tuple[str, str]:
    """Normalise a street address and city before sending to the Census geocoder."""
    if not isinstance(street, str):
        return street, city

    s = street.strip()

    # Remove P.O. Box / PO Box entries — geocoder can't handle them
    if re.match(r"(?i)^p\.?\s*o\.?\s*box", s):
        s = ""

    # Remove trailing "P.O. BOX …" appended to a real street address
    s = re.sub(r"(?i)\s*p\.?\s*o\.?\s*box\s+\S+", "", s).strip()

    # Strip interior periods from abbreviations: "DR." -> "DR", "S.W." -> "SW"
    s = re.sub(r"(?<=[A-Z])\.(?=[A-Z])", "", s)   # mid-word dots  S.W. -> SW
    s = re.sub(r"([A-Za-z]+)\.", r"", s)            # trailing dots  DR. -> DR

    # Expand common city abbreviations
    s = re.sub(r"HWY",      "Highway",  s, flags=re.I)
    s = re.sub(r"RT\.?\s+",   "Route ",   s, flags=re.I)
    s = re.sub(r"RTE\.?\s+",  "Route ",   s, flags=re.I)
    s = re.sub(r"US\s+(\d+)", r"US Highway ", s, flags=re.I)

    # Strip "ROUTE N, BOX N" style rural addresses — geocoder can't resolve them
    s = re.sub(r"(?i)^route\s+\d+,?\s+box\s+\S+$", "", s).strip()

    # Clean city
    c = city.strip() if isinstance(city, str) else city
    c = re.sub(r"(?i)^FT\.\s*",  "Fort ",   c)   # FT. or FT -> Fort
    c = re.sub(r"(?i)^ST\.\s*",  "Saint ",  c)
    c = re.sub(r"(?i)^MT\.\s*",  "Mount ",  c)

    # Collapse extra whitespace
    s = " ".join(s.split())
    c = " ".join(c.split())

    return s, c


# STEP 5b — Geocode addresses (Census Bureau free batch API)
# ─────────────────────────────────────────────────────────────────────────
def geocode_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    log.info("Geocoding via Census Bureau API...")
    df = df.copy()
    df["latitude"]  = None
    df["longitude"] = None

    # Clean addresses before sending to geocoder
    cleaned = [clean_address(s, c) for s, c in zip(df["Address"], df["City"])]
    streets = [x[0] for x in cleaned]
    cities  = [x[1] for x in cleaned]

    geocode_input = pd.DataFrame({
        "id":     range(len(df)),
        "street": streets,
        "city":   cities,
        "state":  df["State"].values,
        "zip":    df["Zip"].values,
    })

    chunk_size = 1000
    for start in range(0, len(geocode_input), chunk_size):
        chunk = geocode_input.iloc[start:start + chunk_size]
        csv_buf = chunk.to_csv(index=False, header=False)
        try:
            resp = requests.post(
                "https://geocoding.geo.census.gov/geocoder/locations/addressbatch",
                files={"addressFile": ("addresses.csv", csv_buf, "text/csv")},
                data={"benchmark": "Public_AR_Current"},
                timeout=120,
            )
            resp.raise_for_status()

            result = pd.read_csv(
                io.StringIO(resp.text),
                header=None,
                names=["id", "input_addr", "match", "matchtype", "outaddr", "coords", "tiger", "side"],
                dtype=str,
            )
            matched = result[result["match"].str.strip().str.lower() == "match"]
            for _, row in matched.iterrows():
                try:
                    idx = int(row["id"])
                    lon, lat = row["coords"].split(",")
                    df.at[idx, "latitude"]  = float(lat.strip())
                    df.at[idx, "longitude"] = float(lon.strip())
                except Exception:
                    continue

            log.info(f"  Chunk {start}: {len(matched)}/{len(chunk)} matched")
        except Exception as e:
            log.warning(f"  Geocoding chunk {start} failed: {e} — skipping")

        time.sleep(0.5)

    total = df["latitude"].notna().sum()
    log.info(f"Geocoding complete: {total}/{len(df)} facilities located")
    return df


# STEP 5b — Export ArcGIS-ready CSV
# ─────────────────────────────────────────────────────────────────────────
def export_arcgis_csv(df: pd.DataFrame, out_path: str) -> bool:
    log.info(f"Writing ArcGIS CSV to: {out_path}")
    try:
        arcgis_df = df[ARCGIS_COLUMNS].copy()
        arcgis_df["Zip"] = arcgis_df["Zip"].astype(str).str.zfill(5).replace("00000", "")

        arcgis_df = geocode_dataframe(arcgis_df)
        arcgis_df = arcgis_df.fillna("")

        for col in ("Total Female Detainment", "Total Male Detainment",
                    "Total Detainment", "Total Non-Criminal"):
            arcgis_df[col] = pd.to_numeric(arcgis_df[col], errors="coerce").fillna(0).astype(int)

        for col in ("FY25 ALOS", "FY26 ALOS"):
            arcgis_df[col] = pd.to_numeric(arcgis_df[col], errors="coerce")

        arcgis_df.to_csv(out_path, index=False, encoding="utf-8")
        log.info(f"ArcGIS CSV saved: {out_path}  ({len(arcgis_df)} rows)")
        return True
    except Exception as e:
        log.error(f"Failed to write ArcGIS CSV: {e}", exc_info=True)
        return False


# STEP 6 — Launch new layer to ArcGIS Online
# ─────────────────────────────────────────────────────────────────────────
def publish_new_agol_layer(
    csv_path: str,
    run_date: str | None = None,
) -> tuple[str, str, "GIS"] | tuple[None, None, None]:
    username = os.environ.get("AGOL_USERNAME")
    password = os.environ.get("AGOL_PASSWORD")
    folder   = os.environ.get("AGOL_FOLDER", None)

    if not all([username, password]):
        log.warning(
            "AGOL publish skipped: set AGOL_USERNAME and AGOL_PASSWORD "
            "environment variables to enable."
        )
        return None, None, None

    date_label  = run_date or datetime.now().strftime("%Y-%m-%d")
    layer_title = f"ICE Detention Facilities {date_label}"
    safe_name   = re.sub(r"[^A-Za-z0-9_]", "_", layer_title)

    try:
        log.info("Connecting to ArcGIS Online...")
        gis = GIS("https://www.arcgis.com", username, password, verify_cert=True)

        log.info(f"Uploading CSV item: {layer_title!r}")
        for old_item in gis.content.search(f'title:"{layer_title}" type:CSV', max_items=5):
            if old_item.title == layer_title:
                log.info(f"Deleting existing CSV item: {old_item.id}")
                old_item.delete()

        for old_fs in gis.content.search(f'title:"{layer_title}"', max_items=10):
            if old_fs.type in ("Feature Service", "Feature Layer Collection"):
                log.info(f"Deleting old feature service: {old_fs.title} ({old_fs.id})")
                old_fs.delete()

        # Upload CSV via REST API directly to avoid gis.content.add() bug
        # in arcgis 2.4.x + Python 3.13 (_is_geoenabled AttributeError).
        token = gis._con.token
        portal_url = gis.url.rstrip("/")
        add_url = f"{portal_url}/sharing/rest/content/users/{username}"
        if folder:
            add_url += f"/{folder}"
        add_url += "/addItem"

        with open(csv_path, "rb") as f:
            upload_resp = requests.post(
                add_url,
                data={
                    "f":           "json",
                    "token":       token,
                    "title":       layer_title,
                    "type":        "CSV",
                    "tags":        "ice,detention,automated,facilities",
                    "description": (
                        f"ICE detention facility data auto-published on {date_label}. "
                        "Source: ice.gov detention management page."
                    ),
                    "snippet": f"ICE detention facilities - {date_label}",
                },
                files={"file": (os.path.basename(csv_path), f, "text/csv")},
                timeout=60,
            )
        upload_resp.raise_for_status()
        upload_json = upload_resp.json()
        if not upload_json.get("success"):
            log.error(f"CSV upload failed: {upload_json}")
            return None, None, None

        csv_item_id = upload_json["id"]
        log.info(f"CSV item uploaded via REST: id={csv_item_id}")
        csv_item = gis.content.get(csv_item_id)

        log.info("Publishing as hosted feature layer...")
        try:
            feature_layer_item = csv_item.publish(
                publish_parameters={
                    "type":               "csv",
                    "name":               safe_name,
                    "locationType":       "coordinates",
                    "latitudeFieldName":  "latitude",
                    "longitudeFieldName": "longitude",
                }
            )
        except Exception as e:
            log.error(f"Publish failed: {e}")
            try:
                csv_item.delete()
            except Exception:
                pass
            return None, None, None

        layer_url = f"{feature_layer_item.url}/0"
        log.info(f"New AGOL layer published: {layer_url}")
        log.info(f"  Item ID : {feature_layer_item.id}")
        log.info(f"  Title   : {feature_layer_item.title}")
        return layer_url, feature_layer_item.id, gis

    except Exception as e:
        log.error(f"AGOL publish failed: {e}", exc_info=True)
        return None, None, None


# STEP 7 — Add the new layer to existing AGOL web map
# ─────────────────────────────────────────────────────────────────────────
def add_layer_to_map(
    gis: "GIS",
    layer_item_id: str,
    layer_url: str,
    layer_title: str,
) -> bool:
    map_item_id = os.environ.get("AGOL_MAP_ITEM_ID")
    if not map_item_id:
        log.warning("Map update skipped: set AGOL_MAP_ITEM_ID to enable.")
        return False

    try:
        log.info(f"Opening Web Map: {map_item_id}")
        map_item = gis.content.get(map_item_id)
        if map_item is None:
            log.error(f"Web Map item not found: {map_item_id}")
            return False

        map_data = map_item.get_data()

        existing = map_data.setdefault("operationalLayers", [])
        before = len(existing)
        map_data["operationalLayers"] = [
            l for l in existing
            if "ICE Detention Facilities" not in l.get("title", "")
        ]
        removed = before - len(map_data["operationalLayers"])
        if removed:
            log.info(f"Removed {removed} old ICE Detention layer(s) from map")

        map_data["operationalLayers"].append({
            "id":        layer_item_id,
            "title":     layer_title,
            "url":       layer_url,
            "layerType": "ArcGISFeatureLayer",
            "visibility": True,
            "opacity":   0.8,
            "layerDefinition": {
                "drawingInfo": {
                    "renderer": {
                        "type":                   "classBreaks",
                        "field":                  "Total_Detainment",
                        "classificationMethod":   "esriClassifyNaturalBreaks",
                        "minValue":               0,
                        "classBreakInfos": [
                            {
                                "classMaxValue": 250,
                                "symbol": {
                                    "type": "esriSMS", "style": "esriSMSCircle",
                                    "color": [200, 30, 30, 160],
                                    "size": 6,
                                    "outline": {"color": [255, 255, 255, 120], "width": 0.5}
                                },
                                "label": "1 - 250",
                            },
                            {
                                "classMaxValue": 750,
                                "symbol": {
                                    "type": "esriSMS", "style": "esriSMSCircle",
                                    "color": [200, 30, 30, 160],
                                    "size": 12,
                                    "outline": {"color": [255, 255, 255, 120], "width": 0.5}
                                },
                                "label": "251 - 750",
                            },
                            {
                                "classMaxValue": 1500,
                                "symbol": {
                                    "type": "esriSMS", "style": "esriSMSCircle",
                                    "color": [200, 30, 30, 160],
                                    "size": 20,
                                    "outline": {"color": [255, 255, 255, 120], "width": 0.5}
                                },
                                "label": "751 - 1500",
                            },
                            {
                                "classMaxValue": 9999,
                                "symbol": {
                                    "type": "esriSMS", "style": "esriSMSCircle",
                                    "color": [200, 30, 30, 160],
                                    "size": 32,
                                    "outline": {"color": [255, 255, 255, 120], "width": 0.5}
                                },
                                "label": "1501+",
                            },
                        ],
                    }
                }
            },
            "popupInfo": {
                "title": "{Name}",
                "fieldInfos": [
                    {"fieldName": "Name", "label": "Name", "visible": True},
                    {"fieldName": "Address","label": "Address", "visible": True},
                    {"fieldName": "City", "label": "City", "visible": True},
                    {"fieldName": "State", "label": "State","visible": True},
                    {"fieldName": "Zip", "label": "Zip", "visible": True},
                    {"fieldName": "Full_Address", "label": "Full Address", "visible": True},
                    {"fieldName": "Type_Detailed", "label": "Type", "visible": True},
                    {"fieldName": "FY25_ALOS", "label": "FY25 ALOS", "visible": True},
                    {"fieldName": "FY26_ALOS", "label": "FY26 ALOS", "visible": True},
                    {"fieldName": "Total_Female_Detainment", "label": "Total Female Detainment", "visible": True},
                    {"fieldName": "Total_Male_Detainment",   "label": "Total Male Detainment", "visible": True},
                    {"fieldName": "Total_Detainment", "label": "Total Detainment", "visible": True},
                    {"fieldName": "Total_Non_Criminal", "label": "Total Non-Criminal", "visible": True},
                    {"fieldName": "last_updated", "label": "Last Updated", "visible": True},
                    {"fieldName": "notes", "label": "Notes", "visible": True},
                ],
                "showAttachments": False,
                "mediaInfos": [],
            },
        })

        map_item.update(data=json.dumps(map_data))
        log.info(f"Layer {layer_title!r} added to map {map_item.title!r} ({map_item_id})")
        return True

    except Exception as e:
        log.error(f"Failed to add layer to map: {e}", exc_info=True)
        return False


# RUN PIPELINE
def run_pipeline(
    output_dir: str  = "output",
    dry_run:    bool = False,
) -> dict:
    results: dict[str, str | None] = {
        "raw_file":         None,
        "main_format_file": None,
        "arcgis_csv_file":  None,
        "agol_layer_url":   None,
    }

    timestamp = datetime.now().strftime("%Y%m%d")
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    raw_dir = os.path.join(output_dir, "raw")

    url = find_detention_stats_link() or FALLBACK_URL
    raw_path, source_date = download_xlsx(url, raw_dir)
    if not raw_path:
        log.error("Pipeline aborted: download failed")
        return results
    results["raw_file"] = raw_path

    df = parse_source_excel(raw_path, source_date)
    if df is None or df.empty:
        log.error("Pipeline aborted: could not parse source file")
        return results

    main_format_path = os.path.join(output_dir, f"MAIN_FORMAT_{timestamp}.xlsx")
    if not write_main_format(df, main_format_path):
        log.error("Pipeline aborted: failed to write MAIN FORMAT")
        return results
    results["main_format_file"] = main_format_path

    arcgis_path = os.path.join(output_dir, f"arcgis_import_{timestamp}.csv")
    if not export_arcgis_csv(df, arcgis_path):
        return results
    results["arcgis_csv_file"] = arcgis_path

    if dry_run:
        log.info("[DRY RUN] Local pipeline complete. Skipping AGOL publish + map update.")
    else:
        layer_url, layer_item_id, gis = publish_new_agol_layer(
            arcgis_path, run_date=source_date
        )
        if layer_url:
            results["agol_layer_url"] = layer_url
            date_label = source_date or datetime.now().strftime("%Y-%m-%d")
            add_layer_to_map(
                gis, layer_item_id, layer_url,
                f"ICE Detention Facilities {date_label}",
            )

    return results


# CLI — command-line interface
def main():
    parser = argparse.ArgumentParser(
        description="ICE Detention Data Pipeline — scrape -> format -> geocode -> ArcGIS",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Full automated run (default)
  python ice_scraper.py

  # Run locally without touching ArcGIS
  python ice_scraper.py --dry-run

  # Verify end-to-end including AGOL publish
  python ice_scraper.py --test-agol

Required environment variables:
  AGOL_USERNAME     your ArcGIS Online username
  AGOL_PASSWORD     your ArcGIS Online password
  AGOL_MAP_ITEM_ID  Web Map item ID (from URL: ?webmap=<THIS_PART>)

Optional:
  AGOL_FOLDER       content folder to publish into (default: root)
        """,
    )
    parser.add_argument("--output-dir", default="output",
                        help="Directory for all outputs (default: ./output)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Run scrape -> download -> parse -> CSV locally; skip all AGOL calls")
    parser.add_argument("--test-agol", action="store_true",
                        help="Run one full pipeline pass including AGOL, then exit")
    args = parser.parse_args()

    log.info("=" * 60)
    log.info("ICE Detention Data Pipeline")
    log.info("=" * 60)

    if args.test_agol:
        log.info("[TEST-AGOL] Running one full pipeline pass including AGOL publish...")
        results = run_pipeline(output_dir=args.output_dir, dry_run=False)
        log.info("[TEST-AGOL] Done. Check above for 'New AGOL layer published' and")
        log.info("[TEST-AGOL] 'Layer ... added to map' to confirm success.")
        for key, val in results.items():
            log.info(f"  {key:22s}: {val or '(not produced)'}")
        return

    results = run_pipeline(output_dir=args.output_dir, dry_run=args.dry_run)

    log.info("=" * 60)
    log.info("Pipeline complete. Outputs:")
    for key, path in results.items():
        log.info(f"  {key:22s}: {path or '(not produced)'}")
    log.info("=" * 60)

    for path in results.values():
        if path:
            print(path)

    if not any(results.values()):
        sys.exit(1)


if __name__ == "__main__":
    main()