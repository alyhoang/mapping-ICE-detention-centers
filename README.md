# MappingICE — ICE Detention Data Pipeline

An automated pipeline that scrapes ICE's public detention statistics, formats and geocodes facility data, and publishes a live bubble map to ArcGIS Online — updated biweekly.

> Data source: [ICE ERO Statistics Dashboard](https://www.ice.gov/detention-management)
> 
---

## Requirements

| Requirement | Details |
|---|---|
| Operating System | Windows, macOS, Linux |
| Python | Version 3.11 or newer |
| ArcGIS Online Account | Must have access to the web map URL and account username/password |
| GitHub Account | Required for biweekly scheduling via GitHub Actions |

---

## Repository Structure

```
MappingICE/
├── ice_scraper.py              # Main pipeline script
├── requirements.txt            # Python dependencies
├── .github/
│   └── workflows/
│       └── scheduler.yml       # GitHub Actions biweekly schedule
└── output/                     # Generated outputs (gitignored)
    ├── raw/                    # Downloaded ICE Excel files
    ├── MAIN_FORMAT_YYYYMMDD.xlsx
    └── arcgis_import_YYYYMMDD.csv
```

---

## Setup & Installation

### Step 1 — Fork and Clone the Repository

**Forking** creates your own copy of this project on GitHub, which you can run and modify independently.

1. Go to [github.com/alyhoang/mapping-ICE-detention-centers](https://github.com/alyhoang/mapping-ICE-detention-centers)
2. Click the **Fork** button in the top-right corner
3. Under "Owner", select your GitHub account, then click **Create fork**

Now clone your fork to your computer:

```bash
git clone https://github.com/<your-github-username>/mapping-ICE-detention-centers.git
cd mapping-ICE-detention-centers
```

Replace `<your-github-username>` with your actual GitHub username.

> If you don't have Git installed, download it at [git-scm.com/downloads](https://git-scm.com/downloads) and re-run the commands above.

### Step 2 — Install Python

Go to [python.org/downloads](https://www.python.org/downloads/) and follow the instructions for your operating system (Windows, macOS, or Linux). Install newest version.

### Step 3 — Open the Terminal

- **macOS**: press `Cmd + Space`, type `Terminal`, press Enter
- **Windows**: press `Win + R`, type `cmd`, press Enter
- **Linux**: open your terminal application

### Step 4 — Navigate to the Project Folder

If you opened a new terminal window after Step 1, navigate back into the project folder by typing:

```bash
cd mapping-ICE-detention-centers
```

### Step 5 — Create and Activate the Python Environment

```bash
python -m venv venv
source venv/bin/activate        # macOS / Linux
```

If successful, you will see `(venv)` appear at the start of your command line:

```
(venv) user@computer MappingICE %
```

### Step 6 — Install Dependencies

```bash
pip install -r requirements.txt
```

> If you get an error, try `pip3 install -r requirements.txt` instead.
> If prompted to upgrade pip, run the suggested upgrade command and then re-run the install.
>
> **Note:** if `pip3` works but `pip` does not, use `python3` anywhere this guide says `python` for the rest of the steps.

This installs the following packages:

```
pandas
requests
beautifulsoup4
openpyxl
arcgis==2.4.3
```

### Step 7 — Set ArcGIS Credentials

The script needs to know which ArcGIS account to use and which map to publish to. Run the following lines one at a time:

```bash
export AGOL_USERNAME='your_username'
export AGOL_PASSWORD='your_password'
export AGOL_MAP_ITEM_ID='your_map_item_id'
```

**Finding your Map Item ID:** open your ArcGIS web map and copy the alphanumeric string after `?webmap=` in the URL:

```
https://user.maps.arcgis.com/apps/mapviewer/index.html?webmap=adb4b9703f5642799c9075d839484980
                                                                      ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
                                                                      this is your AGOL_MAP_ITEM_ID
```

To verify the credentials were set correctly:

```bash
echo "AGOL_USERNAME: $AGOL_USERNAME"
echo "AGOL_PASSWORD: $AGOL_PASSWORD"
echo "AGOL_MAP_ITEM_ID: $AGOL_MAP_ITEM_ID"
```

> **Security note:** these values are not saved anywhere and reset every time you close the terminal. You will need to re-enter them each session.
>
> **Tip:** Best practice is to manually type in these prompts instead of copy-pasting to ensure vertical apostrophes are used -— some terminals insert curly/tilted quotes that break the command.

---

## Running the Pipeline

### Full run — scrape, format, geocode, and publish to ArcGIS

```bash
python ice_scraper.py
```

### Local only — runs everything except the ArcGIS publish step

```bash
python ice_scraper.py --dry-run
```

Use this to verify the scraping and geocoding are working without touching ArcGIS Online.

### End-to-end test including ArcGIS publish

```bash
python ice_scraper.py --test-agol
```

---

## Outputs

| File | Description |
|---|---|
| `raw/FY26_detentionStats_MMDDYYYY.xlsx` | Original file downloaded from ice.gov |
| `MAIN_FORMAT_YYYYMMDD.xlsx` | Cleaned and formatted facility data |
| `arcgis_import_YYYYMMDD.csv` | Geocoded CSV ready for ArcGIS import |

### MAIN FORMAT columns

| Column | Description |
|---|---|
| Name | Facility name |
| Address, City, State, Zip | Location fields |
| Full Address | Combined single-line address |
| Type Detailed | Facility type (BOP, DIGSA, GSA, SPS, USMS) |
| FY25 ALOS / FY26 ALOS | Average length of stay in days |
| Total Male Detainment | Total male detainees (average nightly) |
| Total Female Detainment | Total female detainees (average nightly) |
| Total Detainment | Combined total (used for bubble map sizing) |
| Total Non-Criminal | Detainees with no prior criminal conviction |
| notes | Manual annotations |
| last_updated | Date of last ICE data update |


---

## Geocoding

Facilities are geocoded using the [Census Bureau Batch Geocoder](https://geocoding.geo.census.gov/geocoder/). Addresses are pre-cleaned before submission:

- Trailing periods removed (`DR.` → `DR`, `RD.` → `RD`)
- Directional abbreviations expanded (`S.W.` → `SW`)
- Highway formats normalised (`US 90` → `US Highway 90`)
- City abbreviations expanded (`FT.` → `Fort`, `ST.` → `Saint`)
- P.O. Box, military bases, and rural route addresses are skipped (the geocoder cannot resolve them)
- Facilities in US territories (Guam, CNMI, Puerto Rico) are not matched by the Census geocoder

Typical match rate for the 04/29/2026 update: ~165 / 203 facilities per run. The remaining ~38 are ungeocoded due to the limitations above and will not appear on the map.

---

## Automated Scheduling with GitHub Actions

The pipeline runs automatically every two weeks via GitHub Actions without any manual intervention. ICE updates its data roughly every two weeks on a Wednesday; the scheduler is aligned to match.

> You can also trigger a manual run anytime from the Actions tab in GitHub by clicking **Run workflow**. Manual runs always execute regardless of the biweekly schedule.

### How the schedule works

The workflow fires every Wednesday at 12:00 UTC but includes a **biweekly gate** — a small script that checks whether the current week is an "on" week based on an anchor date. On "off" weeks the job exits immediately without running the pipeline.

---

## Scheduler Setup

### Step 1 — Confirm your fork is on GitHub

If you completed the Fork step in Setup, your repository is already on GitHub at:

```
https://github.com/<your-github-username>/mapping-ICE-detention-centers
```

It will already contain all required files:

```
ice_scraper.py
requirements.txt
.github/workflows/scheduler.yml
```

### Step 2 — Add repository secrets

Your ArcGIS credentials must be stored as encrypted GitHub secrets — never hardcode them in the script or commit them to the repo.

1. Go to your forked repository: `https://github.com/<your-github-username>/mapping-ICE-detention-centers`
2. Click **Settings** → **Secrets and variables** → **Actions**
3. Click **New repository secret** and add each of the following one at a time:

| Secret Name | Value |
|---|---|
| `AGOL_USERNAME` | Your ArcGIS Online username |
| `AGOL_PASSWORD` | Your ArcGIS Online password |
| `AGOL_MAP_ITEM_ID` | Web Map item ID from the AGOL URL |

> Secrets can be updated anytime by clicking the secret name → **Update**. You can never view an existing secret value, only overwrite it.

### Step 3 — Enable Actions

1. Go to the **Actions** tab in your repository
2. If prompted, click **I understand my workflows, go ahead and enable them**
3. You should see **ICE Detention Data Pipeline** listed in the left sidebar

### Step 4 — Test with a manual run

Before waiting for the scheduled Wednesday, trigger a run manually to confirm everything works:

1. Go to **Actions** → **ICE Detention Data Pipeline**
2. Click **Run workflow** → **Run workflow**
3. Watch the logs in real time — a green checkmark means success
4. When complete, click the run → **Artifacts** to download the output files

### Step 5 — Verify the biweekly anchor

The anchor date in `scheduler.yml` controls which Wednesdays are active run weeks. It is currently set to `2026-01-07`. To change it, edit this line and push:

```yaml
anchor="2026-01-07"
```

Pick any Wednesday you want the pipeline to land on.

---

## Data Notes

- **Total Detainment** is average nightly population, not a point-in-time count
- Numbers are rounded integers derived from ICE's published averages
- ~38 facilities per run are ungeocoded and will not appear on the map (PO Box addresses, rural routes, or territory locations outside Census geocoder coverage)

---

## Dependencies

```
pandas
requests
beautifulsoup4
openpyxl
arcgis==2.4.3
```

---

## Credits

Scraping logic adapted from [icewatch](https://github.com/lockdown-systems/icewatch/blob/main/src/icewatch/ice_detention_scraper.py) by Lockdown Systems.
