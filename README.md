# Vulnerability Tracker

Automated CVE tracking tool that runs inside Excel via xlwings. Pulls vulnerabilities from NVD, cross-references with CISA's KEV catalog, prioritizes based on mission impact, and surfaces threats relevant to your tech stack.

## How It Works

On each run the tool:
1. Fetches all CVEs published in the last 30 days from NVD
2. Fetches CISA's live KEV catalog and filters to entries added in the last 30 days
3. For any recently-added KEVs not returned by the NVD 30-day query, fetches them individually from NVD to close the gap
4. Prioritizes each CVE based on KEV status and mission impact setting
5. Writes all results to the Dataset sheet
6. Maps every matching CVE to each technology in your stack on the Dashboard

## Prerequisites

- Python 3.10 or higher
- Microsoft Excel (Windows)
- NVD API key — free from https://nvd.nist.gov/developers/request-an-api-key

## Setup

### 1. Clone the repo

```
git clone https://github.com/FutureZeitgeist/vuln-tracker.git
cd vuln-tracker
```

### 2. Create and activate a virtual environment

```
python -m venv venv
venv\Scripts\activate
```

### 3. Install dependencies

```
pip install -r requirements.txt
```

### 4. Install the xlwings Excel add-in

```
xlwings addin install
```

### 5. Enable VBA project access in Excel

The setup script needs permission to add a macro button to the workbook.

> Excel → Options → Trust Center → Trust Center Settings → Macro Settings
> → Check **Trust access to the VBA project object model**

This only needs to be done once. It can be turned off again after setup if preferred.

### 6. Run the setup script

```
python setup_workbook.py
```

This generates `VulnTracker.xlsm` in the project folder with all sheets, headers, and the **Run Tracker** button pre-configured.

> If VBA access was not enabled, the workbook is still created but without the button.
> See the on-screen instructions printed by the script to add it manually.

## Usage

1. Open `VulnTracker.xlsm`
2. Set mission impact in the **Dashboard** sheet cell B3 (`Low`, `Medium`, or `High`)
3. Add your technologies to the **Input** sheet starting at A11 (one per row)
4. Click **Run Tracker** on the Dashboard and enter your NVD API key when prompted (input is masked)

## Workbook Structure

| Sheet | Purpose |
|-------|---------|
| Dashboard | Mission impact setting (B3), run status (A8), and CVE results by tech stack (A11 onward) |
| Input | Tech stack entries — add one technology per row starting at A11 |
| Dataset | Full CVE dataset from the last 30 days with priority, KEV status, and CVSS scores |

## Dashboard Output (A11 onward)

Each row represents one CVE match for a technology in your stack:

| Column | Content |
|--------|---------|
| A | Technology name |
| B | Priority |
| C | Date Published |
| D | CVE ID |
| E | KEV Status |
| F | CVSS Score |
| G | Description |

If a technology has multiple matching CVEs, it appears on multiple rows sorted highest priority first.

## Prioritization Logic

| Condition | Priority |
|-----------|---------|
| CVE is in CISA KEV catalog | ACT (Immediate) |
| Mission impact = High AND CVSS ≥ 7.0 | ACT (Immediate) |
| CVSS ≥ 7.0 OR mission impact = High | ATTEND (Prioritize) |
| All others | TRACK |

## Security Note

Your NVD API key is entered at runtime via a masked prompt and is never stored in the workbook or codebase.
