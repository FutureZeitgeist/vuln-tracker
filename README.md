

\# Vulnerability Tracker



Automated CVE tracking tool that pulls vulnerabilities from NVD, cross-references with CISA's KEV catalog, applies SSVC prioritization, and identifies threats relevant to your tech stack.



\## Prerequisites



\- Python 3.10 or higher

\- NVD API key (free from https://nvd.nist.gov/developers/request-an-api-key)

\- SSVC decision tree CSV file



\## Setup Instructions



1\. Clone or download this project



2\. Create virtual environment:

&nbsp;  powershell

&nbsp;  python -m venv venv





3\. Activate virtual environment:

&nbsp;  powershell

&nbsp;  .\\venv\\Scripts\\Activate.ps1





4\. Install dependencies:

powershell

&nbsp;  pip install -r requirements.txt





5\. Configure environment variables:

&nbsp;  - Create a `.env` file in the project root

&nbsp;  - Add your NVD API key:



&nbsp;    NVD\_API\_KEY=your-api-key-here





6\. Update SSVC path in vuln\_tracker.py:

&nbsp;  - Open `vuln\_tracker.py`

&nbsp;  - Update line 18 with your SSVC CSV file path:

python

&nbsp;    SSVC\_TREE\_PATH = r"YOUR\_PATH\_HERE\\deployer\_patch\_application\_priority\_1\_0\_0.csv"





7\. Customize tech stack (optional):

&nbsp;  - Edit the `match\_tech\_stack()` function (lines 90-110)

&nbsp;  - Add your specific technologies to the tech\_stack dictionary



\## Usage



Run the script:

powershell

python vuln\_tracker.py





The script will:

\- Fetch CVEs from the last 30 days

\- Download CISA's KEV catalog

\- Apply SSVC prioritization

\- Match vulnerabilities to your tech stack

\- Generate a timestamped Excel report: `YYYY\_MM\_DD\_HHMM\_v\_rpt.xlsx`



\## Output



The Excel report includes:

\- CVE details (ID, dates, description, CVSS scores)

\- KEV status and due dates

\- Tech stack relevance (Yes/No + component name)

\- SSVC prioritization (immediate/out-of-cycle/scheduled/defer)

\- References and affected products



\## Recommended Schedule



Run weekly to stay current on new vulnerabilities.



\## Tech Stack (Default Configuration)



\- React

\- Vercel

\- Supabase/PostgreSQL

\- BackBlaze B2

\- Node.js

\- npm



Customize the `match\_tech\_stack()` function to match your environment.



\## Security Note



Never commit your `.env` file to version control. The `.gitignore` file is configured to exclude it.



