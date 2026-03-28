import xlwings as xw
import requests
import tkinter as tk
from tkinter import simpledialog
from datetime import datetime, timedelta

def main():
    wb = xw.Book.caller()
    dashboard = wb.sheets['Dashboard']
    dataset_sheet = wb.sheets['Dataset']
    inputs_sheet = wb.sheets['Input']

    # 1. Setup and Stakeholder Input
    dashboard.range('A8').value = "Status: Initializing..."
    mission_impact = str(dashboard.range('B3').value or "Medium").strip()

    # Read tech stack from Inputs sheet A11:A1000
    tech_rows = inputs_sheet.range('A11:A1000').value
    if not isinstance(tech_rows, list):
        tech_rows = [tech_rows]
    tech_stack = [str(x).strip().lower() for x in tech_rows if x]

    # 2. Application Programming Interface Key with Masked Input
    root = tk.Tk()
    root.withdraw()
    # The 'show' parameter ensures characters appear as asterisks for security
    api_key = simpledialog.askstring("Security", "Enter National Vulnerability Database API Key:", show='*')
    
    if not api_key:
        dashboard.range('A8').value = "Status: Error - No API Key"
        return

    # 3. Date Range (30 Days)
    end_date = datetime.now()
    start_date = end_date - timedelta(days=30)
    start_str = start_date.strftime('%Y-%m-%dT%H:%M:%S.000')
    end_str = end_date.strftime('%Y-%m-%dT%H:%M:%S.000')

    # 4. Fetch Data
    dashboard.range('A8').value = "Status: Fetching Data..."
    kev_url = "https://www.cisa.gov/sites/default/files/feeds/known_exploited_vulnerabilities.json"
    kev_data = requests.get(kev_url).json().get('vulnerabilities', [])
    kev_lookup = [v['cveID'] for v in kev_data]

    # Filter KEVs added in the last 30 days
    recent_kevs = []
    for v in kev_data:
        try:
            date_added = datetime.strptime(v['dateAdded'], '%Y-%m-%d')
            if date_added >= start_date.replace(hour=0, minute=0, second=0, microsecond=0):
                recent_kevs.append(v['cveID'])
        except:
            pass

    nvd_url = f"https://services.nvd.nist.gov/rest/json/cves/2.0/?pubStartDate={start_str}&pubEndDate={end_str}"
    headers = {'apiKey': str(api_key).strip()}

    try:
        response = requests.get(nvd_url, headers=headers)
        response.raise_for_status()
        vulnerabilities = response.json().get('vulnerabilities', [])
    except:
        dashboard.range('A8').value = "Status: API Error"
        return

    # 4b. Find recently-added KEVs missing from the NVD 30-day window and fetch them individually
    nvd_ids = {item.get('cve', {}).get('id') for item in vulnerabilities}
    missing_kevs = [cve_id for cve_id in recent_kevs if cve_id not in nvd_ids]

    dashboard.range('A8').value = f"Status: Fetching {len(missing_kevs)} missing KEVs..."
    for cve_id in missing_kevs:
        try:
            r = requests.get(
                f"https://services.nvd.nist.gov/rest/json/cves/2.0/?cveId={cve_id}",
                headers=headers
            )
            r.raise_for_status()
            extra = r.json().get('vulnerabilities', [])
            vulnerabilities.extend(extra)
        except:
            pass

    # 5. Process Global Results
    final_output = []
    for item in vulnerabilities:
        cve = item.get('cve', {})
        cve_id = cve.get('id')
        pub_date = cve.get('published', 'N/A')[:10]
        description = cve.get('descriptions', [{}])[0].get('value', "")
        
        metrics = cve.get('metrics', {})
        base_score = "N/A"
        if 'cvssMetricV31' in metrics:
            base_score = metrics['cvssMetricV31'][0].get('cvssData', {}).get('baseScore', "N/A")
        elif 'cvssMetricV30' in metrics:
            base_score = metrics['cvssMetricV30'][0].get('cvssData', {}).get('baseScore', "N/A")
        elif 'cvssMetricV2' in metrics:
            base_score = metrics['cvssMetricV2'][0].get('cvssData', {}).get('baseScore', "N/A")
        
        is_kev = "YES" if cve_id in kev_lookup else "NO"
        
        priority = "TRACK"
        try:
            score_num = float(base_score) if base_score != "N/A" else 0.0
        except:
            score_num = 0.0

        if is_kev == "YES":
            priority = "ACT (Immediate)"
        elif mission_impact.lower() == "high" and score_num >= 7.0:
            priority = "ACT (Immediate)"
        elif score_num >= 7.0 or mission_impact.lower() == "high":
            priority = "ATTEND (Prioritize)"

        final_output.append([pub_date, cve_id, priority, is_kev, base_score, description])

    # 6. Output to Dataset Tab
    dataset_sheet.clear()
    col_headers = ["Date Published", "Common Vulnerabilities and Exposures Identification", "Priority", "Known Exploited Vulnerabilities Status", "Score", "Description"]
    dataset_sheet.range('A1').value = col_headers
    dataset_sheet.range('A2').value = final_output
    dataset_sheet.range('A:F').api.WrapText = False
    dataset_sheet.range('A:F').rows.autofit()

    # 7. Dashboard Mapping - one row per CVE match per tech stack item
    rank_map = {"ACT (Immediate)": 3, "ATTEND (Prioritize)": 2, "TRACK": 1}

    # Clear previous dashboard output area
    dashboard.range('A11:G10000').clear_contents()

    dashboard_rows = []
    for tech_name in tech_stack:
        if not tech_name:
            continue

        tech_matches = [r for r in final_output if tech_name in r[5].lower()]
        tech_matches.sort(key=lambda x: (rank_map.get(x[2], 0), x[4] if x[4] != "N/A" else 0), reverse=True)

        for match in tech_matches:
            dashboard_rows.append([
                tech_name,  # A: Technology
                match[2],   # B: Priority
                match[0],   # C: Date Published
                match[1],   # D: CVE ID
                match[3],   # E: KEV Status
                match[4],   # F: Score
                match[5]    # G: Description
            ])

    if dashboard_rows:
        dashboard.range('A11').value = dashboard_rows

    dashboard.range('A8').value = f"Status: Found {len(final_output)} Threats | {len(dashboard_rows)} Tech Matches"

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import tkinter.messagebox as mb
        mb.showerror("Script Error", str(e))