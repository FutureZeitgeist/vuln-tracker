import os
import requests
from datetime import datetime, timedelta
from dotenv import load_dotenv
from openpyxl import Workbook

# Load API key from .env file
load_dotenv()
NVD_API_KEY = os.getenv('NVD_API_KEY')

print("Vulnerability Tracker Starting...")
print(f"API Key loaded: {'Yes' if NVD_API_KEY else 'No'}")

import csv

# Load SSVC decision tree
SSVC_TREE_PATH = r"C:\Users\Scott Thomas\Documents\cyber_roadmap\SSVC\data\csv\ssvc\deployer_patch_application_priority_1_0_0.csv"

def load_ssvc_tree():
    """Load SSVC decision tree from CSV"""
    tree = []
    with open(SSVC_TREE_PATH, 'r') as f:
        reader = csv.DictReader(f)
        for row in reader:
            tree.append(row)
    print(f"Loaded {len(tree)} SSVC decision paths")
    if tree:
        print(f"Column names: {list(tree[0].keys())}") 
    return tree

def fetch_cves_from_nvd(days_back=21):
    """Fetch CVEs from NVD for the last X days with pagination"""
    end_date = datetime.now()
    start_date = end_date - timedelta(days=days_back)
    
    url = "https://services.nvd.nist.gov/rest/json/cves/2.0"
    headers = {"apiKey": NVD_API_KEY}
    
    all_cves = []
    start_index = 0
    
    print(f"Fetching CVEs from {start_date.date()} to {end_date.date()}...")
    
    while True:
        params = {
            "pubStartDate": start_date.strftime("%Y-%m-%dT%H:%M:%S.000"),
            "pubEndDate": end_date.strftime("%Y-%m-%dT%H:%M:%S.000"),
            "startIndex": start_index
        }
        
        response = requests.get(url, headers=headers, params=params)
        
        if response.status_code == 200:
            data = response.json()
            vulns = data.get("vulnerabilities", [])
            all_cves.extend(vulns)
            
            total = data.get("totalResults", 0)
            print(f"Retrieved {len(all_cves)} of {total} CVEs...")
            
            if len(all_cves) >= total:
                break
            
            start_index += 2000
        else:
            print(f"Error: {response.status_code}")
            break
    
    print(f"Total CVEs fetched: {len(all_cves)}")
    return all_cves
    


def fetch_kevs_from_cisa():
    """Fetch Known Exploited Vulnerabilities from CISA"""
    url = "https://www.cisa.gov/sites/default/files/feeds/known_exploited_vulnerabilities.json"
    
    print("Fetching KEV catalog from CISA...")
    response = requests.get(url)
    
    if response.status_code == 200:
        data = response.json()
        kevs = data.get("vulnerabilities", [])
        print(f"Found {len(kevs)} total KEVs in catalog")
        return {kev["cveID"]: kev for kev in kevs}
    else:
        print(f"Error fetching KEVs: {response.status_code}")
        return {}

def match_tech_stack(vendor, product):
    """Check if CVE applies to our tech stack"""
    tech_stack = {
        'facebook': 'React',
        'meta': 'React', 
        'react': 'React',
        'vercel': 'Vercel',
        'supabase': 'Supabase',
        'postgresql': 'Supabase',
        'postgres': 'Supabase',
        'backblaze': 'BackBlaze B2',
        'node': 'Node.js',
        'nodejs': 'Node.js',
        'npm': 'npm'
    }
    
    vendor_lower = vendor.lower() if vendor != 'N/A' else ''
    product_lower = product.lower() if product != 'N/A' else ''
    
    for key, component in tech_stack.items():
        if key in vendor_lower or key in product_lower:
            return "Yes", component
    
    return "No", "None"


def evaluate_ssvc(cve_id, is_kev, ssvc_tree):
    """Evaluate CVE using SSVC decision tree"""
    # Decision point defaults for your context
    exploitation = "active" if is_kev else "none"
    system_exposure = "open"  # Vercel/Supabase are internet-facing
    automatable = "yes"  # Assume most web vulns are automatable
    human_impact = "high"  # PII exposure = high impact
    
    # Find matching decision in tree
    for decision in ssvc_tree:
        if (decision['Exploitation v1.1.0'] == exploitation and
            decision['System Exposure v1.0.1'] == system_exposure and
            decision['Automatable v2.0.0'] == automatable and
            decision['Human Impact v2.0.2'] == human_impact):
            return decision['Defer, Scheduled, Out-of-Cycle, Immediate v1.0.0']
    
    return "scheduled"  # Default if no match



def export_to_excel(cves, kevs, ssvc_tree, filename=None):
    """Export CVEs to Excel with all available fields"""
    if filename is None:
        timestamp = datetime.now().strftime("%Y_%m_%d_%H%M")
        filename = f"{timestamp}_v_rpt.xlsx"
    
    wb = Workbook()
    ws = wb.active
    ws.title = "CVE Report"
    
# Full headers
    ws.append(["CVE ID", "Published Date", "Last Modified", "Description", 
               "CVSS Version", "Base Score", "Severity", "Vector String",
               "CWE ID", "CWE Name", "Vendor", "Product", "Affected Versions",
               "Is KEV", "KEV Date Added", "KEV Due Date", 
               "Applies to Stack", "Stack Component",
               "SSVC Exploitation", "SSVC System Exposure", "SSVC Automatable", 
               "SSVC Human Impact", "SSVC Priority", "References"])
    
    for vuln in cves:
        cve = vuln['cve']
        cve_id = cve['id']
        
        # Basic info
        pub_date = cve.get('published', 'N/A')
        mod_date = cve.get('lastModified', 'N/A')
        desc = cve.get('descriptions', [{}])[0].get('value', 'No description')
        
        # CVSS scoring - handle both v3.1 and v2
        metrics = cve.get('metrics', {})
        
        # Try CVSS v3.1 first, then v3.0, then v2
        if 'cvssMetricV31' in metrics and metrics['cvssMetricV31']:
            cvss_data = metrics['cvssMetricV31'][0]
            cvss_info = cvss_data.get('cvssData', {})
            severity = cvss_info.get('baseSeverity', 'N/A')
        elif 'cvssMetricV30' in metrics and metrics['cvssMetricV30']:
            cvss_data = metrics['cvssMetricV30'][0]
            cvss_info = cvss_data.get('cvssData', {})
            severity = cvss_info.get('baseSeverity', 'N/A')
        elif 'cvssMetricV2' in metrics and metrics['cvssMetricV2']:
            cvss_data = metrics['cvssMetricV2'][0]
            cvss_info = cvss_data.get('cvssData', {})
            severity = cvss_info.get('baseSeverity', 'N/A')
        else:
            cvss_info = {}
            severity = 'N/A'
        
        cvss_version = cvss_info.get('version', 'N/A')
        base_score = cvss_info.get('baseScore', 'N/A')
        vector = cvss_info.get('vectorString', 'N/A')
        
        # Weakness (CWE)
        weaknesses = cve.get('weaknesses', [{}])[0].get('description', [{}])
        cwe_id = weaknesses[0].get('value', 'N/A') if weaknesses else 'N/A'
        cwe_name = 'N/A'
        
        # Affected products
        configs = cve.get('configurations', [{}])[0].get('nodes', [])
        vendor = product = versions = 'N/A'
        if configs:
            cpe = configs[0].get('cpeMatch', [{}])[0].get('criteria', '')
            parts = cpe.split(':')
            if len(parts) > 4:
                vendor = parts[3]
                product = parts[4]
                versions = parts[5] if len(parts) > 5 else 'N/A'
        
        # References
        refs = cve.get('references', [])
        ref_urls = '; '.join([r.get('url', '') for r in refs[:3]])
        
        # KEV information
        is_kev = "Yes" if cve_id in kevs else "No"
        kev_date = kevs[cve_id].get('dateAdded', '') if cve_id in kevs else ''
        kev_due = kevs[cve_id].get('dueDate', '') if cve_id in kevs else ''
        
        # Tech stack matching
        applies_to_stack, stack_component = match_tech_stack(vendor, product)
        
        # SSVC evaluation

        # SSVC evaluation
        ssvc_result = evaluate_ssvc(cve_id, is_kev == "Yes", ssvc_tree)
        exploitation = "active" if is_kev == "Yes" else "none"
        
        # Add row to Excel
        ws.append([cve_id, pub_date, mod_date, desc, cvss_version, base_score, 
                   severity, vector, cwe_id, cwe_name, vendor, product, versions,
                   is_kev, kev_date, kev_due, 
                   applies_to_stack, stack_component,
                   exploitation, "open", "yes", "high", ssvc_result, ref_urls])
    
    wb.save(filename)
    print(f"Excel report saved: {filename}")


# Test the functions
cves = fetch_cves_from_nvd(days_back=30)
kevs = fetch_kevs_from_cisa()
ssvc_tree = load_ssvc_tree()
export_to_excel(cves, kevs, ssvc_tree)
print(f"Total CVEs: {len(cves)}")
print(f"KEVs in report: {sum(1 for v in cves if v['cve']['id'] in kevs)}")