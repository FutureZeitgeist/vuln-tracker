"""
setup_workbook.py

Run this script once to generate VulnTracker.xlsm with the correct sheet
structure, headers, and Run button wired to vuln_tracker.py via xlwings.

Requirements:
  - pip install -r requirements.txt
  - Microsoft Excel must be installed
  - xlwings add-in must be installed (run: xlwings addin install)
  - 'Trust access to the VBA project object model' must be enabled in
    Excel > Options > Trust Center > Trust Center Settings > Macro Settings
"""

import sys
from pathlib import Path


def check_vba_access(wb):
    """Return True if the VBA project object model is accessible."""
    try:
        _ = wb.VBProject.VBComponents.Count
        return True
    except Exception:
        return False


def add_vba_module(wb):
    """Add the RunPython macro that calls vuln_tracker.main()."""
    vba_code = (
        "Sub RunVulnTracker()\n"
        '    RunPython "import vuln_tracker; vuln_tracker.main()"\n'
        "End Sub\n"
    )
    module = wb.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
    module.Name = "VulnTrackerModule"
    module.CodeModule.AddFromString(vba_code)


def setup_dashboard(ws):
    """Configure the Dashboard sheet layout and headers."""
    # Title
    ws.Range("A1").Value = "Vulnerability Tracker"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16

    # Mission impact label and default value
    ws.Range("A3").Value = "Mission Impact:"
    ws.Range("A3").Font.Bold = True
    ws.Range("B3").Value = "Medium"

    # Dropdown validation for mission impact (Low / Medium / High)
    dv = ws.Range("B3").Validation
    dv.Add(3, 1, 1, "Low,Medium,High")  # 3 = xlValidateList
    dv.ShowError = True
    dv.ErrorMessage = "Please select Low, Medium, or High."

    # Status cell — updated by the script at runtime
    ws.Range("A8").Value = "Status: Ready"

    # Column headers for CVE results (row 10)
    headers = [
        "Technology", "Priority", "Date Published",
        "CVE ID", "KEV Status", "Score", "Description"
    ]
    for col, header in enumerate(headers, start=1):
        cell = ws.Cells(10, col)
        cell.Value = header
        cell.Font.Bold = True

    # Column widths (characters)
    col_widths = [22, 22, 16, 18, 12, 8, 80]
    for col, width in enumerate(col_widths, start=1):
        ws.Columns(col).ColumnWidth = width


def setup_input_sheet(ws):
    """Configure the Input sheet where users enter their tech stack."""
    ws.Range("A1").Value = "Tech Stack Input"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    # Instructions
    ws.Range("A3").Value = (
        "Add one technology per row starting at A11. "
        "Names are matched against CVE descriptions (case-insensitive). "
        "Examples: apache, windows, chrome, openssl"
    )
    ws.Range("A3").WrapText = True
    ws.Rows(3).RowHeight = 40
    ws.Columns("A").ColumnWidth = 50

    ws.Range("A10").Value = "Technology"
    ws.Range("A10").Font.Bold = True


def setup_dataset_sheet(ws):
    """Configure the Dataset sheet that holds the full CVE pull."""
    headers = [
        "Date Published",
        "CVE ID",
        "Priority",
        "KEV Status",
        "Score",
        "Description"
    ]
    for col, header in enumerate(headers, start=1):
        cell = ws.Cells(1, col)
        cell.Value = header
        cell.Font.Bold = True

    col_widths = [16, 20, 22, 12, 8, 80]
    for col, width in enumerate(col_widths, start=1):
        ws.Columns(col).ColumnWidth = width


def add_run_button(ws, macro_name):
    """Add a Run Tracker button to the Dashboard sheet."""
    btn = ws.Buttons().Add(100, 62, 120, 28)  # left, top, width, height (points)
    btn.OnAction = macro_name
    btn.Caption = "Run Tracker"
    btn.Name = "btnRunTracker"


def main():
    try:
        import win32com.client as win32
    except ImportError:
        print("Error: pywin32 is not installed. Run: pip install pywin32")
        sys.exit(1)

    output_path = str(Path(__file__).parent / "VulnTracker.xlsm")

    print("Starting Excel...")
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        print("Creating workbook...")
        wb = excel.Workbooks.Add()

        # Set up three sheets: Dashboard, Input, Dataset
        wb.Worksheets(1).Name = "Dashboard"
        while wb.Worksheets.Count > 1:
            wb.Worksheets(wb.Worksheets.Count).Delete()

        ws_input = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws_input.Name = "Input"

        ws_dataset = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws_dataset.Name = "Dataset"

        ws_dashboard = wb.Worksheets("Dashboard")

        print("Configuring sheets...")
        setup_dashboard(ws_dashboard)
        setup_input_sheet(ws_input)
        setup_dataset_sheet(ws_dataset)

        print("Adding VBA macro...")
        if check_vba_access(wb):
            add_vba_module(wb)
            add_run_button(ws_dashboard, "RunVulnTracker")
            print("Run button added to Dashboard.")
        else:
            print(
                "\nWarning: Could not access the VBA project.\n"
                "The workbook will be created without the Run button.\n\n"
                "To add the button manually:\n"
                "  1. Enable 'Trust access to the VBA project object model':\n"
                "     Excel > Options > Trust Center > Trust Center Settings > Macro Settings\n"
                "  2. Re-run this script, OR open the VBA editor (Alt+F11) and add:\n\n"
                "       Sub RunVulnTracker()\n"
                '           RunPython "import vuln_tracker; vuln_tracker.main()"\n'
                "       End Sub\n\n"
                "  3. Insert a button on the Dashboard sheet and assign it to RunVulnTracker.\n"
            )

        print(f"Saving to {output_path}...")
        wb.SaveAs(output_path, 52)  # 52 = xlOpenXMLWorkbookMacroEnabled
        wb.Close(False)

        print(f"\nDone. Open VulnTracker.xlsm to get started.")
        print("Next steps:")
        print("  1. Install the xlwings add-in if not already done: xlwings addin install")
        print("  2. Add your technologies to the Input sheet (A11 onward)")
        print("  3. Click Run Tracker on the Dashboard and enter your NVD API key")

    finally:
        excel.DisplayAlerts = True
        excel.Quit()


if __name__ == "__main__":
    main()
