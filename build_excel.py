import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Notifications Framework"

headers = [
    "Group",
    "Source",
    "Trigger Logic",
    "Update vs. Alert",
    "Where User Is Notified",
    "How User Resolves",
    "Where Data Is Stored",
    "For How Long",
    "Clarification Flag / Notes",
]

rows = [
    ["Interactions Report", "Interactions Report",
     "Interaction Completed (user is attended); Interaction Report created to log notes",
     "Alert", "Activity Feed, Teams, Int. Hub, Int. Report",
     "Data Source (primary)", "Original data source (primary)", "",
     ""],

    ["Interactions Report", "Interactions Report",
     "Interaction Completed (user not attended; no deal team); Interaction Report created to log notes",
     "Update", "Activity Feed, Teams, Int. Hub, Int. Report",
     "Activity Feed (primary)", "No trail (primary)", "",
     ""],

    ["Interactions Report", "Interactions Report",
     "Upcoming Tasks in next 3 days (user has been assigned the task)",
     "Alert", "Activity Feed",
     "Data Source (primary)", "Original data source (primary)", "",
     ""],

    ["Interactions Report", "Interactions Report",
     "Upcoming Tasks in next 3 days (user assigned the task to someone else)",
     "Update", "Activity Feed",
     "Activity Feed (primary)", "No trail (primary)", "",
     ""],

    ["Interactions Report", "Interactions Report",
     "A task you assigned to another user has been deleted",
     "Update", "Activity Feed",
     "Activity Feed (primary)", "No trail (primary)", "",
     ""],

    ["Opportunity Details / Opportunities", "Opportunity Details / Opportunities",
     "After user creates a change to Accountable Banker, Deal State, Est. Close Date, Jnl, Frc, Probability on a deal this user is staffed on",
     "Update", "Activity Feed",
     "Activity Feed (primary)", "Original data source NEW – Changelog (primary)", "",
     "[C2] Trigger text partially legible in source screenshot — verify wording. [C6] 'NEW – Changelog' label meaning not explained in source."],

    ["Opportunity Details / Opportunities", "Opportunity Details / Opportunities",
     "Change in Deal State to Closed-Won",
     "Alert", "Activity Feed",
     "Data Source (primary)", "Original data source (primary)", "",
     ""],

    ["Opportunity Details / Opportunities", "Opportunity Details",
     "When new opportunity is created or coverage users updated to include banker they are alerted",
     "Update", "Activity Feed",
     "Activity Feed (primary)", "No trail (primary)", "",
     ""],

    ["Engagement Health", "Engagement Health",
     "Banking's last interaction date coming past due in 3 days; My last Interaction date coming past due in 5 days",
     "Alert", "Activity Feed",
     "Data Source (primary)", "Original data source (primary)", "",
     "[C5] Unclear if 3-day and 5-day thresholds are one combined rule or two separate scenarios."],

    ["Signals", "Signals",
     "News item flagged for action",
     "Alert", "Activity Feed",
     "Data Source (primary)", "No trail (primary)", "",
     ""],

    ["Approvals Queue", "Approvals Queue",
     "When another banker approves a client tier change, for example, that this banker's client tier change has been approved",
     "Update", "Operations",
     "Activity Feed (primary)", "No trail (primary)", "",
     ""],

    ["Approvals Queue", "Approvals Queue",
     "When another banker requests a client tier change, for example, that this banker needs to approve",
     "Alert", "Operations",
     "Activity Feed (primary), Operations (primary)", "No trail (primary)", "",
     ""],

    ["Hierarchy Cases", "Hierarchy Cases",
     "When another banker approves a client legal entity change, for example, that this banker requested for an account they cover",
     "Update", "Operations",
     "Activity Feed (primary)", "Consolidated data trail (primary)", "",
     ""],

    ["Hierarchy Cases", "Hierarchy Cases",
     "When another banker requests a client legal entity change, for example, that this banker has clearance to approve",
     "Alert", "Operations",
     "Operations (primary)", "Consolidated data trail (primary)", "",
     "[C3] Operations appears highlighted as primary resolve path; Activity Feed state is visually ambiguous in source screenshot."],
]

# Write headers
header_fill = PatternFill(start_color="1A1A1A", end_color="1A1A1A", fill_type="solid")
header_font = Font(color="FFFFFF", bold=True, size=10)
for col_idx, header in enumerate(headers, start=1):
    cell = ws.cell(row=1, column=col_idx, value=header)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(wrap_text=True, vertical="center")

# Write data rows
group_fill = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")
note_font = Font(color="B45309", italic=True, size=9)
wrap_align = Alignment(wrap_text=True, vertical="top")

current_group = None
row_idx = 2
for row in rows:
    for col_idx, value in enumerate(row, start=1):
        cell = ws.cell(row=row_idx, column=col_idx, value=value)
        cell.alignment = wrap_align
        if col_idx == 9 and value:
            cell.font = note_font
    row_idx += 1

# Column widths
widths = [22, 26, 50, 14, 28, 26, 32, 16, 50]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w

ws.freeze_panes = "A2"
ws.row_dimensions[1].height = 30

# Add a notes/legend sheet
notes_ws = wb.create_sheet("Notes & Legend")
notes_ws["A1"] = "How to use this sheet"
notes_ws["A1"].font = Font(bold=True, size=12)
notes_ws["A3"] = "1. Edit any copy/text cells directly in the 'Notifications Framework' tab."
notes_ws["A4"] = "2. Do not rename column headers — they map directly to the HTML structure."
notes_ws["A5"] = "3. 'For How Long' column is currently blank — fill in retention values if known."
notes_ws["A6"] = "4. The 'Clarification Flag / Notes' column lists open questions [C1]-[C7] from the original review."
notes_ws["A7"] = "5. Send this file back and changes will be applied to notifications-framework.html."
notes_ws.column_dimensions["A"].width = 100
for r in range(3, 8):
    notes_ws[f"A{r}"].alignment = Alignment(wrap_text=True)

wb.save("/home/user/claude-code-workspace/notifications-framework.xlsx")
print("saved")
