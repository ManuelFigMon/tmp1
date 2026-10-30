<#
=====================================================================
  Program Name  : formatExcel.Examples.ps1
  Author        : Manuel Figallo
  Purpose       : Runnable examples for formatExcel.ps1.
  Version       : 1.0beta
  Created       : 2026-10-27
  Last Modified : 2026-10-27

  Dependencies:
    formatExcel.ps1 and cgsUtils.ps1, both in this folder.

  Description:
    Every example below was run against the Planner export
    "Data Analytics Governance_10_27_2026.xlsx". Change $Source and run.

  Usage:
    Edit $Source, then run the block you want. Nothing executes on import
    -- the examples are commented so this file is safe to dot-source.
=====================================================================
#>

$Source = "C:\reports\Data Analytics Governance_10_27_2026.xlsx"
$Sheet  = "Consolidated Data"
$Fields = @("Task Name", "Bucket", "Priority", "Due date", "Start date",
            "Checklist Items", "Notes")

# --- Example 1: the status report management asked for -------------------
# Word, corporate look. format_name defaults to "corporate", so it could be
# left out; it is written here so the intent is on the page.
#
# .\formatExcel.ps1 -ExcelPath $Source -SheetName $Sheet -Columns $Fields `
#     -OutputPath "C:\reports\Data Analytics Governance_10_27_2026.docx" `
#     -format_name corporate -Title "Data Analytics Governance"

# --- Example 2: the same report as PDF -----------------------------------
# NEEDS MICROSOFT WORD on this machine: PowerShell cannot write a PDF on
# its own, so the script drives Word through COM. With no Word it stops and
# tells you, rather than quietly handing back a different file type.
#
# .\formatExcel.ps1 -ExcelPath $Source -SheetName $Sheet -Columns $Fields `
#     -OutputPath "C:\reports\Data Analytics Governance_10_27_2026.pdf"

# --- Example 3: no Word on the server ------------------------------------
# Produce HTML instead and print it to PDF from any browser. -OnWordMissing
# makes the substitution a DECISION rather than a surprise.
#
# .\formatExcel.ps1 -ExcelPath $Source -SheetName $Sheet -Columns $Fields `
#     -OutputPath "C:\reports\governance.pdf" -OnWordMissing html

# --- Example 4: keep long Notes readable ---------------------------------
# A Notes column of 400 characters gives one row half a page. Cutting at
# 220 marks every cut with an ellipsis, so nothing disappears silently.
#
# .\formatExcel.ps1 -ExcelPath $Source -SheetName $Sheet -Columns $Fields `
#     -OutputPath "C:\reports\governance.docx" -MaxCellLength 220

# --- Example 5: every column, plain look ---------------------------------
# Omit -Columns to keep the whole sheet. "plain" drops the colour, which
# suits a document that will be photocopied.
#
# .\formatExcel.ps1 -ExcelPath $Source -SheetName $Sheet `
#     -OutputPath "C:\reports\governance_full.docx" -format_name plain

# --- Example 6: all three formats in one pass ----------------------------
# One read of the workbook per format. Useful for a scheduled task that
# mails the PDF and publishes the HTML.
#
# foreach ($extension in "docx", "html") {
#     .\formatExcel.ps1 -ExcelPath $Source -SheetName $Sheet -Columns $Fields `
#         -OutputPath ("C:\reports\governance.{0}" -f $extension) `
#         -Title "Data Analytics Governance" -MaxCellLength 220
# }

# --- Example 7: from SAS, through the cgs_ai wrapper ----------------------
# %formatExcel(ExcelPath=C:\reports\governance.xlsx,
#              SheetName=%str(Consolidated Data),
#              Columns=%str(Task Name;Bucket;Priority;Notes),
#              OutputPath=C:\reports\governance.docx);
#
# NOTE the SEMICOLONS in Columns. Launched with -File -- which is how SAS
# calls it -- every argument arrives as one string, so a list needs a
# delimiter. Commas work too; semicolon is the convention used by every
# other list parameter in this toolkit.

# --- Example 8: drop rows you do not want in the report ------------------
# -ExcludeWhere names a column and the value(s) to DROP. Named for what it
# does: in Excel "filter to BACKLOG" means KEEP it, so a parameter called
# -Filter would eventually be read as the opposite of what was meant.
#
# .\formatExcel.ps1 -ExcelPath $Source -SheetName $Sheet -Columns $Fields `
#     -OutputPath "C:\reports\governance.pdf" `
#     -Title "Data Analytics Infrastructure and Governance" `
#     -ExcludeWhere "Bucket=BACKLOG"

# --- Example 9: several values, and several columns ----------------------
# A PIPE separates values for one column; a SEMICOLON separates clauses.
# A row is dropped when it matches ANY clause, which is what "filter out
# the backlog and the low priority items" means when said out loud.
#
# -ExcludeWhere "Bucket=BACKLOG|TODO (Up next)"
# -ExcludeWhere "Bucket=BACKLOG;Priority=Low"
#
# An EMPTY value selects the blank cells -- "drop the rows nobody
# categorised" is a real thing to ask for:
# -ExcludeWhere "Bucket="

# --- Example 10: keep only what you want, the inverse --------------------
# -IncludeWhere keeps ONLY matching rows. When both are given, Include runs
# first and Exclude then trims what is left.
#
# .\formatExcel.ps1 -ExcelPath $Source -SheetName $Sheet -Columns $Fields `
#     -OutputPath "C:\reports\closed-items.pdf" `
#     -IncludeWhere "Bucket=DONE|UNDER REVIEW"

# --- Example 11: sort the report -----------------------------------------
# "Column", "Column ASC" or "Column DESC"; several keys with a semicolon,
# applied left to right. Blanks sort LAST in both directions -- a DESC
# column that opens with twenty empty cells is not usefully sorted.
#
# -SortBy "Due date DESC"
# -SortBy "Bucket; Task Name"            # Bucket, then Task Name within it
# -SortBy "Bucket DESC; Due date"
#
# The filter and the sort both run BEFORE columns are chosen, so either may
# name a column the report does not display:
# -Columns "Task Name;Priority" -SortBy "Bucket; Due date DESC"

# --- Example 12: the full status report -----------------------------------
# What produced Data_Analytics_Governance_KANBAN_10_30_2026.pdf.
#
# .\formatExcel.ps1 `
#     -ExcelPath "C:\reports\Data Analytics Infrastructure and Governance - KANBAN_10_30_2026.xlsx" `
#     -SheetName "Consolidated Data" `
#     -Columns "Task Name;Bucket;Priority;Created Date;Labels;Notes" `
#     -OutputPath "C:\reports\Data_Analytics_Governance_KANBAN_10_30_2026.pdf" `
#     -format_name corporate `
#     -Title "Data Analytics Infrastructure and Governance" `
#     -Subtitle "KANBAN status report  |  30 October 2026" `
#     -ExcludeWhere "Bucket=BACKLOG" `
#     -SortBy "Bucket; Task Name" `
#     -MaxCellLength 180
#
# The report states the filter on its own banner:
#   "Showing 21 of 26 rows.  Excluded: Bucket = BACKLOG  Sorted by: ..."
# That line is automatic and not optional. A status report that silently
# omits rows misleads whoever reads it, and nothing else in the PDF would
# reveal it.
