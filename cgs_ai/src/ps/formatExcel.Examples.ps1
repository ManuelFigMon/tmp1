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
