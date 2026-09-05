"""
=====================================================================
  Program Name  : Examples.py
  Author        : Manuel Figallo
  Purpose       : Runnable examples for every cgs_ai function.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Description:
    Start here. Example 1 runs anywhere and proves the import works. The
    rest are commented out because they touch the network, a mail server,
    a database or a UNC share -- un-comment the one you want.
=====================================================================
"""

# %run cgs_ai_setup
import cgs_ai
print('cgs_ai imported from:', cgs_ai.__file__)
result = cgs_ai.basic_hello()
print(result)


# --- Example 1: greetings (no dependencies, always safe to run) --------------
print(cgs_ai.personalized_hello("Manuel"))
print(cgs_ai.detailed_hello("formal"))
print("cgs_ai version:", cgs_ai.__version__)


# --- Example 2: scanFileSystem -- one row per keyword match ------------------
# result = cgs_ai.scanFileSystem(
#     input_folder_root=[r"\\A70admed.com\r1\...\UNIT\DME\Logs"],
#     extract_keyword=["real time", "cpu time"],
#     output_file_path=r"\\...\cgs_ai\data\scan_test.csv",
#     lines_above=5, lines_below=5)
# print(len(result["matches"]), "matches ->", result["output"])


# --- Example 3: metric_profile -- produces EXCEL with a Metrics sheet --------
# result = cgs_ai.scanFileSystem(
#     input_folder_root=[r"\\...\UNIT\DME\Logs"],
#     extract_keyword=["real time"],
#     output_file_path=r"\\...\data\scan_metrics.xlsx",
#     metric_profile="sas_log")


# --- Example 4: formatCSV -- corporate styled Excel report -------------------
# cgs_ai.formatCSV(InputCsvPath=r"\\...\data\scan_test.csv",
#                  OutputExcelPath=r"\\...\data\scan_report.xlsx",
#                  FormatType="corporate")


# --- Example 5: downloadBulkFiles -- CMS comment attachments -----------------
# cgs_ai.downloadBulkFiles(InputCsvPath=r"\\...\data\comments.csv",
#                          OutputFolder=r"\\...\data\attachments",
#                          LinkColumn="attachmentLinks")


# --- Example 6: sendEmail ----------------------------------------------------
# cgs_ai.sendEmail(To="ops@example.com;analyst@example.com",
#                  From="cgs_ai@example.com",
#                  Subject="Scan complete", Body="The nightly scan finished.")


# --- Example 7: runSQLServerQuery -------------------------------------------
# out = cgs_ai.runSQLServerQuery(SQL_Statement="select top 100 * from dbo.Claims",
#                                LOB_Catalog="DataMartKYA")
# print(out["RowCount"], "rows")


# --- Example 8: convertSAS2Pandas -------------------------------------------
# out = cgs_ai.convertSAS2Pandas(InputSas7bdatPath=r"\\...\claims.sas7bdat",
#                                OutputPath=r"\\...\claims.csv")


# --- Example 9: copyExcelSheet2CSV ------------------------------------------
# cgs_ai.copyExcelSheet2CSV(InputExcelPath=r"\\...\reference.xlsx",
#                           SheetName="DenialCodes",
#                           OutputCsvPath=r"\\...\denial_codes.csv")


# --- Example 10: collectSystemMetrics ---------------------------------------
# cgs_ai.collectSystemMetrics(OutputCsvPath=r"\\...\logs\host_metrics.csv",
#                             WriteMode="append")


# --- Example 11: zipFolder ---------------------------------------------------
# cgs_ai.zipFolder(FolderToZip=r"\\...\cgs_ai\data",
#                  OutputZipPath=r"\\...\cgs_ai\data_archive.zip",
#                  AccompanyFiles=[r"\\...\README.docx"])


# --- Example 12: the end-to-end pipeline ------------------------------------
# cgs_ai.runFilescanPipeline(
#     extract_keyword=["real time", "cpu time"],
#     metric_profile="sas_log",
#     email_to="ops@example.com", email_from="cgs_ai@example.com")
