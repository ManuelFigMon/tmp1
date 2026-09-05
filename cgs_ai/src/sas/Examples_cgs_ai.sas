/*=====================================================================
  Program Name : Examples_cgs_ai.sas
  Author       : Manuel Figallo
  Purpose      : Example calls for every cgs_ai SAS wrapper macro.
  Version      : 1.0beta
  Created      : 2026-08-26

  HOW TO USE
    1. Edit the two %INCLUDE paths below.
    2. Optionally override PS_FOLDER_PATH / PYTHON_FOLDER_PATH.
    3. Un-comment ONE example and submit.

  Every example is commented out (%* ... ;) so submitting the whole file
  does nothing by accident. Add debug=1 to any call to build and print the
  command WITHOUT executing it -- the safe way to try a new call.
=====================================================================*/

%include "\\a70admed.com\R1\CGS\APPS\SAS\UNIT\SAS_G\GSIT_Prod\MANUAL\cgs_ai\src\sas\cgsCore.sas";
%include "\\a70admed.com\R1\CGS\APPS\SAS\UNIT\SAS_G\GSIT_Prod\MANUAL\cgs_ai\src\sas\cgsFunctions.sas";

/* Optional per-run overrides */
%*let PS_FOLDER_PATH     = \\a70admed.com\R1\...\cgs_ai\src\ps\;
%*let PYTHON_FOLDER_PATH = \\a70admed.com\R1\...\cgs_ai\src\py\;
%*let PYTHON_EXE         = C:\code\python\cgs_ai\.venv\Scripts\python.exe;


/* --- A. scanFileSystem: HHH Old_logs + DME Logs, metrics to Excel ------- */
%*
%scanFileSystem(
  input_folder_root=%str(\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT\HHH\Old_Programs\Old_logs;\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT\DME\Logs),
  extract_keyword=%str(real time;cpu time),
  output_file_path=\\a70admed.com\R1\CGS\APPS\SAS\UNIT\SAS_G\GSIT_Prod\MANUAL\cgs_ai\data\scan_ps_test5.csv,
  metric_profile=sas_log
);
*;

/* --- B. Same scan, wider context and the 2nd numeric token -------------- */
%*
%scanFileSystem(
  input_folder_root=%str(\\A70admed.com\r1\...\UNIT\DME\Logs),
  extract_keyword=%str(ERROR;WARNING),
  output_file_path=\\...\cgs_ai\data\errors.csv,
  lines_above=10, lines_below=10, numeric_token_after=2
);
*;

/* --- C. Same call, Python engine instead of PowerShell ------------------ */
%*
%scanFileSystem(
  input_folder_root=%str(\\A70admed.com\r1\...\UNIT\DME\Logs),
  extract_keyword=%str(real time),
  output_file_path=\\...\cgs_ai\data\scan_py.csv,
  engine=py
);
*;

/* --- D. runSQLServerQuery ---------------------------------------------- */
%*
%runSQLServerQuery(
  SQL_Statement=%str(select top 100 * from dbo.Claims),
  LOB_Catalog=DataMartKYA,
  OutputCsvPath=\\...\cgs_ai\data\claims_top100.csv
);
*;

/* --- E. formatCSV: styled corporate report ----------------------------- */
%*
%formatCSV(
  InputCsvPath=\\...\cgs_ai\data\scan_ps_test5.csv,
  OutputExcelPath=\\...\cgs_ai\data\scan_report.xlsx,
  FormatType=corporate
);
*;

/* --- F. downloadBulkFiles: CMS comment attachments --------------------- */
%*
%downloadBulkFiles(
  InputCsvPath=\\...\cgs_ai\data\comments.csv,
  OutputFolder=\\...\cgs_ai\data\attachments,
  LinkColumn=attachmentLinks
);
*;

/* --- G. sendEmail ------------------------------------------------------ */
%*
%sendEmail(
  To=%str(ops@example.com;analyst@example.com),
  From=cgs_ai@example.com,
  Subject=%str(Scan complete),
  Body=%str(The nightly scan finished successfully.)
);
*;

/* --- H. convertSAS2Pandas ---------------------------------------------- */
%*
%convertSAS2Pandas(
  InputSas7bdatPath=\\...\data\claims.sas7bdat,
  OutputPath=\\...\data\claims.csv
);
*;

/* --- I. copyExcelSheet2CSV (HeaderRow=2 for formatCSV output) ----------- */
%*
%copyExcelSheet2CSV(
  InputExcelPath=\\...\cgs_ai\data\scan_report.xlsx,
  SheetName=Report,
  OutputCsvPath=\\...\cgs_ai\data\scan_report.csv,
  HeaderRow=2
);
*;

/* --- J. collectSystemMetrics ------------------------------------------- */
%*
%collectSystemMetrics(
  OutputCsvPath=\\...\cgs_ai\logs\host_metrics.csv,
  WriteMode=append
);
*;

/* --- K. zipFolder ------------------------------------------------------ */
%*
%zipFolder(
  FolderToZip=\\...\cgs_ai\data,
  OutputZipPath=\\...\cgs_ai\data_archive.zip,
  AccompanyFiles=%str(\\...\cgs_ai\README.docx)
);
*;

/* --- L. The end-to-end pipeline ---------------------------------------- */
%*
%runFilescanPipeline(
  extract_keyword=%str(real time;cpu time),
  metric_profile=sas_log,
  email_to=%str(ops@example.com),
  email_from=cgs_ai@example.com
);
*;

/* --- M. DEBUG: build and print the command WITHOUT running it ----------- */
%*
%scanFileSystem(
  input_folder_root=%str(\\A70admed.com\r1\...\UNIT\DME\Logs),
  extract_keyword=%str(real time),
  output_file_path=\\...\cgs_ai\data\scan.csv,
  metric_profile=sas_log,
  debug=1
);
*;
