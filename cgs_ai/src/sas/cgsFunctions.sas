/*=====================================================================
  Program Name : cgsFunctions.sas
  Author       : Manuel Figallo
  Purpose      : SAS wrapper macros for every cgs_ai function. Each macro
                 mirrors the Python and PowerShell parameter names exactly,
                 so the same call works in all three languages.
  Version      : 1.0beta
  Created      : 2026-08-26

  DESIGN
    The SAS folder contains WRAPPER CODE ONLY. Every macro here builds an
    argument list and delegates to %cgsRun in cgsCore.sas, which launches
    either the PowerShell (engine=ps) or the Python (engine=py) twin.

    Every macro takes engine= so you can switch implementations without
    changing any other argument:
        %scanFileSystem(..., engine=ps);   /* runs src\ps\scanFileSystem.ps1 */
        %scanFileSystem(..., engine=py);   /* runs src\py\scanFileSystem.py  */

    List parameters take SEMICOLON-DELIMITED strings wrapped in %str(),
    which the target language splits back into a list.

  Usage:
      %include "<path>\src\sas\cgsCore.sas";
      %include "<path>\src\sas\cgsFunctions.sas";
=====================================================================*/


%macro scanFileSystem(
    input_folder_root      =,      /* REQUIRED; semicolon-delimited      */
    extract_keyword        =,      /* REQUIRED; semicolon-delimited      */
    output_file_path       =,      /* .csv or .xlsx, or a directory      */
    file_extensions        =,      /* semicolon-delimited                */
    include_subdirectories = 1,    /* 1 = recurse, 0 = top level only    */
    folder_exclusion_list  =,      /* semicolon-delimited; default none  */
    file_exclusion_list    =,      /* semicolon-delimited                */
    lines_above            = 5,    /* context lines captured BEFORE      */
    lines_below            = 5,    /* context lines captured AFTER       */
    nth_token_after        = 1,    /* which token after the keyword      */
    nth_token_before       = 1,    /* which token before the keyword     */
    numeric_token_after    = 1,    /* which NUMERIC token after          */
    date_from              =,      /* inclusive YYYY-MM-DD               */
    date_to                =,      /* inclusive YYYY-MM-DD               */
    date_field             =,      /* created | modified | accessed      */
    metric_profile         =,      /* none | sas_log -> EXCEL output     */
    engine                 = ps,   /* ps | py                            */
    debug                  = 0
);
/* Scan directory roots for keyword matches; one output row per MATCH.
   Emits SourceDir, FileName, Line, LinesAbove, LinesBelow, FullPath,
   LineNumber, Keyword, ExtractedString, NthTokenAfter, NthTokenBefore,
   NumericTokenAfter, LastToken, FirstToken, FileTimestamp, extension,
   file_size_bytes, created_time, modified_time, accessed_time, scanned_at.
   When metric_profile is set, EXCEL output is produced (announced in the
   log) with a second sheet of structured metrics.
   Use in claims processing: sweep SAS/ETL logs for ERROR, a claim number or
   a file reference and get the matched line plus context in one table.     */
  %cgsResetArgs;
  %local p;
  %if &engine = ps %then %let p = -; %else %let p = --;
  %if &engine = ps %then %do;
    %cgsAddArg(name=-input_folder_root,      value=%superq(input_folder_root));
    %cgsAddArg(name=-extract_keyword,        value=%superq(extract_keyword));
    %cgsAddArg(name=-output_file_path,       value=%superq(output_file_path));
    %cgsAddArg(name=-file_extensions,        value=%superq(file_extensions));
    %cgsAddArg(name=-include_subdirectories, value=&include_subdirectories, always=1);
    %cgsAddArg(name=-folder_exclusion_list,  value=%superq(folder_exclusion_list));
    %cgsAddArg(name=-file_exclusion_list,    value=%superq(file_exclusion_list));
    %cgsAddArg(name=-lines_above,            value=&lines_above, always=1);
    %cgsAddArg(name=-lines_below,            value=&lines_below, always=1);
    %cgsAddArg(name=-nth_token_after,        value=&nth_token_after, always=1);
    %cgsAddArg(name=-nth_token_before,       value=&nth_token_before, always=1);
    %cgsAddArg(name=-numeric_token_after,    value=&numeric_token_after, always=1);
    %cgsAddArg(name=-date_from,              value=%superq(date_from));
    %cgsAddArg(name=-date_to,                value=%superq(date_to));
    %cgsAddArg(name=-date_field,             value=%superq(date_field));
    %cgsAddArg(name=-metric_profile,         value=%superq(metric_profile));
    %cgsRun(engine=ps, script=scanFileSystem.ps1, taskname=scanps, debug=&debug);
  %end;
  %else %do;
    %cgsAddArg(name=--input-folder-root,   value=%superq(input_folder_root));
    %cgsAddArg(name=--extract-keyword,     value=%superq(extract_keyword));
    %cgsAddArg(name=--output-file-path,    value=%superq(output_file_path));
    %cgsAddArg(name=--file-extensions,     value=%superq(file_extensions));
    %if &include_subdirectories = 0 %then
      %cgsAddArg(name=--no-include-subdirectories, value=, always=1);
    %cgsAddArg(name=--folder-exclusion-list, value=%superq(folder_exclusion_list));
    %cgsAddArg(name=--file-exclusion-list,   value=%superq(file_exclusion_list));
    %cgsAddArg(name=--lines-above,           value=&lines_above, always=1);
    %cgsAddArg(name=--lines-below,           value=&lines_below, always=1);
    %cgsAddArg(name=--nth-token-after,       value=&nth_token_after, always=1);
    %cgsAddArg(name=--nth-token-before,      value=&nth_token_before, always=1);
    %cgsAddArg(name=--numeric-token-after,   value=&numeric_token_after, always=1);
    %cgsAddArg(name=--date-from,             value=%superq(date_from));
    %cgsAddArg(name=--date-to,               value=%superq(date_to));
    %cgsAddArg(name=--date-field,            value=%superq(date_field));
    %cgsAddArg(name=--metric-profile,        value=%superq(metric_profile));
    %cgsRun(engine=py, script=scanFileSystem.py, taskname=scanpy, debug=&debug);
  %end;
%mend scanFileSystem;


%macro runSQLServerQuery(
    SQL_Statement =,               /* REQUIRED; the query text            */
    LOB_Catalog   =,               /* REQUIRED; e.g. DataMartKYA          */
    DataSource    =,               /* server,port                         */
    OutputCsvPath =,               /* optional CSV for the result set     */
    engine        = ps,            /* ps needs NO module; py needs pyodbc */
    debug         = 0
);
/* Run a SQL Server query using Windows Integrated Security -- the direct
   equivalent of the PROC SQL / connect to oledb pass-through block. No
   password is handled, stored or logged.
   Use in claims processing: pull a claims or denial extract from the LOB
   data mart into CSV without a manual export step.                        */
  %cgsResetArgs;
  %cgsAddArg(name=-SQL_Statement, value=%superq(SQL_Statement));
  %cgsAddArg(name=-LOB_Catalog,   value=%superq(LOB_Catalog));
  %cgsAddArg(name=-DataSource,    value=%superq(DataSource));
  %cgsAddArg(name=-OutputCsvPath, value=%superq(OutputCsvPath));
  %cgsRun(engine=&engine, script=runSQLServerQuery.%sysfunc(ifc(&engine=ps,ps1,py)),
          taskname=cgssql, debug=&debug);
%mend runSQLServerQuery;


%macro formatCSV(
    InputCsvPath    =,             /* REQUIRED                            */
    OutputExcelPath =,             /* REQUIRED                            */
    FormatType      = corporate,   /* corporate | corporatev2 | plain |   */
                                   /* minimal | ODS1                      */
    SheetName       = Report,
    Title           =,
    engine          = ps,
    debug           = 0
);
/* Render a CSV as a styled Excel workbook with a SAS ODS look and feel:
   navy banner, blue header row, zebra striping.

   FormatType=ODS1 is the ODD ONE OUT. It produces the same look as
   "corporate" but renders it with ODS EXCEL INSIDE SAS, so it needs no
   PowerShell module and no Python. Use it on a server where the
   ImportExcel module is unavailable. engine= is ignored for ODS1 because
   nothing is launched outside SAS.

   Use in claims processing: turn a raw scan or claims extract into a report
   an analyst can open directly, with no hand-formatting each cycle.       */

  %if %upcase(&FormatType) = ODS1 %then %do;
      %cgsFormatCsvOds(InputCsvPath=%superq(InputCsvPath),
                       OutputExcelPath=%superq(OutputExcelPath),
                       SheetName=&SheetName, Title=%superq(Title),
                       debug=&debug);
  %end;
  %else %do;
      %cgsResetArgs;
      %cgsAddArg(name=-InputCsvPath,    value=%superq(InputCsvPath));
      %cgsAddArg(name=-OutputExcelPath, value=%superq(OutputExcelPath));
      %cgsAddArg(name=-FormatType,      value=&FormatType);
      %cgsAddArg(name=-SheetName,       value=&SheetName);
      %cgsAddArg(name=-Title,           value=%superq(Title));
      %cgsRun(engine=&engine, script=formatCSV.%sysfunc(ifc(&engine=ps,ps1,py)),
              taskname=cgsfmt, debug=&debug);
  %end;
%mend formatCSV;


%macro cgsFormatCsvOds(
    InputCsvPath    =,             /* REQUIRED                            */
    OutputExcelPath =,             /* REQUIRED                            */
    SheetName       = Report,
    Title           =,
    debug           = 0
);
/* The ODS1 renderer behind %formatCSV(FormatType=ODS1).

   Reads the CSV with PROC IMPORT and writes the workbook with ODS EXCEL,
   so the whole job stays inside SAS. Colours match FORMAT_STYLES
   "corporate" in the Python twin: banner 1F3864, header 2E75B6, stripe
   DCE6F1.

   NOTE: the banded rows come from a NOPRINT counter column and a compute
   block. PROC REPORT evaluates compute blocks in column order, so the
   counter must be the FIRST item in the COLUMN statement for CALL DEFINE
   to colour the row before it is written.                                */

  %local _cgsVars _cgsBanner _cgsRc;
  %let _cgsRc = 0;

  %if %superq(InputCsvPath) = %str() %then %do;
      %put ERROR: required parameter InputCsvPath is missing or empty.;
      %return;
  %end;
  %if %superq(OutputExcelPath) = %str() %then %do;
      %put ERROR: required parameter OutputExcelPath is missing or empty.;
      %return;
  %end;
  %if %sysfunc(fileexist(%superq(InputCsvPath))) = 0 %then %do;
      %put ERROR: input CSV not found: %superq(InputCsvPath);
      %return;
  %end;

  /* Banner text defaults to the CSV file name, matching the other twins. */
  %if %superq(Title) = %str() %then
      %let _cgsBanner = %scan(%superq(InputCsvPath), -2, %str(\./));
  %else %let _cgsBanner = %superq(Title);

  proc import datafile="%superq(InputCsvPath)" out=work._cgsFmt
              dbms=csv replace;
      getnames=yes;
      guessingrows=max;
  run;

  %if &syserr > 4 %then %do;
      %put ERROR: could not read %superq(InputCsvPath);
      %return;
  %end;

  /* A row counter drives the zebra striping. */
  data work._cgsFmt;
      set work._cgsFmt;
      _cgsRow_ = _n_;
  run;

  /* Every column except the counter, in position order. */
  proc sql noprint;
      select name into :_cgsVars separated by ' '
      from dictionary.columns
      where libname = 'WORK' and memname = '_CGSFMT'
        and upcase(name) ne '_CGSROW_'
      order by varnum;
  quit;

  ods escapechar='^';
  ods excel file="%superq(OutputExcelPath)"
      options(sheet_name="&SheetName"
              embedded_titles="yes"
              frozen_headers="on"
              autofilter="all"
              flow="tables");

  title j=left
      "^S={background=cx1F3864 foreground=cxFFFFFF fontsize=14pt fontweight=bold}&_cgsBanner";

  proc report data=work._cgsFmt nowd missing
       style(report)=[rules=all frame=box cellspacing=0]
       style(header)=[background=cx2E75B6 foreground=cxFFFFFF fontweight=bold
                      vjust=center just=left];
      column _cgsRow_ &_cgsVars;
      define _cgsRow_ / display noprint;
      compute _cgsRow_;
          if mod(_cgsRow_, 2) = 1 then
              call define(_row_, 'style', 'style=[background=cxDCE6F1]');
      endcomp;
  run;

  title;
  ods excel close;

  %if &debug = 0 %then %do;
      proc datasets library=work nolist nowarn;
          delete _cgsFmt;
      quit;
  %end;

  %put NOTE: formatCSV [ODS1] wrote %superq(OutputExcelPath);
%mend cgsFormatCsvOds;


%macro downloadBulkFiles(
    InputCsvPath =,                /* REQUIRED                            */
    OutputFolder =,                /* REQUIRED                            */
    LinkColumn   = attachmentLinks,
    IdColumn     = commentId,
    Overwrite    = false,
    engine       = ps,
    debug        = 0
);
/* Download every file listed in a CSV link column. Cells may be BLANK, a
   single URL, or several joined by '|'. Blank cells are skipped, a failed
   download is logged and the run continues.
   Use in claims processing: pull every public-comment attachment for a CMS
   docket into one folder for bulk text extraction and review.             */
  %cgsResetArgs;
  %cgsAddArg(name=-InputCsvPath, value=%superq(InputCsvPath));
  %cgsAddArg(name=-OutputFolder, value=%superq(OutputFolder));
  %cgsAddArg(name=-LinkColumn,   value=&LinkColumn);
  %cgsAddArg(name=-IdColumn,     value=&IdColumn);
  %cgsAddArg(name=-Overwrite,    value=&Overwrite, always=1);
  %cgsRun(engine=&engine, script=downloadBulkFiles.%sysfunc(ifc(&engine=ps,ps1,py)),
          taskname=cgsdl, debug=&debug);
%mend downloadBulkFiles;


%macro sendEmail(
    To         =,                  /* REQUIRED; semicolon-delimited       */
    From       =,                  /* REQUIRED                            */
    Subject    =,                  /* REQUIRED                            */
    Body       =,                  /* REQUIRED                            */
    SmtpServer =,                  /* default smtp.example.com            */
    Port       =,                  /* default 25                          */
    engine     = ps,
    debug      = 0
);
/* Send an email alert over SMTP. Multiple recipients allowed in To.
   Use in claims processing: notify the operations mailbox when an overnight
   scan or bulk download finishes, including row counts and output paths.  */
  %cgsResetArgs;
  %cgsAddArg(name=-To,         value=%superq(To));
  %cgsAddArg(name=-From,       value=%superq(From));
  %cgsAddArg(name=-Subject,    value=%superq(Subject));
  %cgsAddArg(name=-Body,       value=%superq(Body));
  %cgsAddArg(name=-SmtpServer, value=%superq(SmtpServer));
  %cgsAddArg(name=-Port,       value=%superq(Port));
  %cgsRun(engine=&engine, script=sendEmail.%sysfunc(ifc(&engine=ps,ps1,py)),
          taskname=cgsmail, debug=&debug);
%mend sendEmail;


%macro convertSAS2Pandas(
    InputSas7bdatPath =,           /* REQUIRED                            */
    OutputPath        =,           /* REQUIRED; .parquet | .csv | .pkl    */
    engine            = py,        /* py is native; ps delegates to py    */
    debug             = 0
);
/* Convert a sas7bdat data set to a pandas-readable file.
   NOTE: this is the one function that cannot be standard-library only --
   sas7bdat is a proprietary binary format, so pandas is required.
   Use in claims processing: bring a SAS claims extract into the Python or
   Snowflake side of the stack without a manual export.                    */
  %cgsResetArgs;
  %cgsAddArg(name=-InputSas7bdatPath, value=%superq(InputSas7bdatPath));
  %cgsAddArg(name=-OutputPath,        value=%superq(OutputPath));
  %cgsRun(engine=&engine, script=convertSAS2Pandas.%sysfunc(ifc(&engine=ps,ps1,py)),
          taskname=cgssas2pd, debug=&debug);
%mend convertSAS2Pandas;


%macro copyExcelSheet2CSV(
    InputExcelPath =,              /* REQUIRED                            */
    SheetName      =,              /* REQUIRED                            */
    OutputCsvPath  =,              /* REQUIRED                            */
    HeaderRow      = 1,            /* 2 for formatCSV output (banner row) */
    engine         = ps,
    debug          = 0
);
/* Export one Excel worksheet to CSV, validating FIRST and stopping if the
   sheet is not shaped for flat output (missing, empty, blank or duplicate
   headers, merged header cells).
   Use in claims processing: convert a hand-maintained reference workbook
   (fee schedules, denial-code mappings) into pipeline-ready CSV.          */
  %cgsResetArgs;
  %cgsAddArg(name=-InputExcelPath, value=%superq(InputExcelPath));
  %cgsAddArg(name=-SheetName,      value=%superq(SheetName));
  %cgsAddArg(name=-OutputCsvPath,  value=%superq(OutputCsvPath));
  %cgsAddArg(name=-HeaderRow,      value=&HeaderRow, always=1);
  %cgsRun(engine=&engine, script=copyExcelSheet2CSV.%sysfunc(ifc(&engine=ps,ps1,py)),
          taskname=cgsxls2csv, debug=&debug);
%mend copyExcelSheet2CSV;


%macro collectSystemMetrics(
    OutputCsvPath =,               /* REQUIRED                            */
    ServerName    =,               /* defaults to this host               */
    WriteMode     = append,        /* append | overwrite                  */
    engine        = ps,
    debug         = 0
);
/* Gather host metrics into a CSV time series, failing gracefully: a metric
   unavailable on this server is recorded blank and the run still succeeds.
   Use in claims processing: sample the SAS/ETL server during a nightly run
   to correlate slow steps with CPU, memory or disk pressure.              */
  %cgsResetArgs;
  %cgsAddArg(name=-OutputCsvPath, value=%superq(OutputCsvPath));
  %cgsAddArg(name=-ServerName,    value=%superq(ServerName));
  %cgsAddArg(name=-WriteMode,     value=&WriteMode, always=1);
  %cgsRun(engine=&engine, script=collectSystemMetrics.%sysfunc(ifc(&engine=ps,ps1,py)),
          taskname=cgsmetrics, debug=&debug);
%mend collectSystemMetrics;


%macro zipFolder(
    FolderToZip    =,              /* REQUIRED                            */
    OutputZipPath  =,              /* REQUIRED; full path of the .zip     */
    AccompanyFiles =,              /* semicolon-delimited extra files     */
    engine         = ps,
    debug          = 0
);
/* Archive a folder plus a list of accompanying files into one .zip.
   Use in claims processing: bundle a month of scan output, the formatted
   Excel report and the run log for records retention or hand-off.         */
  %cgsResetArgs;
  %cgsAddArg(name=-FolderToZip,    value=%superq(FolderToZip));
  %cgsAddArg(name=-OutputZipPath,  value=%superq(OutputZipPath));
  %cgsAddArg(name=-AccompanyFiles, value=%superq(AccompanyFiles));
  %cgsRun(engine=&engine, script=zipFolder.%sysfunc(ifc(&engine=ps,ps1,py)),
          taskname=cgszip, debug=&debug);
%mend zipFolder;


%macro runFilescanPipeline(
    input_folder_root =,           /* semicolon-delimited                 */
    extract_keyword   =,           /* semicolon-delimited                 */
    output_file_path  =,
    excel_output_path =,
    metric_profile    = sas_log,
    email_to          =,
    email_from        =,
    email_subject     =,
    engine            = ps,
    debug             = 0
);
/* End-to-end pipeline: scanFileSystem -> formatCSV -> sendEmail.
   Use in claims processing: nightly sweep of SAS job logs for timing and
   error keywords, delivered to the operations mailbox as a workbook.      */
  %cgsResetArgs;
  %cgsAddArg(name=-input_folder_root, value=%superq(input_folder_root));
  %cgsAddArg(name=-extract_keyword,   value=%superq(extract_keyword));
  %cgsAddArg(name=-output_file_path,  value=%superq(output_file_path));
  %cgsAddArg(name=-excel_output_path, value=%superq(excel_output_path));
  %cgsAddArg(name=-metric_profile,    value=&metric_profile);
  %cgsAddArg(name=-email_to,          value=%superq(email_to));
  %cgsAddArg(name=-email_from,        value=%superq(email_from));
  %cgsAddArg(name=-email_subject,     value=%superq(email_subject));
  %cgsRun(engine=&engine, script=filescan_pipeline.%sysfunc(ifc(&engine=ps,ps1,py)),
          taskname=cgspipe, debug=&debug);
%mend runFilescanPipeline;
