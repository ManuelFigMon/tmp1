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
    changing any other argument. engine=ps runs the .ps1 in src\ps and
    engine=py runs the .py in src\py:
        %scanFileSystem(..., engine=ps);
        %scanFileSystem(..., engine=py);

    NOTE: never open a block comment inside this header. SAS comments do not
    nest, so the header would end at that inner close marker and everything
    below it would be parsed as live code.

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
    ColumnLength    = 1024,        /* per-column character width          */
    debug           = 0
);
/* The ODS1 renderer behind %formatCSV(FormatType=ODS1).

   Reads the CSV with a DATA step and writes the workbook with ODS EXCEL,
   so the whole job stays inside SAS. Colours match FORMAT_STYLES
   "corporate" in the Python twin: banner 1F3864, header 2E75B6, stripe
   DCE6F1.

   WHY NOT PROC IMPORT: it samples the file to guess column types, and a
   CSV with a header but no data rows -- which is exactly what
   scanFileSystem writes when nothing matches -- fails with "Unable to
   sample external file, no data in first 5 records. ERROR: Import
   unsuccessful." A DATA step does not sample, so a header-only file
   produces a workbook with headers and no rows, matching the Python twin.

   Every column is read as CHARACTER, which is also what the Python and
   PowerShell twins write. That keeps leading zeros on identifiers such as
   contractor numbers instead of silently turning them into numbers.

   NOTE: the banded rows come from a NOPRINT counter column and a compute
   block. PROC REPORT evaluates compute blocks in column order, so the
   counter must be the FIRST item in the COLUMN statement for CALL DEFINE
   to colour the row before it is written.

   LIMITATION: a column heading containing & or % will be resolved as a
   macro reference when the labels are applied. Rename such a column in the
   source CSV, or use engine=ps / engine=py instead.                      */

  %local _cgsBanner _cgsHdr _cgsN _cgsLabels _cgsVarList _cgsRows;

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

  /* Read the header line only. */
  %let _cgsHdr = ;
  data _null_;
      infile "%superq(InputCsvPath)" lrecl=32767 truncover obs=1;
      input line $char32767.;
      call symputx('_cgsHdr', line, 'L');
  run;

  %if %superq(_cgsHdr) = %str() %then %do;
      %put ERROR: the CSV has no header row (the file is empty): %superq(InputCsvPath);
      %return;
  %end;

  /* Turn the header into column count and labels. Columns are named
     _c1.._cN so that no header text can produce an invalid or duplicate
     SAS name; the original text is carried as the label, which is what
     PROC REPORT prints. */
  data _null_;
      length hdr $ 32767 lab $ 1000 labels $ 32767;
      hdr = symget('_cgsHdr');
      /* Drop a UTF-8 byte-order mark; the Python twin reads utf-8-sig. */
      if substr(hdr, 1, 3) = 'EFBBBF'x then hdr = substr(hdr, 4);
      n = countw(hdr, ',', 'mq');
      do i = 1 to n;
          lab = dequote(strip(scan(hdr, i, ',', 'mq')));
          labels = catx(' ', labels, cats('_c', i, '=') || quote(trim(lab)));
      end;
      call symputx('_cgsN', n, 'L');
      call symputx('_cgsLabels', labels, 'L');
  run;

  /* A single column cannot be written as the range _c1-_c1. */
  %if &_cgsN = 1 %then %let _cgsVarList = _c1;
  %else %let _cgsVarList = _c1-_c&_cgsN;

  data work._cgsFmt;
      length &_cgsVarList $ &ColumnLength;
      infile "%superq(InputCsvPath)" dsd dlm=',' lrecl=32767 truncover
             firstobs=2;
      input &_cgsVarList;
      _cgsRow_ = _n_;
      label &_cgsLabels;
  run;

  %if &syserr > 4 %then %do;
      %put ERROR: could not read %superq(InputCsvPath);
      %return;
  %end;

  %local _cgsDsid _cgsRc;
  %let _cgsDsid = %sysfunc(open(work._cgsFmt));
  %let _cgsRows = %sysfunc(attrn(&_cgsDsid, nlobs));
  %let _cgsRc   = %sysfunc(close(&_cgsDsid));

  %if &_cgsRows = 0 %then %do;
      %put NOTE: the CSV has a header but no data rows - the workbook will hold the banner only.;
  %end;

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
      column _cgsRow_ &_cgsVarList;
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
