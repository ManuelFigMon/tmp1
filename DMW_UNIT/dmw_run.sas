/*=============================================================================
  dmw_run.sas
  -----------------------------------------------------------------------------
  DMW (Data Mart Workbench) -- IT Management Solution
  SAS Stored Process middleware: executes a SQL statement supplied by the
  HTML/JavaScript page and streams the result set back as a JSON array.

  Renamed from dmq_run to dmw_run. Behaviour is unchanged except that the
  database connection is now driven by the configuration block below rather
  than being hard-coded inside PROC SQL.

  -----------------------------------------------------------------------------
  HOW THE PIECES FIT TOGETHER
  -----------------------------------------------------------------------------
      Browser (dmw_it_management.html)
            |
            |  HTTPS GET  .../SASStoredProcess/do?_program=<path>
            |                 &dmq_query=<URL-encoded SQL>
            v
      SAS Stored Process Web Application  -> THIS PROGRAM
            |
            |  SAS/ACCESS to OLE DB, explicit pass-through, TDS port 1433
            v
      SQL Server  A70TUCGSDATA012.A70ADMED.COM / OHM_CGSScheduler

  A browser cannot open a TDS socket to SQL Server directly, so this stored
  process is the bridge. It speaks HTTP on the front and the database on the
  back.

  -----------------------------------------------------------------------------
  REGISTRATION IN SAS MANAGEMENT CONSOLE
  -----------------------------------------------------------------------------
    1. New Stored Process, name: dmw_run
    2. Source code location : the folder holding this .sas file
    3. Output type          : Streaming
    4. Result type          : Streaming      <-- REQUIRED. If this is left at
                              the default, SAS wraps the output in ODS HTML and
                              the browser receives a web page where it expects
                              JSON.
    5. Input parameters     : dmq_query  (Text, not required at prompt)
                              debug_mode (Text, not required at prompt)
    6. Note the metadata path, e.g.
         /User Folders/731o/My Folder/dmw_run
       and enter it as SAS_PROGRAM in the HTML page's CONFIG block.

  -----------------------------------------------------------------------------
  INPUT PARAMETERS (supplied on the query string by the HTML page)
  -----------------------------------------------------------------------------
    dmq_query   The SQL SELECT statement to execute. Retained under this name
                so the program stays drop-in compatible with the existing DMQ
                pages that already post to this parameter.
    debug_mode  0 (default) or 1. When 1, the statement, row count and elapsed
                time are written to the SAS log. Never to the HTTP response.

  -----------------------------------------------------------------------------
  OUTPUT
  -----------------------------------------------------------------------------
    Content-type: application/json
    A bare JSON array of row objects, for example:
        [ { "RequestID":1, "JobName":"NIGHTLY_LOAD", ... }, ... ]
    An empty array [] is returned when the statement contains no SELECT.

  -----------------------------------------------------------------------------
  SECURITY NOTES -- READ BEFORE DEPLOYING
  -----------------------------------------------------------------------------
    * This program executes SQL supplied over HTTP. The SELECT check below is a
      guard, not a sandbox. Grant the stored process identity only the database
      permissions it genuinely needs -- read access to the target view.
    * With Integrated Security the connection runs as the SAS server process
      account (A70ADMED\731o in UNIT). That account's permissions are the real
      security boundary.
    * Credentials are never placed in this file. SSPI uses the process account.
    * Consider tightening DMW_MAXROWS below if the page should not be able to
      pull unbounded result sets.
=============================================================================*/


/*-----------------------------------------------------------------------------
  CONFIGURATION -- edit these to match your environment
  -----------------------------------------------------------------------------
  These values must agree with the CONFIG block at the top of the HTML page for
  the two halves to describe the same target.
-----------------------------------------------------------------------------*/

/* Target SQL Server host or fully qualified domain name. */
%let DMW_SERVER   = A70TUCGSDATA012.A70ADMED.COM;

/* Listening port. SQL Server defaults to 1433. */
%let DMW_PORT     = 1433;

/* Default database (Initial Catalog) the connection attaches to. */
%let DMW_CATALOG  = OHM_CGSScheduler;

/* OLE DB provider installed on the SAS server.
     MSOLEDBSQL19  current Microsoft OLE DB Driver 19
     MSOLEDBSQL    Microsoft OLE DB Driver 18
     SQLNCLI11     legacy SQL Server Native Client
   If a connection fails with "provider not registered", try the next one down. */
%let DMW_PROVIDER = MSOLEDBSQL19;

/* Authentication.
     1 = Windows Authentication (Integrated Security = SSPI). The connection is
         made as the SAS server process account, which in UNIT is A70ADMED\731o.
         No credentials are stored anywhere.
     0 = SQL Server authentication. Supply DMW_USER and DMW_PASS below; prefer a
         SAS authdomain or metadata-stored credentials over literals here. */
%let DMW_TRUSTED  = 1;
%let DMW_USER     = ;
%let DMW_PASS     = ;

/* Maximum rows streamed back to the browser. Guards against an unbounded
   result set flooding the page. Set to 0 for no cap. */
%let DMW_MAXROWS  = 5000;

/* Folder for the server-side execution log. Must be writable by the SAS server
   process account. Set to an empty value to disable logging. */
%let DMW_LOGFOLDER = \\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\Web_Stored_Processes\Logs;

/*---------------------------------------------------------------------------*/


%global dmq_query query debug_mode;

%macro dmw_run();

    /*--------------------------------------------------------------
      Set default debug_mode if not passed via URL
    --------------------------------------------------------------*/
    %if %symexist(debug_mode)=0 %then %let debug_mode=0;
    %if %superq(debug_mode)=     %then %let debug_mode=0;

    /*--------------------------------------------------------------
      Start execution timer
    --------------------------------------------------------------*/
    %let _start_time = %sysfunc(datetime());

    /*--------------------------------------------------------------
      Clean query string safely and remove trailing semicolons
      (Safe for SQL containing commas, quotes, parentheses)
    --------------------------------------------------------------*/
    data _null_;
        length q $32767;
        q = symget('dmq_query');            /* get query from URL parameter */
        q = strip(q);
        do while(length(q) > 0 and substr(q,length(q),1)=';'); /* trailing ; */
            q = substr(q,1,length(q)-1);
        end;
        call symputx('query',q);            /* write back to macro variable */
    run;

    /* Optional debug print */
    %if &debug_mode=1 %then %put NOTE: [DEBUG] Query=&query;

    /*--------------------------------------------------------------
      Set JSON header for web output
    --------------------------------------------------------------*/
    data _null_;
        rc = stpsrv_header('Content-type','application/json');
        rc = stpsrv_header('Pragma','no-cache');
        rc = stpsrv_header('Cache-Control','no-store, no-cache, must-revalidate');
    run;

    /*--------------------------------------------------------------
      Validate query contains SELECT safely
    --------------------------------------------------------------*/
    %local _select_found;
    %let _select_found = 0;
    data _null_;
        if prxmatch("/\bSELECT\b/i", symget("query")) then call symputx('_select_found',1);
    run;

    %if &_select_found=0 %then %do;
        data _null_;
            file _webout lrecl=32767;
            put '[]';
        run;
        %return;
    %end;

    /*--------------------------------------------------------------
      Build the OLE DB init string from the configuration block.
      Integrated Security (SSPI) uses the SAS server process account and
      carries no credentials; the SQL authentication branch supplies them.
    --------------------------------------------------------------*/
    %local _init;
    %if &DMW_TRUSTED = 1 %then %do;
        %let _init = Provider=&DMW_PROVIDER;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=&DMW_CATALOG;Data Source=&DMW_SERVER,&DMW_PORT;
    %end;
    %else %do;
        %let _init = Provider=&DMW_PROVIDER;User ID=&DMW_USER;Password=&DMW_PASS;Persist Security Info=True;Initial Catalog=&DMW_CATALOG;Data Source=&DMW_SERVER,&DMW_PORT;
    %end;

    %if &debug_mode=1 %then %put NOTE: [DEBUG] Target=&DMW_SERVER,&DMW_PORT catalog=&DMW_CATALOG trusted=&DMW_TRUSTED;

    /*--------------------------------------------------------------
      Execute SQL (explicit pass-through)
    --------------------------------------------------------------*/
    proc sql noprint;
        connect to oledb as myconn (init_string="&_init");

        create table work.dmw_result as
        select * from connection to myconn (
            &query
        );
        disconnect from myconn;
    quit;

    /*--------------------------------------------------------------
      If the pass-through failed, return a JSON error object rather than
      letting the browser receive a partial or empty body with no
      explanation. The page surfaces .error to the user.
    --------------------------------------------------------------*/
    %if &sqlrc > 4 or %sysfunc(exist(work.dmw_result)) = 0 %then %do;
        %local _dberr;
        %let _dberr = %superq(sysdbmsg);
        %put ERROR: [dmw_run] Pass-through failed. &_dberr;
        data _null_;
            length msg $2000;
            file _webout lrecl=32767;
            msg = symget('_dberr');
            msg = tranwrd(msg,'\','\\');
            msg = tranwrd(msg,'"','\"');
            msg = compbl(msg);
            put '{"error":"SQL execution failed against &DMW_SERVER/&DMW_CATALOG. ' msg +(-1) '"}';
        run;
        %return;
    %end;

    /*--------------------------------------------------------------
      Apply the row cap
    --------------------------------------------------------------*/
    %if &DMW_MAXROWS > 0 %then %do;
        data work.dmw_result;
            set work.dmw_result (obs=&DMW_MAXROWS);
        run;
    %end;

    /*--------------------------------------------------------------
      Count rows and elapsed time
    --------------------------------------------------------------*/
    %let _rowcount = 0;
    proc sql noprint;
        select count(*) into :_rowcount trimmed from work.dmw_result;
    quit;

    %let _end_time = %sysfunc(datetime());
    %let _elapsed = %sysevalf(&_end_time - &_start_time);

    /* Optional debug print to SAS log only */
    %if &debug_mode=1 %then %do;
        %put NOTE: [DEBUG] Rows returned=&_rowcount, Elapsed seconds=&_elapsed;
    %end;

    /*--------------------------------------------------------------
      Write execution log to the UNC folder (skipped when DMW_LOGFOLDER
      is empty, or silently skipped if the folder is not writable).
    --------------------------------------------------------------*/
    %if %length(&DMW_LOGFOLDER) > 0 %then %do;
        %let logfile = &DMW_LOGFOLDER\dmw_log_%sysfunc(today(),yymmddn8.).log;
        filename logf "&logfile";

        %if %sysfunc(fexist(logf)) or 1 %then %do;
            data _null_;
                file logf mod;
                ts = datetime();
                put "---- DMW Execution Log ----";
                put "Timestamp: " ts datetime20.;
                put "Server: &DMW_SERVER,&DMW_PORT";
                put "Catalog: &DMW_CATALOG";
                put "Rows returned: &_rowcount";
                put "Elapsed seconds: &_elapsed";
                put "Query: &query";
                put "----------------------------";
            run;
        %end;
        filename logf clear;
    %end;

    /*--------------------------------------------------------------
      Export dataset to JSON array only (no wrapper object)
    --------------------------------------------------------------*/
    filename _jsonout temp;

    proc json out=_jsonout nosastags pretty;
        export work.dmw_result / nosastags;
    run;

    /*--------------------------------------------------------------
      Stream JSON array to _webout
    --------------------------------------------------------------*/
    data _null_;
        infile _jsonout lrecl=32767 end=eof;
        file _webout lrecl=32767;
        do while (not eof);
            input;
            put _infile_;
        end;
    run;

    filename _jsonout clear;

    /*--------------------------------------------------------------
      Housekeeping
    --------------------------------------------------------------*/
    proc datasets library=work nolist nowarn;
        delete dmw_result;
    quit;

%mend dmw_run;

/*--------------------------------------------------------------
  Execute the macro
--------------------------------------------------------------*/
%dmw_run();
