/*=============================================================================
  sql_studio_sas_bridge.sas
  =============================================================================
  The server-side half of SQL Studio Browser Runner v5.

  WHY THIS EXISTS
  ---------------
  A web browser cannot connect to Microsoft SQL Server. SQL Server speaks the TDS
  protocol over a raw TCP socket on port 1433, and browser JavaScript has no raw
  TCP socket API -- only HTTP(S), WebSocket and WebRTC. There is no JavaScript
  library that changes this: the reference driver (tedious) requires Node's net,
  tls, dns and dgram modules and ships no browser build.

  A browser therefore needs something that speaks HTTP on the front and the
  database on the back. You already run exactly that: the SAS Stored Process Web
  Application. This program is that bridge. Registering it is a CONFIGURATION
  task in SAS Management Console -- no software is installed on any workstation.

      Browser (HTML page) --HTTP--> SAS Stored Process Web App
                                        |
                                        | SAS/ACCESS OLEDB
                                        v
                                   SQL Server (TDS/1433)

  REGISTRATION (SAS Management Console)
  -------------------------------------
    1. New Stored Process, name: sql_studio_sas_bridge
    2. Source code location: the folder holding this .sas file
    3. Output type:  Streaming
    4. Result type:  Streaming   (we emit raw JSON, not ODS HTML)
    5. Input parameter: dmq_query   (type Text, not required at prompt)
    6. Note the metadata path, e.g.
         /User Folders/731o/My Folder/sql_studio_sas_bridge
       and put it in APP_CONFIG.SAS_PROGRAM in the HTML file.

  CONTRACT
  --------
    REQUEST   GET  .../SASStoredProcess/do?_program=<path>&dmq_query=<URL-encoded SQL>
              (or POST with the same parameters as form fields)

    RESPONSE  Content-type: application/json
              Success: {"columns":[...],"rows":[...],"rowCount":n,"elapsed":"..."}
              Failure: {"error":"message"}

  SECURITY -- READ BEFORE DEPLOYING
  ---------------------------------
    * READONLY_MODE below rejects anything that is not a SELECT. Leave it at 1
      unless you have a specific reason and compensating controls. A stored
      process that executes arbitrary SQL is a serious exposure.
    * The stored process runs under its registered SAS identity. Grant that
      identity only the database permissions it actually needs.
    * Credentials belong in SAS metadata or an authdomain, never in this file.
    * MAXROWS caps the result so a careless SELECT cannot flood the browser.
=============================================================================*/


/*-----------------------------------------------------------------------------
  0. CONFIGURATION -- edit these to match your environment
-----------------------------------------------------------------------------*/

/* Target SQL Server. These must match APP_CONFIG in the HTML file. */
%let SQLSVR_HOST    = A70TUCGSDATA008.A70ADMED.COM;
%let SQLSVR_PORT    = 1433;
%let SQLSVR_CATALOG = DataMartKYA;

/* OLEDB provider. MSOLEDBSQL is the current Microsoft provider; SQLNCLI11 is
   the older Native Client. Use whichever is installed on the SAS server. */
%let SQLSVR_PROVIDER = MSOLEDBSQL;

/* Authentication.
     1 = Integrated Security (SSPI). The SAS server process account is used.
     0 = SQL Server authentication. Supply SQLSVR_USER / SQLSVR_PASS, ideally
         from SAS metadata or an authdomain rather than literals here. */
%let SQLSVR_TRUSTED = 1;
%let SQLSVR_USER    = ;
%let SQLSVR_PASS    = ;

/* 1 = accept SELECT statements only. Strongly recommended. */
%let READONLY_MODE = 1;

/* Maximum rows returned to the browser. */
%let MAXROWS = 5000;


/*-----------------------------------------------------------------------------
  1. Close ODS so nothing wraps our JSON in HTML
-----------------------------------------------------------------------------*/
ods _all_ close;
options nodate nonumber nocenter;


/*-----------------------------------------------------------------------------
  2. Helper: emit a JSON error document and stop
-----------------------------------------------------------------------------*/
%macro emit_error(msg);
  data _null_;
    file _webout;
    length cleaned $2000;
    /* Escape backslashes and double quotes so the JSON stays well formed. */
    cleaned = tranwrd("&msg", '\', '\\');
    cleaned = tranwrd(cleaned, '"', '\"');
    cleaned = compbl(cleaned);
    put 'Content-type: application/json; charset=utf-8';
    put 'Cache-Control: no-store, no-cache, must-revalidate';
    put 'Pragma: no-cache';
    put;
    put '{"error":"' cleaned +(-1) '"}';
  run;
  %put ERROR: [sql_studio_sas_bridge] &msg;
%mend emit_error;


/*-----------------------------------------------------------------------------
  3. Read and validate the inbound statement
-----------------------------------------------------------------------------*/
%global dmq_query;

%macro run_bridge;

  %local sqltext upper_sql started;
  %let started = %sysfunc(datetime());

  %if %superq(dmq_query) = %then %do;
    %emit_error(No SQL supplied. Expected a dmq_query parameter.);
    %return;
  %end;

  %let sqltext = %superq(dmq_query);

  /* Strip comments before inspecting the statement, so a comment cannot hide
     a forbidden keyword from the read-only guard. */
  data _null_;
    length s $32767;
    s = symget('sqltext');
    s = prxchange('s/\/\*[\s\S]*?\*\///', -1, s);   /* block comments */
    s = prxchange('s/--[^\n]*//', -1, s);           /* line comments  */
    s = strip(compbl(s));
    call symputx('sqlclean', s, 'L');
    call symputx('sqlupper', upcase(s), 'L');
  run;

  %if %length(&sqlclean) = 0 %then %do;
    %emit_error(The statement was empty after comments were removed.);
    %return;
  %end;

  %if &READONLY_MODE = 1 %then %do;
    /* Must begin with SELECT or WITH. */
    %if %qsubstr(&sqlupper, 1, 6) ne SELECT and %qsubstr(&sqlupper, 1, 4) ne WITH %then %do;
      %emit_error(READONLY_MODE is enabled: only SELECT or WITH statements are accepted.);
      %return;
    %end;

    /* Reject mutation and batch-separation attempts. */
    data _null_;
      length bad $40;
      s = symget('sqlupper');
      bad = '';
      if      prxmatch('/\bINSERT\b/',   s) then bad = 'INSERT';
      else if prxmatch('/\bUPDATE\b/',   s) then bad = 'UPDATE';
      else if prxmatch('/\bDELETE\b/',   s) then bad = 'DELETE';
      else if prxmatch('/\bDROP\b/',     s) then bad = 'DROP';
      else if prxmatch('/\bTRUNCATE\b/', s) then bad = 'TRUNCATE';
      else if prxmatch('/\bALTER\b/',    s) then bad = 'ALTER';
      else if prxmatch('/\bCREATE\b/',   s) then bad = 'CREATE';
      else if prxmatch('/\bGRANT\b/',    s) then bad = 'GRANT';
      else if prxmatch('/\bREVOKE\b/',   s) then bad = 'REVOKE';
      else if prxmatch('/\bMERGE\b/',    s) then bad = 'MERGE';
      else if prxmatch('/\bEXEC\b/',     s) then bad = 'EXEC';
      else if prxmatch('/\bEXECUTE\b/',  s) then bad = 'EXECUTE';
      else if prxmatch('/\bBACKUP\b/',   s) then bad = 'BACKUP';
      else if prxmatch('/\bRESTORE\b/',  s) then bad = 'RESTORE';
      else if prxmatch('/\bSHUTDOWN\b/', s) then bad = 'SHUTDOWN';
      else if prxmatch('/XP_CMDSHELL/',  s) then bad = 'XP_CMDSHELL';
      else if prxmatch('/SP_CONFIGURE/', s) then bad = 'SP_CONFIGURE';
      call symputx('badkw', bad, 'L');
    run;

    %if %length(&badkw) > 0 %then %do;
      %emit_error(READONLY_MODE is enabled: the statement contains &badkw which is not permitted.);
      %return;
    %end;
  %end;


/*-----------------------------------------------------------------------------
  4. Connect to SQL Server and execute
-----------------------------------------------------------------------------*/
  %local connprops rc;

  %if &SQLSVR_TRUSTED = 1 %then %do;
    %let connprops = %str(properties=("Data Source"="&SQLSVR_HOST,&SQLSVR_PORT"
                                      "Initial Catalog"="&SQLSVR_CATALOG"
                                      "Integrated Security"="SSPI"));
  %end;
  %else %do;
    %let connprops = %str(properties=("Data Source"="&SQLSVR_HOST,&SQLSVR_PORT"
                                      "Initial Catalog"="&SQLSVR_CATALOG"
                                      "User ID"="&SQLSVR_USER"
                                      "Password"="&SQLSVR_PASS"));
  %end;

  libname sqlmart oledb provider=&SQLSVR_PROVIDER &connprops schema=dbo readbuff=1000;

  %if &syslibrc ne 0 %then %do;
    %emit_error(Could not connect to &SQLSVR_HOST,&SQLSVR_PORT catalog &SQLSVR_CATALOG. Check the provider is installed and the SAS identity has access.);
    %return;
  %end;

  /* Execute by passing the statement through to the database. CONNECT USING
     reuses the libname's established connection rather than opening a second. */
  %let syscc = 0;
  proc sql noerrorstop;
    connect using sqlmart as tgt;
    create table work._bridge_out as
      select * from connection to tgt (
        &sqlclean
      );
    disconnect from tgt;
  quit;

  %if &sqlrc > 4 or &syscc > 4 %then %do;
    %local dberr;
    %let dberr = %superq(sysdbmsg);
    libname sqlmart clear;
    %emit_error(SQL execution failed. &dberr);
    %return;
  %end;

  %if %sysfunc(exist(work._bridge_out)) = 0 %then %do;
    libname sqlmart clear;
    %emit_error(The statement executed but produced no result set. Only statements that return rows are supported.);
    %return;
  %end;

  /* Apply the row cap. */
  %local nobs;
  data work._bridge_capped;
    set work._bridge_out (obs=&MAXROWS);
  run;

  proc sql noprint;
    select count(*) into :nobs trimmed from work._bridge_capped;
  quit;


/*-----------------------------------------------------------------------------
  5. Emit the result set as JSON
-----------------------------------------------------------------------------*/
  /* Column names, in positional order, for the browser's grid headers. */
  %local collist;
  proc sql noprint;
    select name into :collist separated by '","'
    from dictionary.columns
    where libname = 'WORK' and memname = '_BRIDGE_CAPPED'
    order by varnum;
  quit;

  %local elapsed;
  %let elapsed = %sysfunc(putn(%sysevalf(%sysfunc(datetime()) - &started), 8.3));

  /* Headers first. PROC JSON writes the body. */
  data _null_;
    file _webout;
    put 'Content-type: application/json; charset=utf-8';
    put 'Cache-Control: no-store, no-cache, must-revalidate';
    put 'Pragma: no-cache';
    put;
  run;

  proc json out=_webout nosastags pretty;
    write open object;
      write values "columns";
      write open array;
        %if %length(&collist) > 0 %then %do;
          %local i thiscol;
          %let i = 1;
          %do %while (%scan(%str(&collist), &i, %str(")) ne %str());
            %let thiscol = %scan(%str(&collist), &i, %str("));
            %if %length(%trim(&thiscol)) > 0 and %trim(&thiscol) ne , %then %do;
              write values "%trim(&thiscol)";
            %end;
            %let i = %eval(&i + 1);
          %end;
        %end;
      write close;
      write values "rowCount" &nobs;
      write values "elapsed" "&elapsed";
      write values "rows";
      export work._bridge_capped / nosastags;
    write close;
  run;

  libname sqlmart clear;

  %put NOTE: ============================================================;
  %put NOTE: [sql_studio_sas_bridge] rows returned : &nobs;
  %put NOTE: [sql_studio_sas_bridge] elapsed       : &elapsed s;
  %put NOTE: [sql_studio_sas_bridge] statement     : &sqlclean;
  %put NOTE: ============================================================;

%mend run_bridge;

%run_bridge;


/*-----------------------------------------------------------------------------
  6. Housekeeping
-----------------------------------------------------------------------------*/
proc datasets library=work nolist nowarn;
  delete _bridge_out _bridge_capped;
quit;
