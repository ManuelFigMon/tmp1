/*=============================================================================
  dmw_update_jobname.sas
  -----------------------------------------------------------------------------
  DMW (Data Mart Workbench) -- IT Management Solution
  Updates ONE field: the job name, from the Data Grid.

  This stored process can change exactly one column on one row. It accepts no
  SQL from the caller. dmw_run.sas keeps its SELECT-only guard and is not
  involved in writing.

  -----------------------------------------------------------------------------
  WHICH TABLE IS ACTUALLY UPDATED
  -----------------------------------------------------------------------------
  The grid shows View_Scheduler_PBI_Dashboard.JobName, but a view is usually
  not updatable, and even where SQL Server permits it the write must resolve to
  exactly one base table. The job name lives on [dbo].[Job], so that is what
  this program writes, keyed on JobID. The view then reflects the change on the
  next read.

  If your site has made the view updatable and you would rather write through
  it, set DMW_TARGET_TABLE to the view name. Confirm with the DBA first: an
  UPDATE through a multi-table view fails at run time rather than at
  registration, so the failure would surface to an end user.

  -----------------------------------------------------------------------------
  WHY THE NEW NAME IS NOT CONCATENATED INTO SQL
  -----------------------------------------------------------------------------
  A job name is free text typed by a user. Building an UPDATE statement around
  it has two failure modes in SAS, one of which is a security hole:

    1. SQL injection -- an apostrophe alone breaks the statement.
    2. Macro resolution -- text reaching SQL as &new_name is resolved by the
       macro processor first, so a name containing & or % is silently
       corrupted or raises "Apparent symbolic reference not resolved".

  Both disappear if the text never becomes SQL. This program writes the new
  value into a SAS dataset with SYMGET and applies it through an OLE DB
  libname with PROC SQL UPDATE ... WHERE, so the driver parameterises the
  value. There is no statement to inject into.

  -----------------------------------------------------------------------------
  INPUT PARAMETERS  (sent by POST)
  -----------------------------------------------------------------------------
    job_id      Integer. Required. The job to rename.
    new_name    The new job name. Required. Capped at DMW_NAME_MAXLEN.
    old_name    Optional. When supplied, the update only applies if the stored
                value still matches. This is an optimistic concurrency check:
                two people renaming the same job will not silently overwrite
                each other.
    debug_mode  0 or 1.

  USE POST, NOT GET. A name reaches the web server log verbatim in a query
  string, and a browser refresh can silently re-submit a GET.

  -----------------------------------------------------------------------------
  OUTPUT
  -----------------------------------------------------------------------------
    Content-type: application/json
    {"status":"OK","job_id":101,"new_name":"...","changed_by":"731o","rows":1}
    {"status":"REJECTED","message":"..."}    validation failed, nothing written
    {"status":"CONFLICT","message":"..."}    old_name did not match
    {"error":"..."}                          database failure

  -----------------------------------------------------------------------------
  REGISTRATION IN SAS MANAGEMENT CONSOLE
  -----------------------------------------------------------------------------
    Name            : dmw_update_jobname
    Result type     : Streaming      <-- REQUIRED
    Input parameters: job_id, new_name, old_name, debug_mode
                      (all Text, none required at prompt)

  -----------------------------------------------------------------------------
  PERMISSIONS
  -----------------------------------------------------------------------------
    GRANT UPDATE ON [dbo].[Job] ([JobName]) TO [A70ADMED\SVC_CGSUSASSRV];

    Column-level UPDATE is deliberate. It allows this one field to change and
    nothing else on the row, so a defect here cannot alter a schedule.
=============================================================================*/


/*-----------------------------------------------------------------------------
  CONFIGURATION -- must match dmw_run.sas
-----------------------------------------------------------------------------*/
%let DMW_SERVER   = A70TUCGSDATA012.A70ADMED.COM;
%let DMW_PORT     = 1433;
%let DMW_CATALOG  = OHM_CGSScheduler;
%let DMW_PROVIDER = MSOLEDBSQL19;
%let DMW_TRUSTED  = 1;
%let DMW_USER     = ;
%let DMW_PASS     = ;

%let DMW_SCHEMA       = dbo;
%let DMW_TARGET_TABLE = Job;          /* base table that owns JobName */
%let DMW_KEY_COLUMN   = JobID;
%let DMW_NAME_COLUMN  = JobName;

/* Maximum accepted length. Must not exceed the column width in SQL Server. */
%let DMW_NAME_MAXLEN  = 200;

/* Require the caller to send old_name and require it to match?
     1 = yes (recommended) -- blocks silent overwrites between two users
     0 = no  -- last write wins */
%let DMW_REQUIRE_MATCH = 1;

/* Audit trail. Every rename is appended here. Set empty to disable. */
%let DMW_LOGFOLDER = \\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\Web_Stored_Processes\Logs;

/*---------------------------------------------------------------------------*/


%global job_id new_name old_name debug_mode;

%macro dmw_update_jobname();

    %if %symexist(debug_mode)=0 %then %let debug_mode=0;
    %if %superq(debug_mode)=     %then %let debug_mode=0;

    data _null_;
        rc = stpsrv_header('Content-type','application/json');
        rc = stpsrv_header('Cache-Control','no-store, no-cache, must-revalidate');
    run;

    /*==============================================================
      1. IDENTITY -- server-side, never from the caller
      --------------------------------------------------------------
      Every write arrives as the shared service account, so the
      database cannot say who renamed a job. SAS can.
    ==============================================================*/
    %local _user;
    data _null_;
        length u $64;
        u = symget('_METAPERSON');
        if missing(u) then u = symget('_SECUREUSERNAME');
        if missing(u) then u = symget('_USERNAME');
        call symputx('_user', strip(u));
    run;

    %if %length(&_user) = 0 %then %do;
        data _null_;
            file _webout lrecl=32767;
            put '{"status":"REJECTED","message":"No SAS user identity on this request. '
                'Sign in to SAS and try again. A rename is never applied without an author."}';
        run;
        %return;
    %end;

    /*==============================================================
      2. VALIDATE
    ==============================================================*/
    %local _jobid _len _reject;
    %let _reject = ;

    data _null_;
        length n $&DMW_NAME_MAXLEN o $&DMW_NAME_MAXLEN reject $300;
        j = input(symget('job_id'), ?? best.);

        /* SYMGET, never &new_name -- see the header. */
        n = strip(symget('new_name'));
        o = strip(symget('old_name'));

        if missing(j) or j <= 0 then
            reject = 'A valid job_id is required.';
        else if n = '' then
            reject = 'The new job name is empty.';
        else if length(n) > &DMW_NAME_MAXLEN then
            reject = "The new job name is longer than &DMW_NAME_MAXLEN characters.";
        else if n = o then
            reject = 'The new name is identical to the current name. Nothing to do.';
        else if &DMW_REQUIRE_MATCH = 1 and o = '' then
            reject = 'The current name was not supplied, so the change cannot be checked for conflicts. Reload the grid and try again.';
        /* Control characters would corrupt the display and the audit log. */
        else if n ne compress(n, , 'c') then
            reject = 'The new job name contains control characters.';

        call symputx('_jobid',  ifn(missing(j), 0, j));
        call symputx('_len',    length(n));
        call symputx('_reject', reject);
    run;

    %if %length(%superq(_reject)) > 0 %then %do;
        data _null_;
            length json $1000;
            file _webout lrecl=32767;
            json = '{"status":"REJECTED","message":"' || strip(symget('_reject')) || '"}';
            jlen = length(json);
            put json $varying1000. jlen;
        run;
        %return;
    %end;

    %if &debug_mode=1 %then
        %put NOTE: [DEBUG] update_jobname job_id=&_jobid len=&_len user=&_user;

    /*==============================================================
      3. CONNECT
      --------------------------------------------------------------
      Init string built in a DATA step, not with %LET: an init string
      is semicolon-delimited and a semicolon ends a %LET statement.
    ==============================================================*/
    %local _init;
    data _null_;
        length init $1000;
    %if &DMW_TRUSTED = 1 %then %do;
        init = "Provider=&DMW_PROVIDER;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=&DMW_CATALOG;Data Source=&DMW_SERVER,&DMW_PORT";
    %end;
    %else %do;
        init = "Provider=&DMW_PROVIDER;User ID=&DMW_USER;Password=&DMW_PASS;Persist Security Info=True;Initial Catalog=&DMW_CATALOG;Data Source=&DMW_SERVER,&DMW_PORT";
    %end;
        call symputx('_init', init);
    run;

    libname dmwdb oledb init_string="&_init" schema=&DMW_SCHEMA;

    %if &syslibrc ne 0 %then %do;
        data _null_;
            file _webout lrecl=32767;
            put '{"error":"Could not connect to &DMW_SERVER/&DMW_CATALOG. '
                'Check that the service account has a SQL Server login and that the provider is installed."}';
        run;
        %return;
    %end;

    /*==============================================================
      4. OPTIMISTIC CONCURRENCY CHECK
      --------------------------------------------------------------
      Read the stored name and compare it to what the browser thought
      it was. If someone else renamed the job since the grid loaded,
      stop rather than overwrite their change.
    ==============================================================*/
    %local _current _match;
    %let _current = ;
    %let _match   = 1;

    %if &DMW_REQUIRE_MATCH = 1 %then %do;

        proc sql noprint;
            create table work.dmw_cur as
            select &DMW_NAME_COLUMN as CurrentName
              from dmwdb.&DMW_TARGET_TABLE
             where &DMW_KEY_COLUMN = &_jobid;
        quit;

        %if %sysfunc(exist(work.dmw_cur)) = 0 %then %do;
            libname dmwdb clear;
            data _null_;
                file _webout lrecl=32767;
                put '{"status":"REJECTED","message":"Could not read the current job name. The update was not applied."}';
            run;
            %return;
        %end;

        %local _nrows;
        proc sql noprint; select count(*) into :_nrows trimmed from work.dmw_cur; quit;

        %if &_nrows = 0 %then %do;
            libname dmwdb clear;
            data _null_;
                file _webout lrecl=32767;
                put '{"status":"REJECTED","message":"No job with that JobID exists."}';
            run;
            %return;
        %end;

        data _null_;
            set work.dmw_cur;
            length o $&DMW_NAME_MAXLEN;
            o = strip(symget('old_name'));
            call symputx('_match', ifn(strip(CurrentName) = o, 1, 0));
            call symputx('_current', strip(CurrentName));
        run;

        %if &_match = 0 %then %do;
            libname dmwdb clear;
            data _null_;
                length json $1200 cur $&DMW_NAME_MAXLEN;
                file _webout lrecl=32767;
                cur  = symget('_current');
                cur  = tranwrd(cur,'\','\\');
                cur  = tranwrd(cur,'"','\"');
                json = '{"status":"CONFLICT","current_name":"' || strip(cur) ||
                       '","message":"This job was renamed by someone else since the grid loaded. Reload and try again."}';
                jlen = length(json);
                put json $varying1200. jlen;
            run;
            %return;
        %end;

    %end;

    /*==============================================================
      5. APPLY
      --------------------------------------------------------------
      The new value is carried in a macro variable read with SYMGET
      inside the DATA step that builds the update source, so it never
      passes through the macro processor as text to be resolved.
    ==============================================================*/
    data work.dmw_newname;
        length NewName $&DMW_NAME_MAXLEN;
        NewName = symget('new_name');      /* SYMGET -- raw value */
        NewName = strip(NewName);
        call symputx('_applied', NewName);
        output;
    run;

    proc sql noprint;
        update dmwdb.&DMW_TARGET_TABLE
           set &DMW_NAME_COLUMN = (select NewName from work.dmw_newname)
         where &DMW_KEY_COLUMN = &_jobid;
    quit;

    %local _ok _rows;
    %let _ok   = 0;
    %let _rows = &sqlobs;
    %if &sqlrc <= 4 %then %let _ok = 1;

    libname dmwdb clear;

    /*==============================================================
      6. AUDIT
      --------------------------------------------------------------
      The row itself carries no history of who renamed it, so the
      before and after values are recorded here.
    ==============================================================*/
    %if %length(&DMW_LOGFOLDER) > 0 %then %do;
        %let logfile = &DMW_LOGFOLDER\dmw_jobname_%sysfunc(today(),yymmddn8.).log;
        filename jlog "&logfile";
        data _null_;
            length old $&DMW_NAME_MAXLEN new $&DMW_NAME_MAXLEN;
            file jlog mod;
            ts  = datetime();
            old = symget('_current');
            new = symget('_applied');
            put "---- DMW JobName Change ----";
            put "Timestamp: " ts datetime20.;
            put "User: &_user";
            put "JobID: &_jobid";
            put "From: " old;
            put "To:   " new;
            put "Result: " "%sysfunc(ifc(&_ok=1,OK,FAILED))" " rows=&_rows";
            put "----------------------------";
        run;
        filename jlog clear;
    %end;

    /*==============================================================
      7. RESPOND
    ==============================================================*/
    data _null_;
        length nm $&DMW_NAME_MAXLEN json $2000;
        file _webout lrecl=32767;
        nm = symget('_applied');
        nm = tranwrd(nm,'\','\\');
        nm = tranwrd(nm,'"','\"');
        %if &_ok = 1 %then %do;
            json = '{"status":"OK","job_id":' || "&_jobid" ||
                   ',"new_name":"' || strip(nm) ||
                   '","changed_by":"' || "&_user" ||
                   '","rows":' || "&_rows" || '}';
        %end;
        %else %do;
            json = '{"status":"FAILED","message":"The update did not complete. '
                || 'The most likely cause is that the service account lacks UPDATE on '
                || "&DMW_SCHEMA..&DMW_TARGET_TABLE" || '. Check the SAS log."}';
        %end;
        jlen = length(json);
        put json $varying2000. jlen;
    run;

    proc datasets library=work nolist nowarn;
        delete dmw_newname dmw_cur;
    quit;

%mend dmw_update_jobname;

%dmw_update_jobname();
