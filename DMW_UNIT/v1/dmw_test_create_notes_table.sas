/*=============================================================================
  dmw_test_create_notes_table.sas                                    *** TEST ***
  -----------------------------------------------------------------------------
  DMW (Data Mart Workbench) -- IT Management Solution
  ONE-TIME SETUP: creates the DMW_JobNotes table used by the Job Notes feature.

  THIS IS A TEST / SETUP ARTEFACT. It is deliberately named with "test" so it is
  never mistaken for part of the running solution. Run it once, confirm the
  table exists, then UNREGISTER it from SAS Management Console. A stored process
  that can issue DDL should not stay reachable from a browser.

  -----------------------------------------------------------------------------
  READ THIS FIRST -- WILL THIS EVEN WORK?
  -----------------------------------------------------------------------------
  This program runs as the SAS Stored Process Server service account, which in
  UNIT is A70ADMED\SVC_CGSUSASSRV. Creating a table requires that account to
  hold CREATE TABLE permission in &DMW_CATALOG, plus ALTER on the target schema.

  The grant you asked the DBA for earlier was SELECT on one view. That is NOT
  enough to create a table. So one of three things is true:

      a) The service account already has db_ddladmin or db_owner -- this works.
      b) It does not -- STEP 1 below reports that, nothing is created, and you
         need the DBA after all. Send them dmw_test_notes_table.sql instead.
      c) The DBA will grant CREATE TABLE just long enough for this to run.

  STEP 1 runs a permission probe BEFORE attempting anything, so a run with
  insufficient rights is a clean report rather than a failed DDL statement.

  -----------------------------------------------------------------------------
  A NOTE ON WHERE THE TABLE GOES
  -----------------------------------------------------------------------------
  OHM_CGSScheduler is an operational database you do not own. Creating objects
  in it is a governance decision, not just a technical one. Preference order:

      1. A separate database owned by your team            (cleanest)
      2. A separate schema in OHM_CGSScheduler, e.g. [dmw] (good)
      3. dbo in OHM_CGSScheduler                           (default below)

  Set DMW_SCHEMA to change this. Whatever you choose, tell the owner of
  OHM_CGSScheduler that the object exists -- an undocumented table in someone
  else's production database is how outages get harder to diagnose.

  -----------------------------------------------------------------------------
  REGISTRATION IN SAS MANAGEMENT CONSOLE
  -----------------------------------------------------------------------------
    Name            : dmw_test_create_notes_table
    Result type     : Streaming      <-- REQUIRED
    Input parameters: none
    AFTER A SUCCESSFUL RUN: delete this stored process registration.

  -----------------------------------------------------------------------------
  OUTPUT
  -----------------------------------------------------------------------------
    Content-type: application/json

    {"step":"probe",  "can_create_table":0, "status":"BLOCKED", "message":"..."}
    {"step":"create", "table_exists":1,     "status":"OK",      "message":"..."}
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

/* Schema that will own the table. See "A NOTE ON WHERE THE TABLE GOES". */
%let DMW_SCHEMA   = dbo;

/* Table name. Keeping "TEST" out of it deliberately -- the TABLE is permanent
   even though this setup program is not. Change only if you must. */
%let DMW_NOTES_TABLE = DMW_JobNotes;

/*-----------------------------------------------------------------------------
  SAFETY INTERLOCK
  -----------------------------------------------------------------------------
  This program issues DDL. It does nothing at all until you deliberately set
  this to 1. Leave it at 0 between runs so an accidental invocation is inert.
-----------------------------------------------------------------------------*/
%let DMW_TEST_CONFIRM = 0;

/*---------------------------------------------------------------------------*/


%macro dmw_test_create_notes_table();

    /*--------------------------------------------------------------
      JSON response header
    --------------------------------------------------------------*/
    data _null_;
        rc = stpsrv_header('Content-type','application/json');
        rc = stpsrv_header('Cache-Control','no-store, no-cache, must-revalidate');
    run;

    /*--------------------------------------------------------------
      Safety interlock
    --------------------------------------------------------------*/
    %if &DMW_TEST_CONFIRM ne 1 %then %do;
        data _null_;
            file _webout lrecl=32767;
            put '{"step":"guard","status":"NOT_CONFIRMED","message":'
                '"Set DMW_TEST_CONFIRM = 1 at the top of dmw_test_create_notes_table.sas and run again. Nothing was changed."}';
        run;
        %return;
    %end;

    /*--------------------------------------------------------------
      Build the OLE DB init string.
      Built in a DATA step, NOT with %LET: an init string is
      semicolon-delimited and a semicolon ends a %LET statement.
    --------------------------------------------------------------*/
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

    /*==============================================================
      STEP 1 -- PERMISSION PROBE
      --------------------------------------------------------------
      Ask SQL Server what this connection may do before trying it.
      HAS_PERMS_BY_NAME(NULL, NULL, 'CREATE TABLE') returns 1 or 0.
      SUSER_SNAME() confirms which account actually connected, which
      is worth recording the first time you run this.
    ==============================================================*/
    %local _cancreate _whoami _exists;
    %let _cancreate = 0;
    %let _whoami    = unknown;
    %let _exists    = 0;

    proc sql noprint;
        connect to oledb as probe (init_string="&_init");

        create table work.dmw_probe as
        select * from connection to probe (
            SELECT CAST(HAS_PERMS_BY_NAME(NULL, NULL, 'CREATE TABLE') AS int) AS CanCreateTable,
                   CAST(SUSER_SNAME() AS varchar(128))                        AS ConnectedAs,
                   CAST(CASE WHEN OBJECT_ID('[&DMW_SCHEMA].[&DMW_NOTES_TABLE]','U') IS NULL
                             THEN 0 ELSE 1 END AS int)                        AS TableExists
        );

        disconnect from probe;
    quit;

    %if &sqlrc > 4 or %sysfunc(exist(work.dmw_probe)) = 0 %then %do;
        %local _dberr;
        %let _dberr = %superq(sysdbmsg);
        %put ERROR: [dmw_test_create_notes_table] Probe failed. &_dberr;
        data _null_;
            length msg $2000 json $4000;
            file _webout lrecl=32767;
            msg  = symget('_dberr');
            msg  = tranwrd(msg,'\','\\');
            msg  = tranwrd(msg,'"','\"');
            msg  = compbl(msg);
            json = '{"step":"probe","status":"CONNECT_FAILED","message":"Could not connect to '
                || "&DMW_SERVER/&DMW_CATALOG" || '. ' || strip(msg) || '"}';
            jlen = length(json);
            put json $varying4000. jlen;
        run;
        %return;
    %end;

    data _null_;
        set work.dmw_probe;
        call symputx('_cancreate', CanCreateTable);
        call symputx('_whoami',    ConnectedAs);
        call symputx('_exists',    TableExists);
    run;

    %put NOTE: [dmw_test] Connected as &_whoami / CanCreateTable=&_cancreate / TableExists=&_exists;

    /*--------------------------------------------------------------
      Already there? Report and stop. Creating it twice is not an
      error here, but silently continuing would hide a surprise.
    --------------------------------------------------------------*/
    %if &_exists = 1 %then %do;
        data _null_;
            file _webout lrecl=32767;
            put '{"step":"probe","status":"ALREADY_EXISTS","connected_as":"' "&_whoami" '",'
                '"table":"&DMW_SCHEMA.&DMW_NOTES_TABLE",'
                '"message":"The table already exists. Nothing was changed."}';
        run;
        %return;
    %end;

    /*--------------------------------------------------------------
      No rights? Say so plainly and name the remedy.
    --------------------------------------------------------------*/
    %if &_cancreate ne 1 %then %do;
        data _null_;
            file _webout lrecl=32767;
            put '{"step":"probe","status":"BLOCKED","can_create_table":0,'
                '"connected_as":"' "&_whoami" '",'
                '"message":"This account cannot create tables in &DMW_CATALOG. '
                'Send dmw_test_notes_table.sql to the DBA instead, or ask for CREATE TABLE to be granted temporarily."}';
        run;
        %return;
    %end;

    /*==============================================================
      STEP 2 -- CREATE THE TABLE
      --------------------------------------------------------------
      One statement per EXECUTE(). A semicolon inside EXECUTE() ends
      the SAS statement, so each T-SQL statement below is written
      WITHOUT a trailing semicolon and with no internal semicolons.
      IF NOT EXISTS makes each step idempotent.
    ==============================================================*/
    proc sql noprint;
        connect to oledb as ddl (init_string="&_init");

        execute (
            IF OBJECT_ID('[&DMW_SCHEMA].[&DMW_NOTES_TABLE]','U') IS NULL
            CREATE TABLE [&DMW_SCHEMA].[&DMW_NOTES_TABLE] (
                NoteID            int IDENTITY(1,1) NOT NULL,
                JobID             int               NOT NULL,
                RequestID         int                   NULL,
                NoteText          nvarchar(2000)    NOT NULL,
                CreatedBy         varchar(64)       NOT NULL,
                CreatedDatetime   datetime2(0)      NOT NULL
                                  CONSTRAINT DF_&DMW_NOTES_TABLE._Created DEFAULT (SYSDATETIME()),
                IsRetracted       bit               NOT NULL
                                  CONSTRAINT DF_&DMW_NOTES_TABLE._Retracted DEFAULT (0),
                RetractedBy       varchar(64)           NULL,
                RetractedDatetime datetime2(0)          NULL,
                CONSTRAINT PK_&DMW_NOTES_TABLE PRIMARY KEY CLUSTERED (NoteID)
            )
        ) by ddl;

        execute (
            IF NOT EXISTS (SELECT 1 FROM sys.indexes
                           WHERE name = 'IX_&DMW_NOTES_TABLE._JobID'
                             AND object_id = OBJECT_ID('[&DMW_SCHEMA].[&DMW_NOTES_TABLE]'))
            CREATE INDEX [IX_&DMW_NOTES_TABLE._JobID]
                ON [&DMW_SCHEMA].[&DMW_NOTES_TABLE] (JobID, CreatedDatetime DESC)
        ) by ddl;

        disconnect from ddl;
    quit;

    %local _ddlrc;
    %let _ddlrc = &sqlrc;

    /*==============================================================
      STEP 3 -- VERIFY, DO NOT ASSUME
    ==============================================================*/
    proc sql noprint;
        connect to oledb as vfy (init_string="&_init");
        create table work.dmw_verify as
        select * from connection to vfy (
            SELECT CAST(CASE WHEN OBJECT_ID('[&DMW_SCHEMA].[&DMW_NOTES_TABLE]','U') IS NULL
                             THEN 0 ELSE 1 END AS int) AS TableExists,
                   (SELECT COUNT(*) FROM sys.columns
                     WHERE object_id = OBJECT_ID('[&DMW_SCHEMA].[&DMW_NOTES_TABLE]')) AS ColumnCount
        );
        disconnect from vfy;
    quit;

    %local _nowexists _ncols;
    %let _nowexists = 0;
    %let _ncols     = 0;
    %if %sysfunc(exist(work.dmw_verify)) %then %do;
        data _null_;
            set work.dmw_verify;
            call symputx('_nowexists', TableExists);
            call symputx('_ncols',     ColumnCount);
        run;
    %end;

    data _null_;
        file _webout lrecl=32767;
        %if &_nowexists = 1 %then %do;
            put '{"step":"create","status":"OK","table_exists":1,'
                '"table":"&DMW_SCHEMA.&DMW_NOTES_TABLE",'
                '"column_count":' "&_ncols" ','
                '"connected_as":"' "&_whoami" '",'
                '"message":"Table created. Now grant SELECT, INSERT and UPDATE on it to the service account, then UNREGISTER this stored process."}';
        %end;
        %else %do;
            put '{"step":"create","status":"FAILED","table_exists":0,'
                '"sqlrc":' "&_ddlrc" ','
                '"message":"The DDL did not produce the table. Check the SAS log for the OLE DB error."}';
        %end;
    run;

    /*--------------------------------------------------------------
      Housekeeping
    --------------------------------------------------------------*/
    proc datasets library=work nolist nowarn;
        delete dmw_probe dmw_verify;
    quit;

%mend dmw_test_create_notes_table;

%dmw_test_create_notes_table();
