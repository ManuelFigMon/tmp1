/*=============================================================================
  dmw_test_note_save.sas                                             *** TEST ***
  -----------------------------------------------------------------------------
  DMW (Data Mart Workbench) -- IT Management Solution
  Job Notes feature: WRITE side. Appends one note, or retracts one.

  This is the ONLY stored process in DMW that writes to the database, and it
  can do exactly two things: insert a note row, or set IsRetracted on one.
  It accepts no SQL from the caller.

  -----------------------------------------------------------------------------
  WHY THE WRITE IS NOT BUILT AS A SQL STRING
  -----------------------------------------------------------------------------
  The obvious implementation -- assemble an INSERT statement with the note text
  in it and send it through pass-through -- has two failure modes in SAS, one
  of which is a security hole:

    1. SQL injection. Note text is free text. An apostrophe alone breaks the
       statement; a crafted note does worse.

    2. Macro resolution. If the text reaches SQL via a macro reference such as
       &note_text, the macro processor resolves it FIRST. A note containing
       "&PROD" or "50% complete" is silently corrupted, or raises
       "WARNING: Apparent symbolic reference not resolved".

  Both disappear if the text never becomes SQL and never becomes a macro
  reference. So this program writes the note into a SAS dataset with SYMGET
  (which returns the raw value without resolving it) and lets PROC APPEND move
  it across an OLE DB libname. The driver parameterises the insert. There is no
  statement to inject into.

  -----------------------------------------------------------------------------
  WHY THE USERNAME IS NOT A PARAMETER
  -----------------------------------------------------------------------------
  Every database write arrives as the shared service account, so the database
  cannot tell who wrote a note. SAS can: _METAPERSON holds the identity of the
  person signed in to the web application.

  CreatedBy is taken from _METAPERSON HERE, inside the stored process. It is
  NEVER accepted from the browser. If the page supplied it, anyone could post a
  note under a colleague's name, and the audit trail would be worthless.

  -----------------------------------------------------------------------------
  INPUT PARAMETERS  (sent by POST -- see note below)
  -----------------------------------------------------------------------------
    action      "add" (default) or "retract"
    job_id      Integer. Required for "add".
    request_id  Integer, optional. Pins the note to one execution.
    note_text   The note. Required for "add". Capped at DMW_NOTE_MAXLEN.
    note_id     Integer. Required for "retract".
    debug_mode  0 or 1.

  USE POST, NOT GET. Three reasons: a 2,000-character note may not survive a
  query string; query strings are written verbatim into web server logs, so
  note text would leak there; and a browser refresh can silently re-submit a
  GET, duplicating the note. SAS stored processes accept form POSTs and expose
  the fields as macro variables in exactly the same way.

  -----------------------------------------------------------------------------
  OUTPUT
  -----------------------------------------------------------------------------
    Content-type: application/json
    {"status":"OK","action":"add","job_id":101,"created_by":"731o"}
    {"status":"REJECTED","message":"..."}   -- validation failed, nothing written
    {"error":"..."}                         -- database failure

  -----------------------------------------------------------------------------
  REGISTRATION IN SAS MANAGEMENT CONSOLE
  -----------------------------------------------------------------------------
    Name            : dmw_test_note_save
    Result type     : Streaming      <-- REQUIRED
    Input parameters: action, job_id, request_id, note_text, note_id,
                      debug_mode   (all Text, none required at prompt)
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

%let DMW_SCHEMA      = dbo;
%let DMW_NOTES_TABLE = DMW_JobNotes;

/* Maximum note length. Must not exceed the nvarchar(2000) column width. */
%let DMW_NOTE_MAXLEN = 2000;

/* May a user retract a note written by someone else?
     0 = only your own notes (default)
     1 = anyone may retract any note
   Retraction is a soft delete -- the original text is always preserved. */
%let DMW_RETRACT_ANY = 0;

/* Server-side log of every write, in addition to the row itself.
   Set to an empty value to disable. */
%let DMW_LOGFOLDER = \\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\Web_Stored_Processes\Logs;

/*---------------------------------------------------------------------------*/


%global action job_id request_id note_text note_id debug_mode;

%macro dmw_test_note_save();

    %if %symexist(debug_mode)=0 %then %let debug_mode=0;
    %if %superq(debug_mode)=     %then %let debug_mode=0;

    data _null_;
        rc = stpsrv_header('Content-type','application/json');
        rc = stpsrv_header('Cache-Control','no-store, no-cache, must-revalidate');
    run;

    /*==============================================================
      1. IDENTITY -- server-side, never from the caller
    ==============================================================*/
    %local _user;
    data _null_;
        length u $64;
        u = symget('_METAPERSON');
        if missing(u) then u = symget('_SECUREUSERNAME');
        if missing(u) then u = symget('_USERNAME');
        u = strip(u);
        call symputx('_user', u);
    run;

    %if %length(&_user) = 0 %then %do;
        data _null_;
            file _webout lrecl=32767;
            put '{"status":"REJECTED","message":"No SAS user identity on this request. '
                'Sign in to SAS and try again. A note is never written without an author."}';
        run;
        %return;
    %end;

    /*==============================================================
      2. VALIDATE
      --------------------------------------------------------------
      Integers are coerced with INPUT(... ?? best.), so anything that
      is not a number becomes missing and is rejected here rather
      than travelling any further.
    ==============================================================*/
    %local _action _jobid _reqid _noteid _len _reject;
    %let _reject = ;

    data _null_;
        length a $16 n $&DMW_NOTE_MAXLEN reject $300;
        a = lowcase(strip(symget('action')));
        if a = '' then a = 'add';
        if a not in ('add','retract') then reject = 'Unknown action. Expected add or retract.';

        j  = input(symget('job_id'),     ?? best.);
        r  = input(symget('request_id'), ?? best.);
        id = input(symget('note_id'),    ?? best.);

        /* SYMGET, never &note_text -- see the header. */
        n = symget('note_text');
        n = strip(n);

        if reject = '' and a = 'add' then do;
            if missing(j) or j <= 0 then
                reject = 'A valid job_id is required.';
            else if n = '' then
                reject = 'The note is empty.';
            else if length(n) > &DMW_NOTE_MAXLEN then
                reject = "The note is longer than &DMW_NOTE_MAXLEN characters.";
        end;

        if reject = '' and a = 'retract' and (missing(id) or id <= 0) then
            reject = 'A valid note_id is required to retract a note.';

        call symputx('_action', a);
        call symputx('_jobid',  ifn(missing(j),  0, j));
        call symputx('_reqid',  ifn(missing(r), ., r));
        call symputx('_noteid', ifn(missing(id), 0, id));
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
        %put NOTE: [DEBUG] note_save action=&_action job_id=&_jobid note_id=&_noteid len=&_len user=&_user;

    /*==============================================================
      3. CONNECT
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
      4a. ADD -- the note text never becomes SQL
      --------------------------------------------------------------
      PROC APPEND hands the row to the OLE DB driver, which
      parameterises the insert. There is no statement to inject into
      and no macro reference to resolve.
    ==============================================================*/
    %local _ok;
    %let _ok = 0;

    %if &_action = add %then %do;

        data work.dmw_newnote;
            length JobID 8 RequestID 8 NoteText $&DMW_NOTE_MAXLEN
                   CreatedBy $64 IsRetracted 8;
            JobID       = &_jobid;
            %if &_reqid ne . %then %do; RequestID = &_reqid; %end;
                                %else %do; RequestID = .;    %end;
            NoteText    = symget('note_text');      /* SYMGET -- raw value */
            NoteText    = strip(NoteText);
            CreatedBy   = "&_user";                  /* server-side identity */
            CreatedDatetime = datetime();
            format CreatedDatetime datetime20.;
            IsRetracted = 0;
            output;
        run;

        proc append base=dmwdb.&DMW_NOTES_TABLE data=work.dmw_newnote force;
        run;

        %if &syserr <= 4 %then %let _ok = 1;

    %end;

    /*==============================================================
      4b. RETRACT -- soft delete, original text preserved
      --------------------------------------------------------------
      A fixed UPDATE with two integers and one quoted identity value.
      The only free text involved is the username, which came from
      SAS and not from the caller.
    ==============================================================*/
    %else %do;

        proc sql noprint;
            connect to oledb as notesdb (init_string="&_init");
            execute (
                UPDATE [&DMW_SCHEMA].[&DMW_NOTES_TABLE]
                   SET IsRetracted       = 1,
                       RetractedBy       = '&_user',
                       RetractedDatetime = SYSDATETIME()
                 WHERE NoteID      = &_noteid
                   AND IsRetracted = 0
                   AND (&DMW_RETRACT_ANY = 1 OR CreatedBy = '&_user')
            ) by notesdb;
            disconnect from notesdb;
        quit;

        %if &sqlrc <= 4 %then %let _ok = 1;

    %end;

    libname dmwdb clear;

    /*==============================================================
      5. SERVER-SIDE LOG
      --------------------------------------------------------------
      The row records who wrote the note. This records that a write
      was attempted at all, including failures and rejections, which
      the table cannot show you.
    ==============================================================*/
    %if %length(&DMW_LOGFOLDER) > 0 %then %do;
        %let logfile = &DMW_LOGFOLDER\dmw_notes_%sysfunc(today(),yymmddn8.).log;
        filename nlog "&logfile";
        data _null_;
            file nlog mod;
            ts = datetime();
            put "---- DMW Note Write ----";
            put "Timestamp: " ts datetime20.;
            put "User: &_user";
            put "Action: &_action";
            put "JobID: &_jobid  NoteID: &_noteid  Length: &_len";
            put "Result: " "%sysfunc(ifc(&_ok=1,OK,FAILED))";
            put "------------------------";
        run;
        filename nlog clear;
    %end;

    /*==============================================================
      6. RESPOND
    ==============================================================*/
    data _null_;
        file _webout lrecl=32767;
        %if &_ok = 1 %then %do;
            put '{"status":"OK","action":"&_action","job_id":' "&_jobid" ','
                '"note_id":' "&_noteid" ',"created_by":"' "&_user" '",'
                '"message":"Saved."}';
        %end;
        %else %do;
            put '{"status":"FAILED","action":"&_action",'
                '"message":"The write did not complete. Check the SAS log. '
                'The most likely cause is that the service account lacks INSERT or UPDATE on &DMW_SCHEMA..&DMW_NOTES_TABLE."}';
        %end;
    run;

    proc datasets library=work nolist nowarn;
        delete dmw_newnote;
    quit;

%mend dmw_test_note_save;

%dmw_test_note_save();
