/*=============================================================================
  dmw_test_note_list.sas                                             *** TEST ***
  -----------------------------------------------------------------------------
  DMW (Data Mart Workbench) -- IT Management Solution
  Job Notes feature: READ side. Returns notes as a JSON array.

  This is the read half of the pair. It takes NO SQL from the caller -- only a
  JobID and a row limit. That is the whole point: dmw_run accepts a statement
  and is therefore guarded by a SELECT-only check; this one accepts no
  statement at all, so there is nothing to guard.

  -----------------------------------------------------------------------------
  INPUT PARAMETERS
  -----------------------------------------------------------------------------
    job_id      Integer. Notes for this JobID. Omit or pass 0 for "all recent
                notes across every job", which is what the Notes tab uses.
    max_rows    Integer, optional. Defaults to DMW_NOTE_MAXROWS below.
    debug_mode  0 or 1. Writes diagnostics to the SAS log, never to the
                response.

  -----------------------------------------------------------------------------
  OUTPUT
  -----------------------------------------------------------------------------
    Content-type: application/json
    [ { "NoteID":1, "JobID":101, "RequestID":null, "NoteText":"...",
        "CreatedBy":"731o", "CreatedDatetime":"2026-11-10T08:15:00",
        "IsRetracted":0 } ]

    [] when the job has no notes.
    {"error":"..."} on failure.

  -----------------------------------------------------------------------------
  REGISTRATION IN SAS MANAGEMENT CONSOLE
  -----------------------------------------------------------------------------
    Name            : dmw_test_note_list
    Result type     : Streaming      <-- REQUIRED
    Input parameters: job_id (Text), max_rows (Text), debug_mode (Text)
                      -- none required at prompt
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

/* Default cap on returned notes. */
%let DMW_NOTE_MAXROWS = 500;

/* Include retracted notes in the response?
     0 = hide them entirely (default)
     1 = return them with IsRetracted = 1 so the page can grey them out */
%let DMW_SHOW_RETRACTED = 0;

/*---------------------------------------------------------------------------*/


%global job_id max_rows debug_mode;

%macro dmw_test_note_list();

    %if %symexist(debug_mode)=0 %then %let debug_mode=0;
    %if %superq(debug_mode)=     %then %let debug_mode=0;

    data _null_;
        rc = stpsrv_header('Content-type','application/json');
        rc = stpsrv_header('Pragma','no-cache');
        rc = stpsrv_header('Cache-Control','no-store, no-cache, must-revalidate');
    run;

    /*--------------------------------------------------------------
      Sanitise the inputs.
      Both are integers, so coerce them to numbers and throw away
      anything else. Nothing the caller sends reaches SQL as text --
      an input that is not a number becomes 0, not a fragment of a
      statement.
    --------------------------------------------------------------*/
    %local _jobid _maxrows;
    data _null_;
        j = input(symget('job_id'),   ?? best.);
        m = input(symget('max_rows'), ?? best.);
        if missing(j) or j < 0 then j = 0;
        if missing(m) or m <= 0 or m > &DMW_NOTE_MAXROWS then m = &DMW_NOTE_MAXROWS;
        call symputx('_jobid',   j);
        call symputx('_maxrows', m);
    run;

    %if &debug_mode=1 %then %put NOTE: [DEBUG] note_list job_id=&_jobid max_rows=&_maxrows;

    /*--------------------------------------------------------------
      Init string -- DATA step, not %LET (semicolons end a %LET).
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

    /*--------------------------------------------------------------
      Read. The statement is fixed; only the two integers vary.
    --------------------------------------------------------------*/
    proc sql noprint;
        connect to oledb as notesdb (init_string="&_init");

        create table work.dmw_notes as
        select * from connection to notesdb (
            SELECT TOP (&_maxrows)
                   NoteID,
                   JobID,
                   RequestID,
                   NoteText,
                   CreatedBy,
                   CONVERT(varchar(19), CreatedDatetime, 126) AS CreatedDatetime,
                   CAST(IsRetracted AS int)                   AS IsRetracted,
                   RetractedBy,
                   CONVERT(varchar(19), RetractedDatetime, 126) AS RetractedDatetime
              FROM [&DMW_SCHEMA].[&DMW_NOTES_TABLE]
             WHERE (&_jobid = 0 OR JobID = &_jobid)
               AND (&DMW_SHOW_RETRACTED = 1 OR IsRetracted = 0)
             ORDER BY CreatedDatetime DESC, NoteID DESC
        );

        disconnect from notesdb;
    quit;

    %if &sqlrc > 4 or %sysfunc(exist(work.dmw_notes)) = 0 %then %do;
        %local _dberr;
        %let _dberr = %superq(sysdbmsg);
        %put ERROR: [dmw_test_note_list] Read failed. &_dberr;
        data _null_;
            length msg $2000 json $4000;
            file _webout lrecl=32767;
            msg  = symget('_dberr');
            msg  = tranwrd(msg,'\','\\');
            msg  = tranwrd(msg,'"','\"');
            msg  = compbl(msg);
            json = '{"error":"Could not read notes from '
                || "&DMW_SCHEMA..&DMW_NOTES_TABLE" || '. ' || strip(msg)
                || ' If the table does not exist yet, run dmw_test_create_notes_table or have the DBA run dmw_test_notes_table.sql."}';
            jlen = length(json);
            put json $varying4000. jlen;
        run;
        %return;
    %end;

    %if &debug_mode=1 %then %do;
        %local _n;
        proc sql noprint; select count(*) into :_n trimmed from work.dmw_notes; quit;
        %put NOTE: [DEBUG] note_list returned &_n rows;
    %end;

    /*--------------------------------------------------------------
      Emit as a bare JSON array, same contract as dmw_run.
    --------------------------------------------------------------*/
    filename _jsonout temp;

    proc json out=_jsonout nosastags pretty;
        export work.dmw_notes / nosastags;
    run;

    data _null_;
        infile _jsonout lrecl=32767 end=eof;
        file _webout lrecl=32767;
        do while (not eof);
            input;
            put _infile_;
        end;
    run;

    filename _jsonout clear;

    proc datasets library=work nolist nowarn;
        delete dmw_notes;
    quit;

%mend dmw_test_note_list;

%dmw_test_note_list();
