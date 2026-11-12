/*=====================================================================
  Program Name : logEGProcess.sas
  Author       : Manuel Figallo
  Purpose      : Write the SAS log of an Enterprise Guide process flow to
                 a file on the server. Put one call at the START of the
                 flow and one at the END. Everything the flow writes to
                 the SAS log between the two calls is captured in the
                 file, with a header (who, where, when) and a footer
                 (elapsed time, count of ERROR and WARNING lines).
                 Used to keep evidence of a long flow (for example the
                 runs behind the Enterprise Guide session drops, CS0478540)
                 after the Enterprise Guide log tabs are closed.
  Version      : 1.0beta
  Created      : 2026-11-12

  DESIGN
    The macro uses PROC PRINTTO to send the session log to a file. It
    runs in the same SAS session as the rest of the flow, so it needs no
    XCMD and no extra software. A flow runs in ONE workspace server
    session, so macro variables set by START are still there when STOP
    runs. The START call remembers the log options and the file; the STOP
    call puts the log back, adds the footer, and restores the options.

    WHAT YOU WILL SEE IN ENTERPRISE GUIDE
      While logging is on, the log goes to the file and NOT to the log
      tab. The log tabs of the nodes between START and STOP are empty.
      That is expected. Use replay=1 on the STOP call to also copy the
      captured log into the STOP node log tab.

    NOTE: never open a block comment inside this header. SAS comments do
    not nest, so the header would end at that inner close marker and
    everything below it would be parsed as live code.

  STEP-BY-STEP INSTRUCTIONS

    Before you start
      1. Pick a folder the SAS SERVER can write to. The path is read by
         the server, not by your PC, so use a UNC path or a server drive
         (for example \\server\share\logs). A path on your own C: drive
         will not work.
      2. Pick a file name that ends in .log, for example
         \\server\share\logs\eg_flow_2026-11-12.log
      3. Save this file where the server can read it, for example
         <path>\src\sas\logEGProcess.sas

    Step 1 - Add a START node at the top of the process flow
      a. Open your project and open the process flow.
      b. Add a new code node (File, New, Program) and name it
         00_START_log.
      c. Paste these lines in it:
             %include "<path>\src\sas\logEGProcess.sas";
             %logEGProcess(
               action=start,
               output_file_path=\\server\share\logs\eg_flow_2026-11-12.log
             );
      d. Run this one node once on its own. Its log tab shows a NOTE
         that logging has started. From now on the log goes to the file.

    Step 2 - Add a STOP node at the bottom of the process flow
      a. Add another code node and name it 99_STOP_log.
      b. Paste this line in it:
             %logEGProcess(action=stop);
         Optional: action=stop, replay=1 also copies the captured log
         into this node log tab so you can read it in Enterprise Guide.

    Step 3 - Put the nodes in the right order
      a. Link the START node to the first node(s) of your real work, and
         link the last node(s) of your real work to the STOP node.
         Enterprise Guide runs nodes in link order, so START runs first
         and STOP runs last. Use the link options on the node right-click
         menu to draw the links.
      b. Optional: add a marker between steps with a code node holding
             %logEGProcess(action=mark, label=Finished loading claims);
         A marker adds one time-stamped line to the file so you can find
         a step. If label is blank, the node label is used when available.

    Step 4 - Run the flow
      a. Run the whole process flow (right-click the flow background and
         choose Run Process Flow). Do not run other flows in the same
         Enterprise Guide session at the same time: the log of every
         submission in the session goes to the same file.
      b. When the STOP node finishes, its log tab shows the file name,
         the number of ERROR and WARNING lines, and the elapsed time.

    Step 5 - Read the file
      a. Open the .log file with Notepad or any text editor, from the
         server path you used.
      b. Each run starts with a line that reads
             >>> logEGProcess RUN <id> START
         and ends with
             >>> logEGProcess RUN <id> END
         Search for ERROR: or WARNING: to find problems.
      c. By default new runs are APPENDED to the same file, one after
         another. Use append=0 on START to start a fresh file.

    If the flow stops early (an error, or you press Stop)
      The STOP node may not run, and the log stays redirected: the log
      tabs stay empty. Fix it by running a code node that holds
             %logEGProcess(action=stop);
      It is safe to run STOP at any time. If logging is not on, it only
      prints a NOTE. Starting a new run while one is still on closes the
      old run first.

    If the SAS session itself ends before STOP runs (server restart,
    connection lost), the file simply ends with no END line. That is
    useful evidence on its own: the last time stamp is the last moment
    the session was alive.

  Parameters
      action           start | stop | mark         (default start)
      output_file_path REQUIRED for start; a .log file on the server
      append           1 = add to an existing file (default), 0 = replace
      verbose          1 = turn on source, source2, notes, stimer,
                       fullstimer and msglevel=i while logging (default);
                       the options are put back at STOP
      replay           1 = at STOP, also copy this run into the STOP node
                       log tab (default 0). If the captured log holds
                       ERROR lines, Enterprise Guide may then mark the STOP
                       node as failed, because it reads the log text.
      label            text for action=mark
      debug            1 = print the plan, change nothing

  Globals created: LOGEGP_ACTIVE, LOGEGP_FILE, LOGEGP_RUNID, LOGEGP_T0,
  LOGEGP_OPTS. They are only created once, so it is safe to %include this
  file in every node.

  Usage (also see Examples_cgs_ai.sas):
      %include "<path>\src\sas\logEGProcess.sas";
      %logEGProcess(action=start, output_file_path=D:\Temp\eg_flow.log);
      ... the rest of the process flow ...
      %logEGProcess(action=stop);
=====================================================================*/


%macro logEGProcess(
    action           = start,      /* start | stop | mark                 */
    output_file_path =,            /* REQUIRED for start; .log, on server */
    append           = 1,          /* 1 = append, 0 = replace the file    */
    verbose          = 1,          /* 1 = full detail log options         */
    replay           = 0,          /* 1 = copy the run into the STOP log  */
    label            =,            /* text for action=mark                */
    debug            = 0           /* 1 = print the plan, change nothing  */
);
/* Capture the SAS log of an Enterprise Guide process flow into a file.
   START redirects the session log to the file and writes a header.
   STOP restores the log, writes a footer with the ERROR and WARNING
   counts and the elapsed time, and puts the log options back.
   MARK writes one time-stamped line so a step is easy to find.
   Use in claims processing: when an analyst reports that a long flow was
   cut off, the file shows how far it got and when the last line was
   written.                                                                  */

  %local _act _stamp _new _direxists _k _v _file _runid _elapsed _elapsedtxt
         _nerr _nwarn _lbl;

  /* create the session globals once; re-including this file never resets them */
  %if not %symexist(LOGEGP_ACTIVE) %then %do;
    %global LOGEGP_ACTIVE LOGEGP_FILE LOGEGP_RUNID LOGEGP_T0 LOGEGP_OPTS;
    %let LOGEGP_ACTIVE = 0;
  %end;

  %let _act = %upcase(%superq(action));

  %put NOTE: cgs_ai SAS wrapper version 1.0beta (logEGProcess, action=&_act);

  /*------------------------------------------------------------------ START */
  %if &_act = START %then %do;

    %if %length(%superq(output_file_path)) = 0 %then %do;
      %put ERROR: output_file_path is required for action=start.;
      %abort cancel 8;
    %end;

    /* the folder must exist on the server (checked from the file name) */
    data _null_;
      length p d $1024;
      p = strip(symget('output_file_path'));
      k = findc(p, '\/', 'b');            /* last path separator */
      if k > 1 then d = substr(p, 1, k - 1);
      else d = ' ';
      if d = ' ' then ok = 1;             /* no folder part: nothing to check */
      else ok = (fileexist(strip(d)) > 0);
      call symputx('_direxists', ok, 'l');
    run;
    %if &_direxists = 0 %then %do;
      %put ERROR: the folder for output_file_path does not exist on the SAS server.;
      %put ERROR- Use a UNC path or a server drive, not a drive on your own PC.;
      %abort cancel 8;
    %end;

    %if &debug = 1 %then %do;
      %put NOTE: debug=1 - nothing was changed.;
      %put NOTE- The session log would be sent to: &output_file_path;
      %if &append = 1 %then %put NOTE- Mode: append. Verbose options: &verbose.;
      %else %put NOTE- Mode: replace. Verbose options: &verbose.;
      %return;
    %end;

    /* a run is already on: close it cleanly before starting the next */
    %if &LOGEGP_ACTIVE = 1 %then %do;
      %put WARNING: logEGProcess was already on. Closing the earlier run first.;
      %logEGProcess(action=stop);
    %end;

    %let _stamp = %sysfunc(datetime(), e8601dt19.);
    %let LOGEGP_RUNID = %sysfunc(compress(&_stamp, -:))_&sysjobid;
    %let LOGEGP_FILE  = &output_file_path;
    %let LOGEGP_T0    = %sysfunc(datetime());

    /* remember the current log options so STOP can put them back */
    %let LOGEGP_OPTS = %sysfunc(getoption(source, keyword))
                       %sysfunc(getoption(source2, keyword))
                       %sysfunc(getoption(notes, keyword))
                       %sysfunc(getoption(stimer, keyword))
                       %sysfunc(getoption(fullstimer, keyword))
                       %sysfunc(getoption(msglevel, keyword));

    %put NOTE: logEGProcess is now writing the SAS log to &output_file_path;
    %put NOTE- The log tabs of the nodes after this one stay empty until you run action=stop.;

    %let _new = ;
    %if &append ne 1 %then %let _new = new;

    /* ---- from here on, the log goes to the file ---- */
    proc printto log="&output_file_path" &_new;
    run;

    %if &verbose = 1 %then %do;
      options source source2 notes stimer fullstimer msglevel=i;
    %end;

    %put ====================================================================;
    %put >>> logEGProcess RUN &LOGEGP_RUNID START;
    %put Started : &_stamp (server local time);
    %put User    : &sysuserid   Host: &syshostname   Session job id: &sysjobid;
    %put SAS     : &sysvlong   OS: &sysscp;
    %do _k = 1 %to 6;
      %let _v = %scan(_CLIENTAPP _CLIENTAPPVERSION _CLIENTMACHINE _CLIENTUSERID _CLIENTPROJECTNAME _CLIENTPROJECTPATH, &_k, %str( ));
      %if %symexist(&_v) %then %put Client  : &_v = &&&_v;
    %end;
    %put Log file: &LOGEGP_FILE;
    %put ====================================================================;

    %let LOGEGP_ACTIVE = 1;
  %end;

  /*------------------------------------------------------------------- MARK */
  %else %if &_act = MARK %then %do;
    %if &LOGEGP_ACTIVE ne 1 %then %do;
      %put NOTE: logEGProcess is not on, so the marker was not written.;
      %return;
    %end;
    %let _lbl = %superq(label);
    %if %length(&_lbl) = 0 and %symexist(_CLIENTTASKLABEL) %then %let _lbl = &_CLIENTTASKLABEL;
    %let _stamp = %sysfunc(datetime(), e8601dt19.);
    %put ---- [&_stamp] MARK: &_lbl ----;
  %end;

  /*------------------------------------------------------------------- STOP */
  %else %if &_act = STOP %then %do;
    %if &LOGEGP_ACTIVE ne 1 %then %do;
      %put NOTE: logEGProcess is not on. Nothing to stop.;
      %return;
    %end;

    %let _file    = &LOGEGP_FILE;
    %let _runid   = &LOGEGP_RUNID;
    %let _stamp   = %sysfunc(datetime(), e8601dt19.);
    %let _elapsed = %sysevalf(%sysfunc(datetime()) - &LOGEGP_T0);
    %let _elapsedtxt = %sysfunc(putn(&_elapsed, time8.));
    %let _nerr  = 0;
    %let _nwarn = 0;

    /* ---- from here on, the log comes back to Enterprise Guide ---- */
    proc printto;
    run;

    %if %length(&LOGEGP_OPTS) > 0 %then %do;
      options &LOGEGP_OPTS;
    %end;

    /* count the ERROR and WARNING lines of THIS run only */
    data _null_;
      infile "&_file" lrecl=1024 truncover end=eof;
      input line $char1024.;
      retain inrun 0 nerr 0 nwarn 0;
      if index(line, ">>> logEGProcess RUN &_runid START") then inrun = 1;
      if inrun then do;
        if substr(line, 1, 5) = 'ERROR' then nerr + 1;
        else if substr(line, 1, 7) = 'WARNING' then nwarn + 1;
      end;
      if eof then do;
        call symputx('_nerr', nerr, 'l');
        call symputx('_nwarn', nwarn, 'l');
      end;
    run;

    /* footer, added to the end of the file */
    data _null_;
      file "&_file" mod;
      put "--------------------------------------------------------------------";
      put ">>> logEGProcess RUN &_runid END";
      put "Ended   : &_stamp (server local time)";
      put "Elapsed : &_elapsedtxt (hh:mm:ss)";
      put "Errors  : &_nerr    Warnings: &_nwarn";
      put "====================================================================";
    run;

    /* optional: copy this run into the STOP node log tab */
    %if &replay = 1 %then %do;
      data _null_;
        infile "&_file" lrecl=1024 truncover;
        input line $char1024.;
        retain inrun 0;
        if index(line, ">>> logEGProcess RUN &_runid START") then inrun = 1;
        if inrun then do;
          n = lengthn(line);
          if n > 0 then put line $varying1024. n;
          else put ' ';
        end;
      run;
    %end;

    %let LOGEGP_ACTIVE = 0;
    %put NOTE: logEGProcess finished. Log file: &_file;
    %put NOTE- Errors: &_nerr  Warnings: &_nwarn  Elapsed: &_elapsedtxt.;
  %end;

  %else %do;
    %put ERROR: action must be start, stop or mark. Received: &action;
    %abort cancel 8;
  %end;

%mend logEGProcess;
