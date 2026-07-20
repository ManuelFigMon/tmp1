/*==============================================================================
  Program : run_sas_jobs.sas
  Purpose : Launch one or more Windows SAS batch jobs from a control (driver)
            SAS dataset.

            By DEFAULT the macro submits each job with the SYSTASK COMMAND
            statement (asynchronous, trackable via a task/status name).  A
            Boolean flag lets the caller switch to the classic X command
            instead.

  Key features
  ------------
    * SYSTASK COMMAND is the default launch mechanism.
    * USE_X= (Boolean 0/1) optionally switches to the X command.
    * EXEPATH= parameter defines the SAS executable
      (default: D:\SASHome\SASFoundation\9.4\sas.exe).
    * Each job is described by a row in a driver dataset -> "dataset driven".
    * The SASB switches built for every job:
          -SYSIN    "....sas"      (the program to run)
          -LOG      "....log"      (the log, with timestamp - see below)
          -NOSPLASH -NOLOGO -ICON  (optional, per-job flags)
    * The log file name carries a timestamp suffix, e.g. _20260719_023040.
      In the driver dataset that suffix is written as the token {DTN}; the
      macro replaces {DTN} with the current YYYYMMDD_HHMMSS stamp at run time.

  Author  : CGS SAS Support
==============================================================================*/


/*------------------------------------------------------------------------------
  MACRO : run_sas_jobs
------------------------------------------------------------------------------*/
%macro run_sas_jobs(
    data=                                              /* driver dataset (required)                 */,
    use_x=0                                            /* 0 = SYSTASK (default), 1 = X command      */,
    exepath=D:\SASHome\SASFoundation\9.4\sas.exe       /* full path to sas.exe                       */,
    wait=0                                             /* 0 = launch async, 1 = wait for each job    */,
    dryrun=0                                           /* 1 = only print the command, do NOT launch  */,
    /* -- names of the columns in the driver dataset -- */
    name_var=jobname                                   /* char : friendly job name                   */,
    sysin_var=sysin                                    /* char : path to the .sas program (-SYSIN)   */,
    log_var=log                                        /* char : path to the .log file   (-LOG)      */,
    nosplash_var=nosplash                              /* num  : 1 -> add -NOSPLASH                   */,
    nologo_var=nologo                                  /* num  : 1 -> add -NOLOGO                     */,
    icon_var=icon                                      /* num  : 1 -> add -ICON                       */,
    dtntoken=%str({DTN})                               /* token in the log path replaced by the stamp*/
);

  %local dtn nobs i q1 q2 sw cmd cmddisp tname sname waitopt;

  /*--- validate --------------------------------------------------------------*/
  %if %superq(data)= %then %do;
      %put ERROR: [run_sas_jobs] DATA= (the driver dataset) is required.;
      %return;
  %end;
  %if not %sysfunc(exist(&data)) %then %do;
      %put ERROR: [run_sas_jobs] Driver dataset &data does not exist.;
      %return;
  %end;

  /*--- 1. Build the {DTN} timestamp : YYYYMMDD_HHMMSS ------------------------*/
  /*     b8601dt15. -> 20260719T023040 ; translate the "T" separator to "_"    */
  %let dtn = %sysfunc(translate(
                 %sysfunc(putn(%sysfunc(datetime()),b8601dt15.)),
                 _,T));

  /*--- 2. Read the driver dataset into indexed macro variables ---------------*/
  data _null_;
      set &data end=__eof;
      call symputx(cats('m_name',    _n_), &name_var,                        'L');
      call symputx(cats('m_sysin',   _n_), &sysin_var,                       'L');
      /* substitute the {DTN} token in the log path with the timestamp        */
      call symputx(cats('m_log',     _n_), tranwrd(&log_var,"&dtntoken","&dtn"), 'L');
      call symputx(cats('m_nosplash',_n_), coalesce(&nosplash_var,0),        'L');
      call symputx(cats('m_nologo',  _n_), coalesce(&nologo_var,0),          'L');
      call symputx(cats('m_icon',    _n_), coalesce(&icon_var,0),            'L');
      if __eof then call symputx('nobs', _n_, 'L');
  run;

  %if &nobs=0 or &nobs= %then %do;
      %put WARNING: [run_sas_jobs] Driver dataset &data is empty - nothing to do.;
      %return;
  %end;

  /*  Two literal double-quote characters ("").  Inside a SAS double-quoted    */
  /*  string every "" collapses to a single ", so the OS ultimately receives   */
  /*  correctly quoted paths even though they contain spaces.                  */
  %let q2 = %str(%"%");
  /*  A single double-quote (") - used only to build a readable preview line.  */
  %let q1 = %str(%");

  /*--- 3. Loop over the jobs and launch each one -----------------------------*/
  %do i=1 %to &nobs;

      /* optional switches, added only when the flag column = 1 */
      %let sw=;
      %if &&m_nosplash&i = 1 %then %let sw=&sw -NOSPLASH;
      %if &&m_nologo&i   = 1 %then %let sw=&sw -NOLOGO;
      %if &&m_icon&i     = 1 %then %let sw=&sw -ICON;

      /* full command line (inner quotes doubled - see note above) */
      %let cmd=&q2.&exepath.&q2. -SYSIN &q2.&&m_sysin&i.&q2. -LOG &q2.&&m_log&i.&q2.&sw;
      /* readable copy (single quotes) for the log / preview only */
      %let cmddisp=&q1.&exepath.&q1. -SYSIN &q1.&&m_sysin&i.&q1. -LOG &q1.&&m_log&i.&q1.&sw;

      %put NOTE: [run_sas_jobs] ---------------------------------------------------;
      %put NOTE: [run_sas_jobs] Job &i of &nobs : &&m_name&i;
      %put NOTE: [run_sas_jobs] Log            : &&m_log&i;
      %put NOTE: [run_sas_jobs] Command        : &cmddisp;

      %if &dryrun=1 %then %do;
          %put NOTE: [run_sas_jobs] DRYRUN=1 - command shown only, not executed.;
      %end;
      %else %if &use_x=1 %then %do;
          /*---- X command path -------------------------------------------------*/
          /*  NOXSYNC/NOXWAIT -> fire-and-forget ; XSYNC/XWAIT -> wait for it.   */
          %if &wait=1 %then %do; options xsync   xwait;   %end;
          %else            %do; options noxsync  noxwait; %end;
          x "&cmd";
      %end;
      %else %do;
          /*---- SYSTASK COMMAND path (DEFAULT) --------------------------------*/
          %let tname=task_&i;
          %let sname=stat_&i;
          %if &wait=1 %then %let waitopt=wait;
          %else            %let waitopt=nowait;
          systask command "&cmd" taskname=&tname status=&sname &waitopt;
      %end;

  %end;

  %local how;
  %if &use_x=1 %then %let how=the X command;
  %else               %let how=SYSTASK COMMAND;
  %put NOTE: [run_sas_jobs] Submitted &nobs job(s) using &how..;

%mend run_sas_jobs;


/*------------------------------------------------------------------------------
  MACRO : run_sas_job   (thin convenience wrapper for a SINGLE job)
          Builds a one-row driver dataset and calls run_sas_jobs.
------------------------------------------------------------------------------*/
%macro run_sas_job(
    sysin=                                             /* path to the .sas program                   */,
    log=                                               /* path to the .log (may contain {DTN})        */,
    name=adhoc_job,
    use_x=0,
    exepath=D:\SASHome\SASFoundation\9.4\sas.exe,
    wait=0,
    dryrun=0,
    nosplash=1,
    nologo=1,
    icon=0
);
    data __one_job;
        length jobname $60 sysin log $400;
        jobname  = "&name";
        sysin    = "&sysin";
        log      = "&log";
        nosplash = &nosplash;
        nologo   = &nologo;
        icon     = &icon;
    run;

    %run_sas_jobs(data=__one_job, use_x=&use_x, exepath=&exepath,
                  wait=&wait, dryrun=&dryrun);

    proc datasets lib=work nolist; delete __one_job; quit;
%mend run_sas_job;
