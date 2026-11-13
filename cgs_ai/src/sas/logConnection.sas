/*=====================================================================
  Program Name : logConnection.sas
  Author       : Manuel Figallo
  Purpose      : Log the network connections to the SAS Object Spawner
                 port (default 8591) from the SAS server's side, once per
                 interval, into a CSV time series. Used to find out WHEN
                 an Enterprise Guide connection disappears, so the drop
                 can be matched to an idle timeout (CS0478540), a network
                 event, or a spawner outage.
  Version      : 1.0beta
  Created      : 2026-11-05

  DESIGN
    Runs on the SAS server (compute tier), so it sees the server end of
    each connection and cannot see the client PC. Pair it with the
    PowerShell client logger to see both ends. It asks Windows for the
    connections with Get-NetTCPConnection, reads the answer through a
    PIPE, and appends one row per connection per interval. When nothing
    is connected a row with State=none is written, so a gap in the log
    always means the macro was not running.

    The macro keeps the SAS session busy for the whole run. Submit it
    from a SEPARATE Enterprise Guide session, or run it in batch.

    NOTE: never open a block comment inside this header. SAS comments do
    not nest, so the header would end at that inner close marker and
    everything below it would be parsed as live code.

  REQUIREMENT: the PIPE needs XCMD. Check with
      %put %sysfunc(getoption(xcmd));

  Output columns: Time, State, RemoteAddress, RemotePort
    Time is local server time, ISO 8601.

  Usage:
      %include "<path>\src\sas\logConnection.sas";
      %logConnection(
        port=8591, interval_seconds=60, duration_minutes=480,
        output_file_path=D:\Temp\sas_conn_log_before.csv
      );
=====================================================================*/


%macro logConnection(
    port             = 8591,       /* spawner port to watch               */
    interval_seconds = 60,         /* seconds between samples             */
    duration_minutes = 60,         /* total run time                      */
    output_file_path =,            /* REQUIRED; .csv, appended if exists  */
    client_ip        =,            /* optional; log only this client      */
    debug            = 0           /* 1 = print the command, do not run   */
);
/* Sample the connections to a TCP port on this server and append them to
   a CSV. Each sample records the connection state and the client address
   and port, or State=none when nothing is connected.
   Use in claims processing: when analysts report Enterprise Guide
   disconnects, run it before and after a KeepAlive change to show whether
   idle sessions now survive.                                              */

  %local _n _i _stamp _filter _rc;

  %put NOTE: cgs_ai SAS wrapper version 1.0beta (logConnection, port=&port);

  %if %length(%superq(output_file_path)) = 0 %then %do;
    %put ERROR: output_file_path is required.;
    %abort cancel 8;
  %end;

  %if %sysfunc(getoption(xcmd)) ne XCMD %then %do;
    %put ERROR: SAS is running with NOXCMD, so the PIPE cannot launch PowerShell.;
    %put ERROR- Ask the SAS administrators to run it, or use the PowerShell logger from the client PC.;
    %abort cancel 8;
  %end;

  %if %length(%superq(client_ip)) > 0 %then %let _filter = -RemoteAddress &client_ip;
  %else %let _filter = ;

  %let _n = %sysevalf(&duration_minutes * 60 / &interval_seconds, floor);

  /* PIPE command: header line is dropped; fields come back quoted */
  filename cgscmd pipe "powershell -NoProfile -NonInteractive -Command ""Get-NetTCPConnection -LocalPort &port &_filter -ErrorAction SilentlyContinue | Select-Object State,RemoteAddress,RemotePort | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1""";

  %if &debug = 1 %then %do;
    %put NOTE: debug=1 - the command was built but NOT executed.;
    %put NOTE- Samples that would run: &_n (every &interval_seconds seconds).;
    %put NOTE- Output file: &output_file_path;
    %put NOTE- Command: %sysfunc(pathname(cgscmd));
    filename cgscmd clear;
    %return;
  %end;

  /* write the header once */
  %if not %sysfunc(fileexist(&output_file_path)) %then %do;
    data _null_;
      file "&output_file_path";
      put "Time,State,RemoteAddress,RemotePort";
    run;
  %end;

  %do _i = 1 %to &_n;
    %let _stamp = %sysfunc(datetime(), e8601dt19.);

    data _null_;
      infile cgscmd truncover eof=done;
      file "&output_file_path" mod;
      length line $200;
      retain n 0;
      input line $200.;
      n + 1;
      line = compress(line, '"');
      put "&_stamp," line;
      return;
    done:
      if n = 0 then put "&_stamp,none,,";
      stop;
    run;

    %let _rc = %sysfunc(sleep(&interval_seconds, 1));
  %end;

  filename cgscmd clear;
  %put NOTE: logConnection finished. Samples written: &_n. Output: &output_file_path;
%mend logConnection;
