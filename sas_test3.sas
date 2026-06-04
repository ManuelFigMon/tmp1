/***********************************************************************
  Program  : sas_test3.sas
  Purpose  : Call the proc_sql_guest stored process via PROC HTTP
             and save the HTML response locally. Timed end-to-end.
***********************************************************************/

/* ── 1. Start timer ── */
%let _t3_start = %sysfunc(datetime());
%put NOTE: [sas_test3] START — %sysfunc(datetime(), datetime20.);

/* ── 2. Configuration ── */
%let sp_server   = cgssaswebu.a70admed.com;
%let sp_port     = 8080;
%let sp_program  = /Shared Data/Public/proc_sql_guest;
%let out_file    = %sysfunc(getoption(work))/proc_sql_guest_output.html;

/* ── 3. Stored Process parameters ── */
%let p_nrows   = 1000;
%let p_ncols   = 30;
%let p_seed    = 12345;
%let p_out_fmt = TABLE;

/* ── 4. Call the stored process via PROC HTTP ── */
filename resp "&out_file";

proc http
    url    = "https://&sp_server./SASStoredProcess/guest"
    method = "GET"
    out    = resp;

    query
        "_program" = "&sp_program"
        "_action"  = "execute"
        "nrows"    = "&p_nrows"
        "ncols"    = "&p_ncols"
        "seed"     = "&p_seed"
        "out_fmt"  = "&p_out_fmt";
run;

/* ── 5. Check HTTP return code ── */
%put NOTE: [sas_test3] HTTP Status Code   = &SYS_PROCHTTP_STATUS_CODE.;
%put NOTE: [sas_test3] HTTP Status Phrase = &SYS_PROCHTTP_STATUS_PHRASE.;

%if &SYS_PROCHTTP_STATUS_CODE. = 200 %then %do;
    %put NOTE: [sas_test3] SUCCESS — response written to &out_file.;
%end;
%else %do;
    %put ERROR: [sas_test3] HTTP call failed. Status: &SYS_PROCHTTP_STATUS_CODE. &SYS_PROCHTTP_STATUS_PHRASE.;
    %let _t3_end     = %sysfunc(datetime());
    %let _t3_elapsed = %sysevalf(&_t3_end - &_t3_start);
    %put NOTE: [sas_test3] Elapsed: &_t3_elapsed seconds (failed);
    filename resp clear;
    %return;
%end;

/* ── 6. Read first 50 lines of response to SAS log ── */
data _null_;
    infile "&out_file" truncover lrecl=32767;
    input line $char32767.;
    if _n_ <= 50 then
        put "RESP[" _n_ "] " line;
run;

/* ── 7. Stop timer and report ── */
%let _t3_end     = %sysfunc(datetime());
%let _t3_elapsed = %sysevalf(&_t3_end - &_t3_start);

%put NOTE: [sas_test3] END     — %sysfunc(datetime(), datetime20.);
%put NOTE: [sas_test3] Elapsed: &_t3_elapsed seconds;

filename resp clear;
