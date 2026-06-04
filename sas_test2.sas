/***********************************************************************
  Program  : sas_test2.sas
  Purpose  : Implicit PROC SQL — generates &nrows x &ncols random
             numeric data inside SAS using ranuni(). No remote DBMS.
             Timed end-to-end for comparison with sas_test1.sas.
  Method   : Implicit (SAS DATA step + PROC SQL, all in-process)
***********************************************************************/

/* ── 1. Configurable parameters ── */
%let nrows = 1000;    /* number of rows to generate       */
%let ncols = 30;      /* number of random numeric columns  */
%let seed  = 12345;   /* random seed for reproducibility   */

/* ── 2. Start timer ── */
%let _t2_start = %sysfunc(datetime());
%put NOTE: [sas_test2] START — %sysfunc(datetime(), datetime20.);

/* ── 3. Build column assignment statements as a macro variable ──
        Each col_N is assigned ranuni(0) * 100 rounded to 4 decimals. */
%macro gen_assignments(n);
    %local i;
    %do i = 1 %to &n;
        col&i = round(ranuni(&seed) * 100, 0.0001);
    %end;
%mend gen_assignments;

/* ── 4. Generate data implicitly via DATA step ── */
data work.test2_result;
    length row_id 8;
    %gen_assignments(&ncols)
    do row_id = 1 to &nrows;
        %local j;
        %do j = 1 %to &ncols;
            col&j = round(ranuni(&seed) * 100, 0.0001);
        %end;
        output;
    end;
    drop _i_;
run;

/* ── 5. Implicit PROC SQL query against the SAS work dataset ── */
proc sql noprint;
    create table work.test2_summary as
    select
        count(*)        as total_rows,
        mean(col1)      as mean_col1,
        min(col1)       as min_col1,
        max(col1)       as max_col1,
        std(col1)       as std_col1
    from work.test2_result;
quit;

/* ── 6. Stop timer and report ── */
%let _t2_end     = %sysfunc(datetime());
%let _t2_elapsed = %sysevalf(&_t2_end - &_t2_start);

%put NOTE: [sas_test2] END   — %sysfunc(datetime(), datetime20.);
%put NOTE: [sas_test2] Elapsed: &_t2_elapsed seconds;

proc sql noprint;
    select count(*) into :obs trimmed from work.test2_result;
quit;
%put NOTE: [sas_test2] Observation count confirmed: &obs;
