/***********************************************************************
  Program  : sas_test1.sas
  Purpose  : Explicit PROC SQL passthrough — generates &nrows x &ncols
             random numeric data inside SQL Server via OLEDB and returns
             it to a SAS work dataset. Timed end-to-end.
  Method   : Explicit passthrough (SQL executes on the remote DBMS)
***********************************************************************/

/* ── 1. Configurable parameters ── */
%let nrows   = 1000;    /* number of rows to generate            */
%let ncols   = 30;      /* number of random numeric columns       */
%let seed    = 12345;   /* not used in SQL Server NEWID approach  */

/* ── 2. Build the SELECT column list as a macro variable ──
        Resolved here so the string is passed literally to SQL Server. */
%macro build_collist(n);
    %local i result;
    %let result = ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS row_id;
    %do i = 1 %to &n;
        %let result = &result
            , CAST(ABS(CHECKSUM(NEWID())) % 10000 / 100.0 AS DECIMAL(10,4)) AS col&i;
    %end;
    &result
%mend build_collist;

%let collist = %build_collist(&ncols);

/* ── 3. Start timer ── */
%let _t1_start = %sysfunc(datetime());
%put NOTE: [sas_test1] START — %sysfunc(datetime(), datetime20.);

/* ── 4. Explicit passthrough ── */
proc sql noprint;
    connect to oledb as myconn
        (init_string = "Provider=MSOLEDBSQL19;
                        Integrated Security=SSPI;
                        Persist Security Info=True;
                        Initial Catalog=DataMartKYA;
                        Data Source=A70TUCGSDATA001.A70ADMED.COM, 1433");

    create table work.test1_result as
    select * from connection to myconn
    (
        WITH n AS (
            SELECT TOP &nrows
                   ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS rn
            FROM   sys.all_columns a
                   CROSS JOIN sys.all_columns b
        )
        SELECT &collist
        FROM   n
    );

    disconnect from myconn;
quit;

/* ── 5. Stop timer and report ── */
%let _t1_end     = %sysfunc(datetime());
%let _t1_elapsed = %sysevalf(&_t1_end - &_t1_start);

%put NOTE: [sas_test1] END   — %sysfunc(datetime(), datetime20.);
%put NOTE: [sas_test1] Elapsed: &_t1_elapsed seconds;
%put NOTE: [sas_test1] Rows returned: %sysfunc(attrn(%sysfunc(open(work.test1_result)), nobs));

/* ── 6. Quick validation ── */
proc sql noprint;
    select count(*) into :obs trimmed from work.test1_result;
quit;
%put NOTE: [sas_test1] Observation count confirmed: &obs;
