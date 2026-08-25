/*=====================================================================
  Program Name : Find_python_exe.sas
  Author       : Manuel Figallo
  Purpose      : Locate python.exe from SAS and set a macro variable
                 (default PYTHON_EXE) that %scanFileSystem() can use.
  Version      : 1.3.3

  IMPORTANT - which python.exe do you actually want?
    If you installed the scanner's dependencies into a virtual environment
    (pip install -r requirements.txt inside .venv), you must point at the
    VENV interpreter:

        <project>\.venv\Scripts\python.exe

    NOT the system-wide python.exe that "where python" finds first. The
    system interpreter will not have openpyxl, so .xlsx output would
    silently fall back to CSV. %findPython() checks the venv FIRST for
    exactly this reason.

  Usage:
      %include "<path>\sas\Find_python_exe.sas";
      %findPython(project_root=C:\code\python\cgs_ai\scanFileSystem);
      %put &=PYTHON_EXE;

  Notes:
    FILENAME PIPE and X require the XCMD system option. If your site runs
    with NOXCMD, the pipe probes are skipped automatically and only the
    filesystem probes run (those always work).
=====================================================================*/


%macro findPython(
    project_root =,              /* optional: folder containing .venv     */
    out          = PYTHON_EXE,   /* macro variable to create              */
    validate     = 1             /* 1 = check openpyxl is importable      */
);

    %global &out;
    %local xcmd_ok cand i n;
    %let &out = ;
    %let xcmd_ok = %eval(%sysfunc(getoption(xcmd)) = XCMD);

    %if not &xcmd_ok %then
        %put NOTE: XCMD is disabled at this site; using filesystem probes only.;

    /*-----------------------------------------------------------------
      PROBE 1 - the project virtual environment (preferred).
    -----------------------------------------------------------------*/
    %if %superq(project_root) ne %then %do;
        %let cand = %superq(project_root)\.venv\Scripts\python.exe;
        %if %sysfunc(fileexist(%superq(cand))) %then %do;
            %let &out = %superq(cand);
            %put NOTE: Found venv interpreter: &&&out;
        %end;
    %end;

    /*-----------------------------------------------------------------
      PROBE 2 - "where python.exe" (first match on PATH).
    -----------------------------------------------------------------*/
    %if %superq(&out) = and &xcmd_ok %then %do;
        filename _fpwhere pipe 'where python.exe 2>&1' lrecl=1024;
        data _null_;
            infile _fpwhere truncover;
            input line $1024.;
            /* a real hit looks like a path: has a colon and ends in .exe */
            if index(line, ':') and upcase(scan(line, -1, '.')) = 'EXE' then do;
                call symputx("&out", strip(line), 'G');
                stop;
            end;
        run;
        filename _fpwhere clear;
        %if %superq(&out) ne %then %put NOTE: Found on PATH: &&&out;
    %end;

    /*-----------------------------------------------------------------
      PROBE 3 - the Windows Python launcher: "py -0p" lists every
                installed version and its path.
    -----------------------------------------------------------------*/
    %if %superq(&out) = and &xcmd_ok %then %do;
        filename _fppy pipe 'py -0p 2>&1' lrecl=1024;
        data _null_;
            infile _fppy truncover;
            input line $1024.;
            if index(upcase(line), 'PYTHON.EXE') then do;
                /* the path is the last whitespace-delimited token */
                call symputx("&out", strip(scan(line, -1, ' ')), 'G');
                stop;
            end;
        run;
        filename _fppy clear;
        %if %superq(&out) ne %then %put NOTE: Found via py launcher: &&&out;
    %end;

    /*-----------------------------------------------------------------
      PROBE 4 - common install locations.
    -----------------------------------------------------------------*/
    %if %superq(&out) = %then %do;
        %let n = 0;
        %do i = 313 %to 39 %by -1;   /* 3.13 down to 3.9 */
            %if %superq(&out) = %then %do;
                %local v; %let v = %substr(&i, 1, 1)%substr(&i, 2);
                %let cand = C:\Python&v\python.exe;
                %if %sysfunc(fileexist(%superq(cand))) %then %let &out = %superq(cand);

                %if %superq(&out) = %then %do;
                    %let cand = %sysget(LOCALAPPDATA)\Programs\Python\Python&v\python.exe;
                    %if %sysfunc(fileexist(%superq(cand))) %then %let &out = %superq(cand);
                %end;

                %if %superq(&out) = %then %do;
                    %let cand = C:\Program Files\Python&v\python.exe;
                    %if %sysfunc(fileexist(%superq(cand))) %then %let &out = %superq(cand);
                %end;
            %end;
        %end;
        %if %superq(&out) ne %then %put NOTE: Found in a standard location: &&&out;
    %end;

    /*-----------------------------------------------------------------
      Report / validate.
    -----------------------------------------------------------------*/
    %if %superq(&out) = %then %do;
        %put ERROR: Could not locate python.exe.;
        %put ERROR- Run this in a terminal where Python works, and paste the result:;
        %put ERROR-     python -c "import sys; print(sys.executable)";
        %return;
    %end;

    %put NOTE: &out = %superq(&out);

    %if &validate and &xcmd_ok %then %do;
        %local vrc;
        systask command """%superq(&out)"" -c ""import sys, openpyxl; print(sys.version)"""
                taskname=_pyck status=_pyrc wait;
        waitfor _pyck;
        %let vrc = &_pyrc;
        %if &vrc = 0 %then
            %put NOTE: Interpreter works and openpyxl is available (.xlsx output OK).;
        %else %do;
            %put WARNING: %superq(&out) could not import openpyxl (rc=&vrc).;
            %put WARNING- .xlsx output will fall back to CSV. Install it with:;
            %put WARNING-     "%superq(&out)" -m pip install openpyxl;
        %end;
    %end;

%mend findPython;


/*=====================================================================
  EXAMPLES - un-comment one and submit.
=====================================================================*/

/* A. Prefer the project venv, fall back to PATH / launcher / standard dirs */
%*
%findPython(project_root=C:\code\python\cgs_ai\scanFileSystem);
%put &=PYTHON_EXE;
*;

/* B. Just search the system (no venv) */
%*
%findPython();
%put &=PYTHON_EXE;
*;

/* C. Wire it straight into the scanner wrapper */
%*
%include "C:\code\python\cgs_ai\scanFileSystem\sas\Run_scanFileSystem_v1.sas";
%findPython(project_root=C:\code\python\cgs_ai\scanFileSystem);
%scanFileSystem(
  input_folder_root=%str(\\A70admed.com\r1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT\DME\Logs),
  output_file_path=C:\code\python\cgs_ai\tests\scanFileSystem\scan.xlsx,
  metric_profile=sas_log
);
*;


/*=====================================================================
  ONE-LINERS you can paste directly (no macro needed)
=====================================================================*/

/* Show every python.exe on PATH */
%*
filename _p pipe 'where python.exe 2>&1';
data _null_; infile _p truncover; input l $256.; put l=; run;
filename _p clear;
*;

/* Show every installed Python and its path (Windows Python launcher) */
%*
filename _p pipe 'py -0p 2>&1';
data _null_; infile _p truncover; input l $256.; put l=; run;
filename _p clear;
*;

/* Ask a specific interpreter where it lives and what it can import */
%*
filename _p pipe '"C:\code\python\cgs_ai\scanFileSystem\.venv\Scripts\python.exe" -c "import sys,openpyxl;print(sys.executable);print(sys.version)" 2>&1';
data _null_; infile _p truncover; input l $256.; put l=; run;
filename _p clear;
*;

/* Is XCMD even enabled at this site? (needed for PIPE and X) */
%*
%put %sysfunc(getoption(xcmd));
*;
