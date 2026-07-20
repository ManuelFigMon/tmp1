/*==============================================================================
  Program : example_run_sas_jobs.sas
  Purpose : Dataset-driven demonstration of %run_sas_jobs.

            Step 1 builds a single SAS dataset (WORK.SASJOBS) whose rows ARE the
            job definitions - exactly the values supplied in the request:

               *NAME  -> jobname   (friendly name)
               -SYSIN -> sysin     (the .sas program to run)
               -LOG   -> log       (the .log path, timestamp written as {DTN})
               -NOSPLASH/-NOLOGO/-ICON -> nosplash / nologo / icon  (1 = add it)

            Step 2 simply calls the macro against that dataset.  Because the job
            list lives entirely in the dataset, adding / removing / editing jobs
            never requires touching the launcher code -> "dataset driven".

  Notes
  -----
    * {DTN} in each LOG path is replaced at run time with a YYYYMMDD_HHMMSS
      stamp, e.g.  PB_Pending_ADR_{DTN}.log -> PB_Pending_ADR_20260719_023040.log
    * DRYRUN=1 below just PRINTS the commands so this can be demonstrated on any
      platform.  Set DRYRUN=0 on the Windows SAS server to actually launch them.
==============================================================================*/

%include "run_sas_jobs.sas";      /* pull in the %run_sas_jobs / %run_sas_job macros */


/*--- Step 1 : the driver dataset -------------------------------------------*/
/*  nosplash / nologo / icon : 1 = include that switch, 0 = omit it.          */
/*  (Per the request only PB_Pending_ADR carries the three switches; set the  */
/*   flags to 1 on any other row if you want the same switches there too.)    */
data sasjobs;
    length jobname $60 sysin $300 log $320 nosplash nologo icon 8;

    infile datalines dsd dlm='|' truncover;
    input jobname $ sysin $ log $ nosplash nologo icon;
    datalines4;
PB_Pending_ADR|\\a70admed.com\R1\CGS\APPS\SAS\PROD\SAS_G\GSIT_PROD\PB\PB_Pending_ADR.sas|\\a70admed.com\R1\CGS\APPS\SAS\PROD\SAS_G\GSIT_PROD\PB\Logs\PB_Pending_ADR_{DTN}.log|1|1|1
PB_Pending_ADR_150_RSNAT|\\a70admed.com\R1\CGS\APPS\SAS\PROD\SAS_G\GSIT_PROD\PB\PB_Pending_ADR_150_RSNAT.sas|\\a70admed.com\R1\CGS\APPS\SAS\PROD\SAS_G\GSIT_PROD\PB\Logs\PB_Pending_ADR_150_RSNAT_{DTN}.log|0|0|0
PA_ADR_CERT_MRTS|\\a70admed.com\R1\CGS\APPS\SAS\PROD\SAS_G\GSIT_PROD\PA\PA_ADR_CERT_MRTS.sas|\\a70admed.com\R1\CGS\APPS\SAS\PROD\SAS_G\GSIT_PROD\PA\Logs\PA_ADR_CERT_MRTS_{DTN}.log|0|0|0
PA_OPDServices_Daily|\\a70admed.com\R1\CGS\APPS\SAS\PROD\SAS_G\GSIT_PROD\PA\PA_OPDServices_Daily.sas|\\a70admed.com\R1\CGS\APPS\SAS\PROD\SAS_G\GSIT_PROD\PA\Logs\PA_OPDServices_Daily_{DTN}.log|0|0|0
PB_Pending_Processor|\\a70admed.com\R1\CGS\APPS\SAS\PROD\SAS_G\GSIT_PROD\PB\PB_Pending_Processor.sas|\\a70admed.com\R1\CGS\APPS\SAS\PROD\SAS_G\GSIT_PROD\PB\Logs\PB_Pending_Processor_{DTN}.log|0|0|0
PA_SNFTPE_Weekly_Update|\\a70admed.com\R1\CGS\APPS\SAS\PROD\SAS_G\GSIT_PROD\PA\PA_SNFTPE_Weekly_Update.sas|\\a70admed.com\R1\CGS\APPS\SAS\PROD\SAS_G\GSIT_PROD\PA\Logs\PA_SNFTPE_Weekly_Update_{DTN}.log|0|0|0
PA_ProvDetail_Weekly|\\a70admed.com\R1\CGS\APPS\SAS\PROD\SAS_G\GSIT_PROD\PA\PA_ProvDetail_Weekly.sas|\\a70admed.com\R1\CGS\APPS\SAS\PROD\SAS_G\GSIT_PROD\PA\Logs\PA_ProvDetail_Weekly_{DTN}.log|0|0|0
PB_Pending_ADR_156_157_ASC|\\a70admed.com\R1\CGS\APPS\SAS\PROD\SAS_G\GSIT_PROD\PB\PB_Pending_ADR_156_157_ASC.sas|\\a70admed.com\R1\CGS\APPS\SAS\PROD\SAS_G\GSIT_PROD\PB\Logs\PB_Pending_ADR_156_157_ASC_{DTN}.log|0|0|0
;;;;
run;


/*--- Step 2 : launch every job in the dataset ------------------------------*/

/* (a) DEFAULT = SYSTASK COMMAND, asynchronous.  DRYRUN=1 to preview safely.  */
%run_sas_jobs(data=sasjobs, dryrun=1);

/* (b) Same jobs, but using the X command instead of SYSTASK:                 */
/* %run_sas_jobs(data=sasjobs, use_x=1, dryrun=1);                            */

/* (c) Real run on the Windows server (SYSTASK, fire-and-forget):             */
/* %run_sas_jobs(data=sasjobs, dryrun=0);                                     */

/* (d) Real run, wait for each job to finish before starting the next:        */
/* %run_sas_jobs(data=sasjobs, wait=1, dryrun=0);                             */

/* (e) Override the executable path:                                          */
/* %run_sas_jobs(data=sasjobs,                                                */
/*               exepath=D:\SASHome\SASFoundation\9.4\sas.exe, dryrun=0);      */
