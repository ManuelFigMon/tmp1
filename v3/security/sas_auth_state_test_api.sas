/*=============================================================
  Stored Process: sas_auth_state_test_api.sas
  Purpose : Read / write a per-user JSON datastore that backs the
            data-driven dropdown in sas_auth_state_test.html.

            The store lives at:
              meta/profiles/<userid>/data/sas_auth_state_test_state.json

            Shape:
              { "userid":"731o",
                "items":[ {"id":1,"value":"..."}, ... ],
                "updated":"02FEB2026:09:15:00" }

  Registration (SAS Management Console):
    - Output type : Streaming
    - Result type : Streaming (we emit application/json)
    - Grant the stored-process identity read/write on profile_root.

  Query parameters:
    action = read | save           (default: read)
    userid = <folder-safe userid>  (falls back to &_USERNAME)
    items  = <URL-encoded JSON array of {id,value}>   (action=save only)

  Examples:
    .../do?_program=/.../sas_auth_state_test_api&action=read&userid=731o
    .../do?_program=/.../sas_auth_state_test_api&action=save&userid=731o
           &items=%5B%7B%22id%22%3A1%2C%22value%22%3A%22Region%3A%20Midwest%22%7D%5D

  SECURITY NOTES (read before production use):
    - In production, IGNORE the client-supplied &userid and trust ONLY
      &_USERNAME so a user cannot read/write another user's store.
      The userid parameter here is a convenience for offline testing.
    - Validate/escape &items; this demo writes the received JSON text
      after a light bracket check. Harden before exposing externally.
=============================================================*/


/* ----------------------------------------------------------
   0.  CONFIG — must match sas_auth_state_test_redirect.sas
---------------------------------------------------------- */
%let profile_root = /custom/projects/dmq/temp/v3/meta/profiles;
%let state_file   = sas_auth_state_test_state.json;


/* ----------------------------------------------------------
   1.  Resolve action + a filesystem-safe userid.
       PRODUCTION: replace the userid fallback chain with
                   %let safe_user = <derived strictly from &_USERNAME>;
---------------------------------------------------------- */
%global action userid items;
%if %superq(action) = %then %let action = read;

%let src_user = %superq(userid);
%if %length(&src_user.) = 0 %then %let src_user = &_USERNAME.;

data _null_;
  u = symget('src_user');
  if index(u,'\')>0 then u = substr(u, index(u,'\')+1);
  if index(u,'/')>0 then u = substr(u, index(u,'/')+1);
  if index(u,'@')>0 then u = substr(u, 1, index(u,'@')-1);
  u = prxchange('s/[^A-Za-z0-9_\-]//', -1, strip(u));
  call symputx('safe_user', u, 'G');
run;

%let udir  = &profile_root./&safe_user.;
%let ddir  = &udir./data;
%let fpath = &ddir./&state_file.;


/* ----------------------------------------------------------
   2.  Make sure the target folder exists (idempotent).
---------------------------------------------------------- */
data _null_;
  base="&profile_root."; uid="&safe_user."; udir="&udir."; ddir="&ddir.";
  if not fileexist(udir) then rc1 = dcreate(uid, base);
  if not fileexist(ddir) then rc2 = dcreate('data', udir);
run;


/* Close ODS so we control the raw stream. */
ods listing close;
ods html    close;


/* ==========================================================
   ACTION: SAVE
   ========================================================== */
%macro do_save;
  /* Pull the raw items JSON the client sent (already URL-decoded
     by SAS into the &items macro var). Guard against empties. */
  %let items_in = %superq(items);
  %if %length(&items_in.) = 0 %then %let items_in = [];

  /* Light validation: must start with '[' and end with ']'. */
  data _null_;
    s = strip(symget('items_in'));
    ok = (first(s)='[') and (substr(s,length(s),1)=']');
    call symputx('items_ok', ok, 'G');
  run;

  %if &items_ok. = 0 %then %do;
    data _null_;
      file _webout;
      put 'Content-type: application/json; charset=utf-8';
      put 'Cache-Control: no-store';
      put '';
      put '{"status":"error","message":"items must be a JSON array"}';
    run;
    %return;
  %end;

  /* Write the store file: { userid, items, updated }. We stream the
     received array text verbatim between the wrapper braces. */
  %let now_ts = %sysfunc(datetime(), datetime20.);
  data _null_;
    length line $32767;
    file "&fpath." lrecl=32767;
    put '{';
    put "  ""userid"": ""&safe_user."",";
    line = '  "items": ' || strip(symget('items_in')) || ',';
    put line;
    put "  ""updated"": ""&now_ts.""";
    put '}';
  run;

  /* Echo success + the saved payload back to the caller. */
  data _null_;
    length line $32767;
    file _webout lrecl=32767;
    put 'Content-type: application/json; charset=utf-8';
    put 'Cache-Control: no-store';
    put '';
    put '{';
    put "  ""status"": ""saved"",";
    put "  ""userid"": ""&safe_user."",";
    line = '  "items": ' || strip(symget('items_in')) || ',';
    put line;
    put "  ""updated"": ""&now_ts.""";
    put '}';
  run;
%mend do_save;


/* ==========================================================
   ACTION: READ
   ========================================================== */
%macro do_read;
  %if %sysfunc(fileexist(&fpath.)) %then %do;
    /* Stream the existing file straight back. */
    data _null_;
      length line $32767;
      infile "&fpath." lrecl=32767 pad end=eof;
      file _webout lrecl=32767;
      if _n_=1 then do;
        put 'Content-type: application/json; charset=utf-8';
        put 'Cache-Control: no-store';
        put '';
      end;
      input;
      line = _infile_;
      put line;
    run;
  %end;
  %else %do;
    /* No store yet — return an empty, well-formed document. */
    data _null_;
      file _webout;
      put 'Content-type: application/json; charset=utf-8';
      put 'Cache-Control: no-store';
      put '';
      put '{';
      put "  ""userid"": ""&safe_user."",";
      put '  "items": [],';
      put '  "updated": null';
      put '}';
    run;
  %end;
%mend do_read;


/* ----------------------------------------------------------
   3.  Dispatch.
---------------------------------------------------------- */
%macro dispatch;
  %if %lowcase(&action.) = save %then %do; %do_save; %end;
  %else %do; %do_read; %end;
%mend dispatch;
%dispatch;


/* ----------------------------------------------------------
   4.  Audit log.
---------------------------------------------------------- */
%put NOTE: ============================================================;
%put NOTE: [sas_auth_state_test_api] action   : &action.;
%put NOTE: [sas_auth_state_test_api] _USERNAME: &_USERNAME.;
%put NOTE: [sas_auth_state_test_api] safe_user: &safe_user.;
%put NOTE: [sas_auth_state_test_api] file     : &fpath.;
%put NOTE: ============================================================;
