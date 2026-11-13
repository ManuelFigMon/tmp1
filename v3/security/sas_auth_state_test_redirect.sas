/*=============================================================
  Stored Process: sas_auth_state_test_redirect.sas
  Purpose : Authenticate the user via SAS Web Auth, capture
            &_USERNAME, ENSURE the user's profile/data folder
            exists (meta/profiles/<userid>/data), then redirect
            back to the originating HTML portal page with the
            userid as a URL parameter.

  Registration (SAS Management Console):
    - Output type : Streaming
    - Result type : (leave blank -- we emit raw HTTP headers)

  URL called by the HTML page:
    /SASStoredProcess/do?_program=/YourFolder/sas_auth_state_test_redirect
        &returnUrl=https://yourserver/.../sas_auth_state_test.html

  Companion files:
    - sas_auth_state_test.html        (the portal page)
    - sas_auth_state_test_api.sas     (reads/writes the datastore)
    - meta/profiles/<userid>/data/sas_auth_state_test_state.json
=============================================================*/


/* ----------------------------------------------------------
   0.  CONFIG — server-side root of the profile store.
       This is the filesystem directory that the web path
       meta/profiles/ maps to on the SAS content server.
       Adjust to your deployment.
---------------------------------------------------------- */
%let profile_root = /custom/projects/dmq/temp/v3/meta/profiles;


/* ----------------------------------------------------------
   1.  Read the return URL passed in by the HTML page.
       %superq() avoids special-character resolution problems.
---------------------------------------------------------- */
%let default_return = https://cgssaswebu.a70admed.com/custom/projects/dmq/temp/v3/security/sas_auth_state_test.html;

%let return_url = %superq(returnUrl);
%if %length(&return_url.) = 0 %then %do;
  %let return_url = &default_return.;
%end;


/* ----------------------------------------------------------
   2.  Capture and URL-encode the authenticated userid.
       &_USERNAME is set automatically by SAS for every
       authenticated Stored Process call. compress() strips
       any DOMAIN\ prefix-friendly whitespace; we also strip a
       leading domain and trailing @realm for the folder name.
---------------------------------------------------------- */
%let raw_user   = &_USERNAME.;
%let clean_user = %sysfunc(compress(&raw_user.));

/* derive a filesystem-safe userid: drop DOMAIN\ and @realm */
data _null_;
  u = symget('clean_user');
  /* strip everything up to and including a backslash or slash */
  if index(u,'\')>0 then u = substr(u, index(u,'\')+1);
  if index(u,'/')>0 then u = substr(u, index(u,'/')+1);
  /* strip @realm */
  if index(u,'@')>0 then u = substr(u, 1, index(u,'@')-1);
  /* keep only safe characters */
  u = prxchange('s/[^A-Za-z0-9_\-]//', -1, strip(u));
  call symputx('safe_user', u, 'G');
run;

%let enc_user = %sysfunc(urlencode(&safe_user.));


/* ----------------------------------------------------------
   3.  Login timestamp passed back to the HTML page.
---------------------------------------------------------- */
%let login_ts = %sysfunc(datetime(), datetime20.);
%let enc_ts   = %sysfunc(urlencode(&login_ts.));


/* ----------------------------------------------------------
   4.  Ensure meta/profiles/<userid>/data exists so the API
       stored process has somewhere to persist state.
       dcreate() is a no-op-safe create (we guard with fileexist).
---------------------------------------------------------- */
data _null_;
  base   = "&profile_root.";
  uid    = "&safe_user.";
  udir   = catx('/', base, uid);
  ddir   = catx('/', udir, 'data');

  if not fileexist(udir) then rc1 = dcreate(uid, base);
  if not fileexist(ddir) then rc2 = dcreate('data', udir);

  put "NOTE: [sas_auth_state_test_redirect] ensured folder: " ddir;
run;


/* ----------------------------------------------------------
   5.  Build the redirect URL (data step avoids macro quoting
       issues with the "?" vs "&" separator on the URL).
---------------------------------------------------------- */
data _null_;
  return_url = symget('return_url');
  if index(return_url, '?') > 0 then sep = '&'; else sep = '?';
  enc_user = symget('enc_user');
  enc_ts   = symget('enc_ts');
  redirect_url = strip(return_url)
              || strip(sep)
              || 'userid='     || strip(enc_user)
              || '&loginTime=' || strip(enc_ts)
              || '&sasAuth=1';
  call symputx('redirect_url', redirect_url, 'G');
run;


/* ----------------------------------------------------------
   6.  Emit HTTP 302 redirect (close ODS so SAS does not wrap
       our raw headers in HTML).
---------------------------------------------------------- */
ods listing close;
ods html    close;

data _null_;
  file _webout;
  put 'HTTP/1.0 302 Moved Temporarily';
  put 'Content-type: text/html; charset=utf-8';
  put "Location: &redirect_url.";
  put 'Cache-Control: no-store, no-cache, must-revalidate';
  put 'Pragma: no-cache';
  put '';
  put '<!DOCTYPE html><html><head>';
  put "<meta http-equiv=""refresh"" content=""0;url=&redirect_url."">";
  put '</head><body>';
  put '<p style="font-family:sans-serif">Completing sign-in, please wait...</p>';
  put "<p><a href=""&redirect_url."">Click here if not redirected.</a></p>";
  put '</body></html>';
run;


/* ----------------------------------------------------------
   7.  Server-side audit log.
---------------------------------------------------------- */
%put NOTE: ============================================================;
%put NOTE: [sas_auth_state_test_redirect] User authenticated : &_USERNAME.;
%put NOTE: [sas_auth_state_test_redirect] Safe userid        : &safe_user.;
%put NOTE: [sas_auth_state_test_redirect] Login timestamp    : &login_ts.;
%put NOTE: [sas_auth_state_test_redirect] Redirecting to     : &redirect_url.;
%put NOTE: ============================================================;
