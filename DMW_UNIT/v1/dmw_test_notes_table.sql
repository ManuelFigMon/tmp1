/*=============================================================================
  dmw_test_notes_table.sql                                           *** TEST ***
  -----------------------------------------------------------------------------
  DMW (Data Mart Workbench) -- IT Management Solution
  Job Notes feature: table definition and grants.

  SEND THIS TO THE DBA. It is the preferred route.

  Use dmw_test_create_notes_table.sas only if you cannot get a DBA to run this
  and the service account happens to hold CREATE TABLE. The SAS route exists
  because the service account is the only identity you can invoke, not because
  it is the better way to create a production table.

  Target : A70TUCGSDATA012.A70ADMED.COM
  Database: OHM_CGSScheduler
  Runs as : a DBA / db_owner. NOT the application service account.

  -----------------------------------------------------------------------------
  WHAT THIS CREATES AND WHY IT IS SHAPED THIS WAY
  -----------------------------------------------------------------------------
  One row per note, append-only. Notes are never updated in place and never
  hard-deleted -- a retraction sets IsRetracted and keeps the original text.

  Append-only removes three problems at once:
    * No lost updates when two people type a note on the same job at once.
    * No row-level concurrency control needed in the application.
    * A complete audit trail, free.

  CreatedBy holds the SAS _METAPERSON value -- the person signed in to the web
  application. It is written server-side by the stored process and is NEVER
  accepted from the browser, because a browser-supplied username can be forged.
  All database writes arrive as the shared service account, so this column is
  the only thing that makes a note attributable to a person.
=============================================================================*/

USE [OHM_CGSScheduler];
GO

/*-----------------------------------------------------------------------------
  OPTIONAL BUT RECOMMENDED -- give DMW its own schema.
  OHM_CGSScheduler is an operational database. Keeping application objects out
  of dbo makes it obvious which objects belong to DMW and lets you grant on the
  schema rather than object by object.

  If you use this, change DMW_SCHEMA in the .sas files from dbo to dmw and
  change [dbo] to [dmw] throughout this script.
-----------------------------------------------------------------------------*/
-- IF NOT EXISTS (SELECT 1 FROM sys.schemas WHERE name = 'dmw')
--     EXEC('CREATE SCHEMA [dmw] AUTHORIZATION [dbo]');
-- GO


/*-----------------------------------------------------------------------------
  1. The table
-----------------------------------------------------------------------------*/
IF OBJECT_ID('[dbo].[DMW_JobNotes]','U') IS NULL
BEGIN
    CREATE TABLE [dbo].[DMW_JobNotes] (
        NoteID            int IDENTITY(1,1) NOT NULL,

        /* The job the note is about. Durable knowledge: "this job always fails
           on month-end close, that is expected." Survives every run.          */
        JobID             int               NOT NULL,

        /* Optional: pins the note to ONE execution. Incident-specific:
           "run 4471 failed because the upstream file was late." Nullable, so a
           note can be about the job in general. Costs nothing now and is
           painful to retrofit later.                                          */
        RequestID         int                   NULL,

        NoteText          nvarchar(2000)    NOT NULL,

        /* SAS _METAPERSON. Written server-side. Never from the browser.       */
        CreatedBy         varchar(64)       NOT NULL,
        CreatedDatetime   datetime2(0)      NOT NULL
                          CONSTRAINT DF_DMW_JobNotes_Created   DEFAULT (SYSDATETIME()),

        /* Soft delete. People retract notes they regret, and you want that
           visible rather than silently gone.                                  */
        IsRetracted       bit               NOT NULL
                          CONSTRAINT DF_DMW_JobNotes_Retracted DEFAULT (0),
        RetractedBy       varchar(64)           NULL,
        RetractedDatetime datetime2(0)          NULL,

        CONSTRAINT PK_DMW_JobNotes PRIMARY KEY CLUSTERED (NoteID)
    );
END
GO

/*-----------------------------------------------------------------------------
  2. Index -- every read is "notes for this job, newest first"
-----------------------------------------------------------------------------*/
IF NOT EXISTS (SELECT 1 FROM sys.indexes
               WHERE name = 'IX_DMW_JobNotes_JobID'
                 AND object_id = OBJECT_ID('[dbo].[DMW_JobNotes]'))
BEGIN
    CREATE INDEX [IX_DMW_JobNotes_JobID]
        ON [dbo].[DMW_JobNotes] (JobID, CreatedDatetime DESC);
END
GO

/*-----------------------------------------------------------------------------
  3. Grants -- the narrowest set the feature needs
  -----------------------------------------------------------------------------
  SELECT  read the thread
  INSERT  add a note
  UPDATE  retract a note (sets IsRetracted only)

  DELETE is deliberately NOT granted: nothing in the application removes rows.

  This grant is on THIS TABLE ONLY. The service account keeps its read-only
  access to View_Scheduler_PBI_Dashboard and gains no other write capability
  anywhere in the database.
-----------------------------------------------------------------------------*/
IF NOT EXISTS (SELECT 1 FROM sys.database_principals
               WHERE name = 'A70ADMED\SVC_CGSUSASSRV')
BEGIN
    CREATE USER [A70ADMED\SVC_CGSUSASSRV] FOR LOGIN [A70ADMED\SVC_CGSUSASSRV];
END
GO

GRANT SELECT, INSERT, UPDATE ON [dbo].[DMW_JobNotes] TO [A70ADMED\SVC_CGSUSASSRV];
GO


/*-----------------------------------------------------------------------------
  4. Verification -- run this and eyeball the output before telling anyone
     the feature is ready
-----------------------------------------------------------------------------*/
SELECT  TableExists  = CASE WHEN OBJECT_ID('[dbo].[DMW_JobNotes]','U') IS NULL
                            THEN 0 ELSE 1 END,
        ColumnCount  = (SELECT COUNT(*) FROM sys.columns
                         WHERE object_id = OBJECT_ID('[dbo].[DMW_JobNotes]')),
        IndexExists  = CASE WHEN EXISTS (SELECT 1 FROM sys.indexes
                                          WHERE name = 'IX_DMW_JobNotes_JobID'
                                            AND object_id = OBJECT_ID('[dbo].[DMW_JobNotes]'))
                            THEN 1 ELSE 0 END;

SELECT  grantee     = USER_NAME(dp.grantee_principal_id),
        permission  = dp.permission_name,
        state       = dp.state_desc
FROM    sys.database_permissions dp
WHERE   dp.major_id = OBJECT_ID('[dbo].[DMW_JobNotes]')
ORDER BY grantee, permission;
GO


/*-----------------------------------------------------------------------------
  ROLLBACK -- if the pilot is abandoned
-----------------------------------------------------------------------------*/
-- DROP TABLE [dbo].[DMW_JobNotes];
-- GO
