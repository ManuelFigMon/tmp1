# SAS Log Performance Crawler — Architecture

How `Get-SasLogPerformance_v1.ps1`, its configuration CSVs, the SAS launcher,
and the outputs fit together.

## Deployment layout

```text
\\a70admed.com\R1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\code\ps\sas_crawler\
├── Get-SasLogPerformance_v1.ps1     the crawler (all logic, param() at top)
├── input_directories.csv            root folders to crawl   (FULL_PATH_DIR,NOTES)
├── folder_exclusion_list.csv        folders to skip         (EXCLUDE_FOLDER,NOTES)
└── robocopy_config.csv              optional staging jobs   (Source,Target,FileExtension,
                                     UpdateIfNewer,IncludeSubfolders,Notes)
                                     used only with -EnableRobocopy $true

Run_SasLogPerformance_v1.sas         SAS wrapper (can live anywhere SAS can read)
```

The three CSVs are found *next to the .ps1* by default; each location can be
overridden with `-InputDirectoriesCsv`, `-FolderExclusionCsv`,
`-RobocopyConfigCsv`. If `input_directories.csv` is missing, the crawler falls
back to the single `-LogRoot` path
(default `\\a70admed.com\R1\CGS\APPS\SAS\UNIT\SAS_G\SAS\Manuel\data\logs\UNIT`).

## System diagram

```mermaid
flowchart TB
    subgraph CALLERS["Launchers"]
        TS["Windows Task Scheduler<br/>powershell.exe -NoProfile -ExecutionPolicy Bypass -File ..."]
        SASW["SAS: Run_SasLogPerformance_v1.sas<br/>data _null_;  rc = system(...)"]
    end

    subgraph CONFIG["Configuration files (CSV, next to the .ps1)"]
        IN["input_directories.csv<br/>FULL_PATH_DIR, NOTES"]
        EXC["folder_exclusion_list.csv<br/>EXCLUDE_FOLDER, NOTES"]
        ROBO["robocopy_config.csv<br/>Source, Target, FileExtension,<br/>UpdateIfNewer, IncludeSubfolders, Notes<br/><i>(only when -EnableRobocopy $true)</i>"]
    end

    PS1["Get-SasLogPerformance_v1.ps1<br/><i>all behavior driven by param() block</i>"]

    LOGS[("SAS log directories<br/>\\\\a70admed.com\\...\\data\\logs\\UNIT<br/>*.log  *.txt  *.sas")]

    subgraph OUTPUTS["Outputs (folder of -OutputPath)"]
        OUT["sas_log_inventory.csv<br/>or .xlsx (sheet -ExcelSheetName)<br/>one row per SAS step"]
        RLOG["Get-SasLogPerformance.log<br/>rolling run log (5 MB roll)"]
    end

    MAIL["Completion / failure email<br/>Send-MailMessage via<br/>smtprelay.bcbssc.com:25"]

    RC["Exit code<br/>0 = success<br/>1 = parse errors<br/>2 = fatal"]

    TS --> PS1
    SASW --> PS1
    IN --> PS1
    EXC --> PS1
    ROBO --> PS1
    ROBO -. "robocopy staging (optional)" .-> LOGS
    LOGS --> PS1
    PS1 --> OUT
    PS1 --> RLOG
    PS1 --> MAIL
    PS1 --> RC
    RC --> SASW
```

## Internal pipeline (functions inside the .ps1)

```mermaid
flowchart TB
    START(["Start / param() defaults resolved"]) --> INITLOG["Initialize-RunLog<br/>open rolling Get-SasLogPerformance.log"]
    INITLOG --> ROBOQ{"-EnableRobocopy?"}
    ROBOQ -- "yes" --> COPY["Invoke-LogCopy<br/>robocopy per config row<br/>args as array, exit 0-7 = OK"]
    ROBOQ -- "no" --> DIRS
    COPY --> DIRS["Read-InputDirectories<br/>CSV FULL_PATH_DIR, else -LogRoot"]
    DIRS --> EXCL["Get-ExclusionList<br/>normalize excluded prefixes"]
    EXCL --> FIND["Get-LogFiles<br/>extension filter + exclusions"]
    FIND --> LOOP{"for each log file<br/>(own try/catch)"}

    LOOP --> PARSE["Parse-SasLog<br/>StreamReader Windows-1252"]
    PARSE --> MARKERS["step markers:<br/>PROCEDURE &lt;NAME&gt; used / DATA statement used /<br/>SAS initialization used / The SAS System used"]
    MARKERS --> CONV["ConvertTo-Seconds<br/>ss.ff | mm:ss.ff | h:mm:ss.ff -&gt; decimal"]
    MARKERS --> NAME["Get-SasProgramName<br/>schedule suffix -&gt; version suffix -&gt;<br/>literal tokens -&gt; trim _ and ."]
    CONV --> ROWS["rows: OK (per step) /<br/>NO_STEPS_FOUND (stub) /<br/>PARSE_ERROR (bad file, run continues)"]
    NAME --> ROWS
    ROWS --> LOOP

    LOOP -- "done" --> EXPORT["Export-Inventory<br/>.csv Export-Csv | .xlsx ImportExcel<br/>-&gt; Excel COM -&gt; CSV fallback"]
    EXPORT --> EMAIL["Send-CompletionEmail<br/>(Send-EmailAlert pattern)"]
    EMAIL --> EXIT(["exit 0 / 1 / 2"])
```

## File reference

| File | Role | Consumed by | Required |
|---|---|---|---|
| `Get-SasLogPerformance_v1.ps1` | Crawler/parser/exporter | Task Scheduler or SAS wrapper | yes |
| `Run_SasLogPerformance_v1.sas` | SAS launcher; maps exit code to NOTE/WARNING/ERROR, optional `abort return 2` | SAS batch / EG | optional |
| `input_directories.csv` | Root directories to crawl (`FULL_PATH_DIR`) | `Read-InputDirectories` | no — falls back to `-LogRoot` |
| `folder_exclusion_list.csv` | Folder prefixes to skip (`EXCLUDE_FOLDER`) | `Get-ExclusionList` | no |
| `robocopy_config.csv` | Staging copy jobs | `Invoke-LogCopy` | only with `-EnableRobocopy $true` |
| `sas_log_inventory.csv` / `.xlsx` | Output inventory (one row per SAS step) | reporting / Excel | produced |
| `Get-SasLogPerformance.log` | Rolling run log next to the output | troubleshooting | produced |
| `sample_logs/` | Synthetic SAS 9.4 fixtures for validation | testing only | no |
