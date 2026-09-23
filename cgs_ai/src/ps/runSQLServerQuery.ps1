<#
=====================================================================
  Program Name  : runSQLServerQuery.ps1
  Author        : Manuel Figallo
  Purpose       : Run a SQL Server query against a LOB catalog using
                  Windows Integrated Security and return the result set.
  Version       : 1.0beta
  Created       : 2026-08-26
  Last Modified : 2026-08-26

  Dependencies:
    NONE. Uses the built-in .NET System.Data.OleDb provider, so this works
    on a locked-down server where pip/Install-Module are not permitted.
    (The Python twin needs pyodbc -- prefer this version where installs are
    restricted.)

  Description:
    PowerShell twin of src/py/runSQLServerQuery.py, and the direct
    equivalent of this SAS pass-through block:

        connect to oledb as myconn
        (init_string="Provider=MSOLEDBSQL19;
         Integrated Security=SSPI;
         Persist Security Info=True;
         Initial Catalog=&LOB_Catalog;
         Data Source=my.dbserver.CCOM, 1433");
        select * from connection to myconn (&SQL_Statement);

    Integrated Security is used, so NO password is handled, stored or logged.

  Input Parameters (required first):
    -SQL_Statement (REQUIRED)  -LOB_Catalog (REQUIRED)
    -DataSource (default my.dbserver.CCOM,1433 or SQL_DATA_SOURCE from .env)
    -OutputCsvPath (optional)  -Provider (default MSOLEDBSQL19)
  Exit codes: 0 = success, 2 = config error, 3 = connection/query error.
=====================================================================
#>
[CmdletBinding()]
param(
    [string] $SQL_Statement  = '',
    [string] $LOB_Catalog    = '',
    [string] $DataSource     = '',
    [string] $OutputCsvPath  = '',
    [string] $Provider       = 'MSOLEDBSQL19',
    [int]    $TimeoutSeconds = 300
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\cgsUtils.ps1"

function Get-CgsConnectionString {
    <# .SYNOPSIS Build the OLE DB init string (Integrated Security, no password).
       .OUTPUTS  [string] -- contains no credentials, safe to log. #>
    param([string] $Catalog, [string] $Source, [string] $ProviderName)
    return ("Provider=$ProviderName;Integrated Security=SSPI;" +
            "Persist Security Info=True;Initial Catalog=$Catalog;Data Source=$Source")
}

function Invoke-Main {
    <# .SYNOPSIS Connect, execute, return rows. .OUTPUTS [int] exit code. #>
    if (-not $SQL_Statement) { Write-CgsError "required parameter 'SQL_Statement' is missing or empty"; return 2 }
    if (-not $LOB_Catalog)   { Write-CgsError "required parameter 'LOB_Catalog' is missing or empty"; return 2 }
    $source = if ($DataSource) { $DataSource } else { Get-CgsConfig -Key 'SQL_DATA_SOURCE' -Default 'my.dbserver.CCOM,1433' }

    Add-Type -AssemblyName System.Data
    $connectionString = Get-CgsConnectionString -Catalog $LOB_Catalog -Source $source -ProviderName $Provider
    Write-CgsInfo "connecting to catalog '$LOB_Catalog' on $source (Integrated Security)"

    $connection = New-Object System.Data.OleDb.OleDbConnection $connectionString
    $table = New-Object System.Data.DataTable
    try {
        $connection.Open()
        $command = $connection.CreateCommand()
        $command.CommandText    = $SQL_Statement
        $command.CommandTimeout = $TimeoutSeconds
        $adapter = New-Object System.Data.OleDb.OleDbDataAdapter $command
        [void]$adapter.Fill($table)
    } finally {
        if ($connection.State -eq 'Open') { $connection.Close() }
    }

    $columns = @($table.Columns | ForEach-Object { $_.ColumnName })
    Write-CgsInfo ("returned {0} row(s) x {1} column(s)" -f $table.Rows.Count, $columns.Count)

    if ($OutputCsvPath) {
        $rows = foreach ($row in $table.Rows) {
            $item = [ordered]@{}
            foreach ($column in $columns) { $item[$column] = $row[$column] }
            [PSCustomObject]$item
        }
        [void](Write-CgsCsv -Rows @($rows) -Columns $columns -Target $OutputCsvPath)
        Write-CgsInfo "wrote result set to $OutputCsvPath"
    }
    # Emit the rows so a caller can pipe them.
    $table | Write-Output
    return 0
}
try { exit (Invoke-Main) }
catch { Write-CgsError ("query failed: {0}" -f $_.Exception.Message); exit 3 }
