function Export-DbaCsv {
    <#
    .SYNOPSIS
        Efficiently export very large (and small) table data or querie to CSV file.

    .DESCRIPTION
        Import-DbaCsv takes advantage of .NET's super fast SqlBulkCopy class to import CSV files into SQL Server.

        The entire import is performed within a transaction, so if a failure occurs or the script is aborted, no changes will persist.

        If the table or view specified does not exist and -AutoCreateTable, it will be automatically created using slow and inefficient but accommodating data types.

        This importer supports fields spanning multiple lines. The only restriction is that they must be quoted, otherwise it would not be possible to distinguish between malformed data and multi-line values.

        Able to read gzip compressed CSV files if the filename ends with ".csv.gz"

    .PARAMETER Path
        Specifies path to the CSV file(s) to be imported. Multiple files may be imported at once.

    .PARAMETER NoHeaderRow
        By default, the first row is used to determine column names for the data being imported.

        Use this switch if the first row contains data and not column names.

    .PARAMETER Delimiter
        Specifies the delimiter used in the imported file(s). If no delimiter is specified, comma is assumed.

        Valid delimiters are '`t`, '|', ';',' ' and ',' (tab, pipe, semicolon, space, and comma).

    .PARAMETER SingleColumn
        Specifies that the file contains a single column of data. Otherwise, the delimiter check bombs.

    .PARAMETER SqlInstance
        The SQL Server Instance to import data into.

    .PARAMETER SqlCredential
        Login to the target instance using alternative credentials. Accepts PowerShell credentials (Get-Credential).

        Windows Authentication, SQL Server Authentication, Active Directory - Password, and Active Directory - Integrated are all supported.

        For MFA support, please use Connect-DbaInstance.

    .PARAMETER Database
        Specifies the name of the database the CSV will be imported into. Options for this this parameter are  auto-populated from the server.

    .PARAMETER Schema
        Specifies the schema in which the SQL table or view where CSV will be imported into resides. Default is dbo.

        If a schema does not currently exist, it will be created, after a prompt to confirm this. Authorization will be set to dbo by default.

        This parameter overrides -UseFileNameForSchema.

    .PARAMETER Table
        Specifies the SQL table or view where CSV will be imported into.

        If a table name is not specified, the table name will be automatically determined from the filename.

        If the table specified does not exist and -AutoCreateTable, it will be automatically created using slow and inefficient but accommodating data types.

        If the automatically generated table datatypes do not work for you, please create the table prior to import.

        If you want to import specific columns from a CSV, create a view with corresponding columns.

    .PARAMETER Column
        Import only specific columns. To remap column names, use the ColumnMap.

    .PARAMETER ColumnMap
        By default, the bulk copy tries to automap columns. When it doesn't work as desired, this parameter will help. Check out the examples for more information.

    .PARAMETER KeepOrdinalOrder
        By default, the importer will attempt to map exact-match columns names from the source document to the target table. Using this parameter will keep the ordinal order instead.

    .PARAMETER AutoCreateTable
        Creates a table if it does not already exist. The table will be created with sub-optimal data types such as nvarchar(max)

    .PARAMETER Truncate
        If this switch is enabled, the destination table will be truncated prior to import.

    .PARAMETER NotifyAfter
        Specifies the import row count interval for reporting progress. A notification will be shown after each group of this many rows has been imported.

    .PARAMETER BatchSize
        Specifies the batch size for the import. Defaults to 50000.

    .PARAMETER UseFileNameForSchema
        If this switch is enabled, the script will try to find the schema name in the input file by looking for a period (.) in the file name.

        If used with the -Table parameter you may still specify the target table name. If -Table is not used the file name after the first period will
        be used for the table name.

        For example test.data.csv will import the csv contents to a table in the test schema.

        If it finds one it will use the file name up to the first period as the schema. If there is no period in the filename it will default to dbo.

        If a schema does not currently exist, it will be created, after a prompt to confirm this. Authorization will be set to dbo by default.

        This behaviour will be overridden if the -Schema parameter is specified.

    .PARAMETER TableLock
        If this switch is enabled, the SqlBulkCopy option to acquire a table lock will be used.

        Per Microsoft "Obtain a bulk update lock for the duration of the bulk copy operation. When not
        specified, row locks are used."

    .PARAMETER CheckConstraints
        If this switch is enabled, the SqlBulkCopy option to check constraints will be used.

        Per Microsoft "Check constraints while data is being inserted. By default, constraints are not checked."

    .PARAMETER FireTriggers
        If this switch is enabled, the SqlBulkCopy option to allow insert triggers to be executed will be used.

        Per Microsoft "When specified, cause the server to fire the insert triggers for the rows being inserted into the database."

    .PARAMETER KeepIdentity
        If this switch is enabled, the SqlBulkCopy option to keep identity values from the source will be used.

        Per Microsoft "Preserve source identity values. When not specified, identity values are assigned by the destination."

    .PARAMETER KeepNulls
        If this switch is enabled, the SqlBulkCopy option to keep NULL values in the table will be used.

        Per Microsoft "Preserve null values in the destination table regardless of the settings for default values. When not specified, null values are replaced by default values where applicable."

    .PARAMETER NoProgress
        The progress bar is pretty but can slow down imports. Use this parameter to quietly import.

    .PARAMETER Quote
        Defines the default quote character wrapping every field.
        Default: double-quotes

    .PARAMETER Escape
        Defines the default escape character letting insert quotation characters inside a quoted field.

        The escape character can be the same as the quote character.
        Default: double-quotes

    .PARAMETER Comment
        Defines the default comment character indicating that a line is commented out.
        Default: hashtag

    .PARAMETER TrimmingOption
        Determines which values should be trimmed. Default is "None". Options are All, None, UnquotedOnly and QuotedOnly.

    .PARAMETER BufferSize
        Defines the default buffer size. The default BufferSize is 4096.

    .PARAMETER ParseErrorAction
        By default, the parse error action throws an exception and ends the import.

        You can also choose AdvanceToNextLine which basically ignores parse errors.

    .PARAMETER Encoding
        By default, set to UTF-8.

        The encoding of the file.

    .PARAMETER NullValue
        The value which denotes a DbNull-value.

    .PARAMETER MaxQuotedFieldLength
        The maximum length (in bytes) for any quoted field.

    .PARAMETER SkipEmptyLine
        Skip empty lines.

    .PARAMETER SupportsMultiline
        Indicates if the importer should support multiline fields.

    .PARAMETER UseColumnDefault
        Use the column default values if the field is not in the record.

    .PARAMETER NoTransaction
        Do not use a transaction when performing the import.

    .PARAMETER WhatIf
        Shows what would happen if the command were to run. No actions are actually performed.

    .PARAMETER Confirm
        Prompts you for confirmation before executing any changing operations within the command.

    .PARAMETER EnableException
        By default, when something goes wrong we try to catch it, interpret it and give you a friendly warning message.
        This avoids overwhelming you with "sea of red" exceptions, but is inconvenient because it basically disables advanced scripting.
        Using this switch turns this "nice by default" feature off and enables you to catch exceptions with your own try/catch.

    .NOTES
        Tags: Import, Data, Utility
        Author: Chrissy LeMaire (@cl), netnerds.net

        Website: https://dbatools.io
        Copyright: (c) 2018 by dbatools, licensed under MIT
        License: MIT https://opensource.org/licenses/MIT

    .LINK
        https://dbatools.io/Import-DbaCsv

    .EXAMPLE
        PS C:\> Import-DbaCsv -Path C:\temp\housing.csv -SqlInstance sql001 -Database markets

        Imports the entire comma-delimited housing.csv to the SQL "markets" database on a SQL Server named sql001, using the first row as column names.

        Since a table name was not specified, the table name is automatically determined from filename as "housing".

    .EXAMPLE
        PS C:\> Import-DbaCsv -Path .\housing.csv -SqlInstance sql001 -Database markets -Table housing -Delimiter "`t" -NoHeaderRow

        Imports the entire tab-delimited housing.csv, including the first row which is not used for colum names, to the SQL markets database, into the housing table, on a SQL Server named sql001.

    .EXAMPLE
        PS C:\> Import-DbaCsv -Path C:\temp\huge.txt -SqlInstance sqlcluster -Database locations -Table latitudes -Delimiter "|"

        Imports the entire pipe-delimited huge.txt to the locations database, into the latitudes table on a SQL Server named sqlcluster.

    .EXAMPLE
        PS C:\> Import-DbaCsv -Path c:\temp\SingleColumn.csv -SqlInstance sql001 -Database markets -Table TempTable -SingleColumn

        Imports the single column CSV into TempTable

    .EXAMPLE
        PS C:\> Get-ChildItem -Path \\FileServer\csvs | Import-DbaCsv -SqlInstance sql001, sql002 -Database tempdb -AutoCreateTable

        Imports every CSV in the \\FileServer\csvs path into both sql001 and sql002's tempdb database. Each CSV will be imported into an automatically determined table name.

    .EXAMPLE
        PS C:\> Get-ChildItem -Path \\FileServer\csvs | Import-DbaCsv -SqlInstance sql001, sql002 -Database tempdb -AutoCreateTable -WhatIf

        Shows what would happen if the command were to be executed

    .EXAMPLE
        PS C:\> Import-DbaCsv -Path c:\temp\dataset.csv -SqlInstance sql2016 -Database tempdb -Column Name, Address, Mobile

        Import only Name, Address and Mobile even if other columns exist. All other columns are ignored and therefore null or default values.

    .EXAMPLE
        PS C:\> Import-DbaCsv -Path C:\temp\schema.data.csv -SqlInstance sql2016 -database tempdb -UseFileNameForSchema

        Will import the contents of C:\temp\schema.data.csv to table 'data' in schema 'schema'.

    .EXAMPLE
        PS C:\> Import-DbaCsv -Path C:\temp\schema.data.csv -SqlInstance sql2016 -database tempdb -UseFileNameForSchema -Table testtable

        Will import the contents of C:\temp\schema.data.csv to table 'testtable' in schema 'schema'.

    .EXAMPLE
        PS C:\> $columns = @{
        >> Text = 'FirstName'
        >> Number = 'PhoneNumber'
        >> }
        PS C:\> Import-DbaCsv -Path c:\temp\supersmall.csv -SqlInstance sql2016 -Database tempdb -ColumnMap $columns

        The CSV field 'Text' is inserted into SQL column 'FirstName' and CSV field Number is inserted into the SQL Column 'PhoneNumber'. All other columns are ignored and therefore null or default values.

    .EXAMPLE
        PS C:\> $columns = @{
        >> 0 = 'FirstName'
        >> 1 = 'PhoneNumber'
        >> }
        PS C:\> Import-DbaCsv -Path c:\temp\supersmall.csv -SqlInstance sql2016 -Database tempdb -NoHeaderRow -ColumnMap $columns

        If the CSV has no headers, passing a ColumnMap works when you have as the key the ordinal of the column (0-based).
        In this example the first CSV field is inserted into SQL column 'FirstName' and the second CSV field is inserted into the SQL Column 'PhoneNumber'.
    #>
    [CmdletBinding(DefaultParameterSetName = 'TableOrView')]
    param (
        [Parameter(Mandatory)]
        $SqlInstance,

        [PSCredential]$SqlCredential,

        [string]$Database,

        [string]$Schema,

        [Parameter(ParameterSetName = 'TableOrView', Mandatory)]
        [Parameter(ValueFromPipeline)]
        [Alias('View')]
        $Table,

        [Parameter(ParameterSetName = 'TableOrView', Mandatory)]
        [ValidateScript({ Test-Path $_ -Type Container })]
        [string]$Path,

        [Parameter(ParameterSetName = 'OneQuery', Mandatory)]
        [int]$Query,

        [Parameter(ParameterSetName = 'TableOrView')]
        [Parameter(ParameterSetName = 'OneQuery', Mandatory)]
        [string]$FilePath,

        [string]$Delimiter = ',',

        [int]$BatchSize = 100000,

        [ValidateSet('ASCII', 'BigEndianUnicode', 'Byte', 'String', 'Unicode', 'UTF7', 'UTF8', 'Unknown')]
        [string]$Encoding = 'UTF8',

        [switch]$IncludeTypeInformation,


        [string]$DatabaseCulture = 'en-US', # https://www.fincher.org/Utilities/CountryLanguageList.shtml

        [ValidateScript({ $PSVersionTable.PSVersion.Major -gt 5 })]
        [switch]$NoHeader,

        [ValidateScript({ $PSVersionTable.PSVersion.Major -gt 5 })]
        [ValidateSet('Never', 'AsNeeded', 'Always')]  # If PS5, then it will be Always
        [string]$UseQuotes,

        [ValidateScript({ $PSVersionTable.PSVersion.Major -gt 5 })]
        [string[]]$QuoteFields,

        [int64]$MaxFileSize,

        [Parameter(ParameterSetName = 'TableOrView')]
        [int64]$Top,  # To limit SELECTed rows from a table or view

        [Parameter(ParameterSetName = 'TableOrView')]
        [string]$Nomenclature = '{Path}\{Database}-{Schema}-{Name}-{FileId}',

        [switch]$EnableException
    )
    begin {
        #region Initialization

        $fullnameFormat = $Nomenclature -replace 'Path', '0' -replace 'Database', '1' `
            -replace 'Schema', '2' -replace 'Name', '3' -replace 'FileId', '4:d3'

        Write-Message -Level Verbose -Message "Nomenclature = '$fullnameFormat' (should not have orginal text)"

        if ($DatabaseCulture) {
            [CultureInfo]::CurrentCulture = $DatabaseCulture
        }

        $timer = [Diagnostics.Stopwatch]::StartNew()

        #endregion

        #region Private functions

        function ConvertToTsqlColumnList {
            # Returns a list of columns for SELECT but CONVERT binary datatypes to VARCHAR
            # This is way faster than to convert locally
            param (
                $Columns
            )
            $colArray = New-Object System.Collections.ArrayList
            foreach ($col in $Columns) {
                if ($col.DataType.Name -like '*binary') {
                    #Write-Verbose "ICI $($col.DataType.Name)"
                    $null = $arr.Add($("convert(varchar(100),[{0}],1) as {0}" -f $col.Name))
                    $varcharMax = if ($col.DataType.MaximumLength -eq -1) { 'max' } else { $col.MaximumLength * 2 }
                    $null = $colArray.Add('convert(varchar({0}),[{1}],1) as [{1}]' -f $varcharMax, $col.Name)
                } else {
                    $null = $colArray.Add('[{0}]' -f $col.Name)
                }
            }
            return $colArray -join ','
        }
        function ConvertToHexadecimal {
            # If too diffcult to CONVERT BINARY to VARCHAR in SELECT. 0x is added because TSQL does the same. We might need to keep data integrity
            param (
                [byte[]]$Bytes
            )
            "0x$([Convert]::ToHexString($Bytes))"
        }
        function ExportQuery {
            # Run SqlReader then Export only one query result to FilePath with overwrite
            # Show a progress bar then return an object with RowCount per file
            param (
                [Parameter(Mandatory)]
                [Microsoft.Data.SqlClient.SqlConnection]$sqlConn,

                [Parameter(Mandatory)]
                [string]$Query,             # Only one SELECT query

                [Parameter(Mandatory)]
                $FilePath,                  # Fullname must be provided with existing directory

                $BatchSize,
                $Encoding,
                $Delimiter = ',',
                $UseQuote,                  # Available only on PS core 6+ (validated higher in the stack)
                $QuoteFields,               # Available only on PS core 6+ (validated higher in the stack)
                [int32]$CommandTimeout = 0, # Zero will wait forever (not sure this will be useful to change)
                [int64]$MaxFileSize = 0,    # Will split result into files if set (you can use 2GB and more)
                [int64]$TableRowCount = 0,  # Will provide a percent completed progress bar when provided
                [switch]$IncludeTypeInformation,
                [switch]$NoProgress,
                [switch]$NoHeader,
                [switch]$EnableException
            )
            $splatExportCsv = @{
                Delimiter         = $Delimiter
                NoTypeInformation = -not $IncludeTypeInformation
                Encoding          = $Encoding
                Append            = $true
            }
            if ($UseQuotes) { $splatExportCsv.UseQuote = $UseQuotes }
            if ($QuoteFields) { $splatExportCsv.QuoteFields = $QuoteFields }

            # Running query...
            try {
                $sqlCommand = $sqlConn.CreateCommand()
                $sqlCommand.CommandText = $Query
                $sqlCommand.CommandTimeout = $CommandTimeout
                $sqlDataReader = $sqlCommand.ExecuteReader()
                if ($sqlDataReader.HasRows) {

                    # Init data buffer
                    $schema = $SqlDataReader.GetSchemaTable()
                    $data = New-Object System.Data.DataTable

                    # Use schema to define data buffer
                    foreach ($col in $schema.Rows) {
                        $colName = $col.ColumnName
                        switch ($col.DataType) {
                            # binary types will be converted to string hexadecimal (takes longer than convert it in the query)
                            ([System.Byte[]]) { $t = [System.String]; break }
                            default { $t = $col.DataType }
                        }
                        [void]$data.Columns.Add($colName, $t)
                    }

                    [int64]$rowId = 0
                    [int32]$fileId = 1
                    [int32]$batchId = 0
                    [int32]$fileRowCount = 0

                    # Create first or only file
                    $csvPath = if ($MaxFileSize -eq 0) {
                        $FilePath -replace '(\.\w+$)', ('-{0:d3}$1' -f $fileId)
                    } else {
                        $FilePath
                    }
                    # Do not Test-Path if Force
                    if (-not $Force -and (Test-Path $csvFile)) {
                        Stop-Function -Message "File already exists $csvFile. Use -Force to overwrite."
                    }

                    # Foreach row
                    while ($SqlDataReader.Read()) {
                        $rowId++
                        $fileRowCount++

                        # Add the current row to the batch data table
                        $newRow = $data.Rows.Add();
                        foreach ($col in $data.Columns) {
                            $value = $SqlDataReader[$col.ColumnName]
                            if ($value -is [System.Byte[]]) {
                                $newRow[$col.ColumnName] = ConvertToHexadecimal $value
                            } else {
                                $newRow[$col.ColumnName] = $value
                            }
                        }

                        # When rowid reaches -BatchSize or -MaxFileSize then save batch to file
                        if (($rowid % $BatchSize -eq 0) -or $createNewFile) {
                            $batchId++

                            if ($createNewFile) {
                                # No need to check MaxFilesSize because it's not the first file
                                $csvPath = $FilePath -replace '(\.\w+$)', ('-{0:d3}$1' -f $fileId)
                                if (-not $Force -and (Test-Path $csvPath)) {
                                    Stop-Function -Message "File already exists $csvPath. Use -Force to overwrite."
                                }
                                $null = mkdir -Path (Split-Path $csvPath) -ErrorAction SilentlyContinue
                                $createNewFile = $false
                            }

                            $data | Export-Csv -Path $csvPath @splatExportCsv

                            if (-not $NoProgress) {
                                $splatProgress = @{
                                    Activity = 'Export-DbaCsv...'
                                    Status   = 'File {0}' -f $csvPath
                                }
                                # Percent completed is available only if rowcount is given
                                if ($TableRowCount -gt 0) {
                                    $splatProgress.PercentComplete = ($rowid * 100 / $TableRowCount)
                                }
                                Write-Progress @splatProgress
                            }

                            [int64]$csvSize = (Get-ChildItem $csvPath).Length
                            if (-not $avgRowSize) {
                                [int32]$avgRowSize = $csvSize / $rowid
                            }

                            # MaxFileSize minus estimated rowsize to avoid spilling
                            if ($createNewFile = $csvSize -gt ($MaxFileSize - $avgRowSize)) {
                                # Actual file will not receive any more updates
                                [PSCustomObject]@{
                                    FileName      = $csvFile
                                    FileRowCount  = $fileRowCount
                                    TotalRowCount = $rowid
                                    Length        = $csvSize
                                }
                                $fileRowCount = 0
                            }

                            # Then flush data buffer
                            $data.Clear()
                            $fileRowCount = 0
                        }
                    }
                }
            } catch {

            }
        }

        #endregion
    }
    process {
        foreach ($instance in $SqlInstance) {
            try {
                $server = Connect-DbaInstance -SqlInstance $instance -SqlCredential $SqlCredential -Database $Database -MinimumVersion 9
                $sqlConn = $server.ConnectionContext.SqlConnectionObject
                #if ($sqlConn.State -ne 'Open') { $sqlConn.Open() } # This is an abondance of precaution. Connect-DbaInstance opens the connection

                if ($PSCmdlet.ParameterSetName -eq 'TableOrView') {

                    # $tbl can be a table or a view. They all gets transformed to queries and sent to ExportQuery
                    foreach ($tbl in $Table) {
                        if ($obj = Get-DbaDbTable -SqlInstance $SqlInstance -Database $Database -Schema $Schema -Table $tbl -Verbose:$false) {
                            $commandText = ("SELECT $(if($Top){"TOP $Top "}; ConvertToTsqlColumnList -Columns $obj.Columns)
                            FROM [{0}].[{1}].[{2}];" -f $obj.Database, $obj.Schema, $obj.Name) -replace ' {2,}', ' '
                            $tableRowCount = [math]::Min($tbl.RowCount, $Top)
                        } elseif ($obj = Get-DbaDbView -SqlInstance $SqlInstance -Database $tbl.Database -Schema $tbl.Schema -View $tbl.Name -Verbose:$false) {
                            $commandText = "SELECT $(if($Top){"TOP $Top "})`nFROM [{0}].[{1}].[{2}];" -f $obj.Database, $obj.Schema, $obj.Name
                        } else {
                            Write-Message -Level Warning -Message "$tbl was not found in $SqlInstance $Database $Schema" -Continue
                        }

                        # $obj was found
                        # $fullname = '[{0}].[{1}].[{2}]' -f $obj.Database, $obj.Schema, $obj.Name

                        # Running query...
                        $splatExportQuery = @{
                            sqlConn = $sqlConn
                            Query   = $commandText
                        }
                        if ($tableRowCount) { $splatExportQuery.Add('TableRowCount', $tableRowCount) }

                        $PSBoundParameters.GetEnumerator() | Where-Object Key -in ('Delimiter,BatchSize,Encoding,IncludeTypeInformation,NoHeader,
                        UseQuote,QuoteFields,CommandTimeout,MaxFileSize' -replace '\s+' -split ',') | ForEach-Object {
                            $splatExportQuery.Add($_.Key, $_.Value)
                        }

                        ExportQuery @splatExportQuery

                    }
                } elseif ($PSCmdlet.ParameterSetName -eq 'OneQuery') {
                    # Running query...
                    $splatExportQuery = @{
                        sqlConn  = $sqlConn
                        Query    = $commandText
                        FilePath = $FilePath
                    }
                    $PSBoundParameters.GetEnumerator() | Where-Object Key -in ('Delimiter,BatchSize,Encoding,IncludeTypeInformation,NoHeader,
                    UseQuote,QuoteFields,CommandTimeout,MaxFileSize,NoHeader' -replace '\s+' -split ',') | ForEach-Object {
                        $splatExportQuery.Add($_.Key, $_.Value)
                    }

                    ExportQuery @splatExportQuery

                    [PSCustomObject]@{
                        FullName    = '{0}.{1}.{2}' -f $obj.Database, $obj.Schema, $obj.Name
                        CommandText = "SELECT $(if($Top){"TOP $Top "})`nFROM [{0}].[{1}].[{2}];" -f $obj.Database, $obj.Schema, $obj.Name
                        Object      = $obj
                    }
                }
            } catch {
                throw $_
                #Stop-Function -Message "Failure" -Category ConnectionError -ErrorRecord $_ -Target $instance -Continue
            }
        }
    }
    end {
        # Close everything just in case & ignore errors
        $ErrorActionPreference = 'SilentlyContinue'
        $null = $SqlDataReader.Close()
        $null = $sqlDataReader.Dispose()
        $null = $sqlConn.Close()
        $null = $sqlConn.Dispose()

        # Script is finished. Show elapsed time.
        $timer.Stop()
        Write-Message -Level Verbose -Message "Total Elapsed Time $($timer.Elapsed.ToString())"
    }
}
<#
        } else {
        }
        foreach ($qry in $queries) {
            # $qry =@{Name, CommandText, Object}


                # Prepare Export-Csv parameters (for all rows)
                $splatExportCsv = @{
                    Delimiter         = $Delimiter
                    NoTypeInformation = -not $IncludeTypeInformation
                    Encoding          = $Encoding
                    Append            = $true
                }
                if ($PSVersionTable.PSVersion.Major -gt 5) {
                    if ($UseQuotes) { $splatExportCsv.UseQuote = $UseQuotes }
                    if ($QuoteFields) { $splatExportCsv.QuoteFields = $QuoteFields }
                }

                # RBAR Transfer process
                while ($SqlDataReader.Read()) {
                    $rowId++
                    $newRow = $data.Rows.Add()

                    # Copy row value in data buffer foreach column
                    foreach ($col in $data.Columns) {
                        # Copy value in the right column
                        $value = $SqlDataReader[$col.ColumnName]
                        if ($value -is [Byte[]]) {
                            # Binary is transformed to hexadecimal
                            $newRow[$col.ColumnName] = ConvertToHexadecimal $value
                        } else {
                            $newRow[$col.ColumnName] = $value
                        }
                    }

                    # Start a new batch or a new csv file
                    if (($rowId % $BatchSize -eq 0) -or $isNewFile) {
                        $batch++

                        # Flush data buffer to csv file
                        Write-Verbose ($data.Columns -join ',')
                        $data | Export-Csv -Path $csvPath @splatExportCsv

                        # Stats about this batch
                        Write-Message -Level Verbose -Message ('{0} rows in {1}' -f $rowId, $csvPath)

                        # Cleanup
                        $data.Clear()
                        $isNewFile = $false
                    }

                    # Start a new file when file size reach MaxFileSize
                    $csvSize = (Get-ChildItem $FilePath).Length
                    if ($isNewFile = $csvSize -ge $MaxFileSize) {
                        $csvPath = NewCsvFile -Path $Path -FileId (++$fid) -Obj $qry.Object
                    }

                    $data.Clear()
                }
                # Flush pending rows
                if ($data.Rows.Count -gt 0) {
                    Write-Message -Level Verbose -Message ('Flushing pending {0} rows' -f $data.Rows.Count)
                }
            }
        }
    }
    end {
        # Close everything just in case & ignore errors
        $ErrorActionPreference = 'SilentlyContinue'
        $null = $SqlDataReader.Close()
        $null = $sqlDataReader.Dispose()
        $null = $sqlConn.Close()
        $null = $sqlConn.Dispose()

        # Script is finished. Show elapsed time.
        $timer.Stop()
        Write-Message -Level Verbose -Message "Total Elapsed Time $($timer.Elapsed.ToString())"
    }
}
#>
