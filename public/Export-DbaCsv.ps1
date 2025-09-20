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
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [DbaInstanceParameter[]]$SqlInstance,   # il faut prévoir comment connecter à azure ?
        [PSCredential]$SqlCredential,

        [Parameter(Mandatory)]
        [object[]]$Path,
        [string]$Nomenclature = '{Path}\{Database}\{Database}_{Schema}_{Name}_{FileId}',

        [string]$Database,
        [string]$Schema,

        [Parameter(ValueFromPipeline)]
        [string]$Table,  # will accept three parts table or view name ex: Database.Schema.Name

        [char]$Delimiter = ',',
        #[ValidateSet('ASCII', 'BigEndianUnicode', 'Byte', 'String', 'Unicode', 'UTF7', 'UTF8', 'Unknown')]
        [string]$Encoding = 'utf8',
        [switch]$IncludeTypeInformation,
        [switch]$NoHeader,
        [string]$DatabaseCulture = "en-US",

        [ValidateSet('Never', 'AsNeeded', 'Always')]
        [string]$UseQuotes,

        [string[]]$QuoteFields,

        [int]$BatchSize = 50000,
        [int64]$MaxFileSize = 2GB,  # To generate multiple csv files
        [switch]$Force,             # Overwrite files
        [switch]$EnableException
    )
    begin {
        if ($PSVersionTable.PSVersion.Major -le 5 -and ($UseQuotes -or $QuoteFields)) {
            Write-Message -Level Warning -Message 'Parameter -UseQuotes or -QuoteFields works only with PSv6+'
        }

        #region Private functions

        function New-CsvFile {
            param (
                [int]$FileId,
                $tblObj
            )
            $format = $Nomenclature -replace 'Path', '0' -replace 'Database', '1' `
                -replace '$Schema', '2' -replace 'Name', '3' -replace 'FileId', '4'

            if (-not $script:once) { Write-Verbose "Nomenclature = '$format'"; $once = $true }
            return ($format -f $Path, $tbl.Database, $tbl.Schema, $tbl.Name, $FileId)
        }
        function ConvertTo-TsqlColumnList {
            # Returns a list of columns for SELECT but CONVERT binary datatypes to VARCHAR
            # This is way faster than to convert locally
            param (
                [Microsoft.SqlServer.Management.Smo.Column]$Columns
            )
            $arr = New-Object System.Collections.ArrayList
            foreach ($col in $Columns) {
                if ($col.DataType.Name -like '*binary') {
                    $null = $arr.Add('convert(varchar({0}}),[{0}],1) as [{1}]' -f $col.Name)
                }
                else {
                    $null = $arr.Add('[{0}]' -f $col.Name)
                }
            }
            return $arr -join ','
        }
        function ConvertTo-Hexadecimal {
            # If too complicated to CONVERT BINARY to VARCHAR
            param (
                [byte[]] $BytesArray
            )
            ($BytesArray | ForEach-Object { $_.ToString('X2') }) -join ''
        }
        #endregion

        # Connecting...
        $SqlConnectionString = 'Data Source={0};Initial Catalog={1};Integrated Security=SSPI;Connection Timeout={2}' -f $SqlInstance, $Database, $ConnectionTimeout
        $SqlConnection = New-Object -TypeName System.Data.SqlClient.SqlConnection -ArgumentList $SqlConnectionString
        $SqlConnection.Open()

        Write-Message -Level Verbose -Message "Started at $(Get-Date)"
    }
    process {
        if ($Query) {
            $SqlQuery = $Query
        } else {
            foreach ($tbl in $Table) {
                #$Database = $tbl.Database
                #$Schema = $tbl.Schema
                #$Name = $tbl.Name
                $SqlQuery = "SELECT $(
                if($Top){"TOP $Top "}; ObtenirCols -columns $tbl.Columns
                )`nFROM {0}.{1}.{2};" -f $tbl.Database, $tbl.Schema, $tbl.Name
            }
        }
        # Running query...
        $SqlCommand = $SqlConnection.CreateCommand()
        $SqlCommand.CommandText = $SqlQuery
        $SqlCommand.CommandTimeout = 0
        $SqlDataReader = $SqlCommand.ExecuteReader()

        if ($SqlDataReader.HasRows) {

            # Get the table schema from the data reader
            $schema = $SqlDataReader.GetSchemaTable()

            # Define the data table that will contain the batch rows
            $data = New-Object System.Data.DataTable

            # Define columns in the data buffer
            foreach ($colDef in $schema.Rows) {
                $colName = $colDef.ColumnName
                switch ($colDef.DataType) {
                    # Binary needs to be transformed in hexadecimal
                    ([System.Byte[]]) { $colType = [System.String]; break }
                    default { $colType = $colDef.DataType }
                }
                [void]$data.Columns.Add($colName, $colType)
            }
            [int64]$rows = 0
            [int]$batches = 0
            [int]$rid = 1
            [int]$fid = 1

            $filePath = New-CsvFile -FileId $fid -tblObj $tbl

        }











    }
    end {
        # Close everything just in case & ignore errors

        $null = $SqlConnection.Close()
        $null = $SqlConnection.Dispose()

        # Script is finished. Show elapsed time.
        $totaltime = [math]::Round($scriptelapsed.Elapsed.TotalSeconds, 2)
        Write-Message -Level Verbose -Message "Total Elapsed Time for everything: $totaltime seconds"
    }
}