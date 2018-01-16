﻿function Remove-DbaPfDataCollectorCounter {
    <#
        .SYNOPSIS
            Removes a Performance Data Collector Counter right from the server itself, no templates required.

        .DESCRIPTION
            Removes a Performance Data Collector Counter right from the server itself, no templates required.
    
            Copies line for line from a source server. For more configurable options, use Remove-DbaPfDataCollectorCounterTemplate.

        .PARAMETER ComputerName
            The target computer. Defaults to localhost.

        .PARAMETER Credential
            Allows you to login to $ComputerName using alternative credentials.
    
        .PARAMETER CollectorSet
            The Collector Set name

        .PARAMETER Collector
            The Collector name
    
        .PARAMETER Counter
            The Counter name - in the form of '\Processor(_Total)\% Processor Time'. This field is required.
    
        .PARAMETER InputObject
            Enables piped results from Get-DbaPfDataCollector. This field is required
    
        .PARAMETER WhatIf
            Shows what would happen if the command were to run. No actions are actually performed.

            .PARAMETER Confirm
            Prompts you for confirmation before executing any changing operations within the command.
        
        .PARAMETER EnableException
            By default, when something goes wrong we try to catch it, interpret it and give you a friendly warning message.
            This avoids overwhelming you with "sea of red" exceptions, but is inconvenient because it basically disables advanced scripting.
            Using this switch turns this "nice by default" feature off and enables you to catch exceptions with your own try/catch.
    
        .NOTES
            Tags: PerfMon
            Website: https://dbatools.io
            Copyright: (C) Chrissy LeMaire, clemaire@gmail.com
            License: GNU GPL v3 https://opensource.org/licenses/GPL-3.0
    
        .LINK
            https://dbatools.io/Remove-DbaPfDataCollectorCounter

        .EXAMPLE
            Remove-DbaPfDataCollectorCounter -ComputerName sql2017 -CollectorSet 'System Correlation' -Collector DataCollector01  -Counter '\LogicalDisk(*)\Avg. Disk Queue Length'
    
            Removes the '\LogicalDisk(*)\Avg. Disk Queue Length' counter within the datacollector1 collector within the system correlation collector set on sql2017
    
        .EXAMPLE
            Get-DbaPfDataCollectorCounter | Out-GridView -PassThru | Remove-DbaPfDataCollectorCounter -Confirm:$false
    
            Allows you to select which counters you'd like on localhost and does not prompt for confirmation

    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = "High")]
    param (
        [DbaInstance[]]$ComputerName=$env:COMPUTERNAME,
        [PSCredential]$Credential,
        [Alias("DataCollectorSet")]
        [string[]]$CollectorSet,
        [Alias("DataCollector")]
        [string[]]$Collector,
        [Alias("Name")]
        [parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [object[]]$Counter,
        [parameter(ValueFromPipeline)]
        [object[]]$InputObject,
        [switch]$EnableException
    )
    begin {
        $setscript = {
            $setname = $args[0]; $removexml = $args[1]
            $CollectorSet = New-Object -ComObject Pla.DataCollectorSet
            $CollectorSet.SetXml($removexml)
            $CollectorSet.Commit($setname, $null, 0x0003) #add or modify.
            $CollectorSet.Query($setname, $Null)
        }
    }
    process {
        if ($InputObject.Credential -and (Test-Bound -ParameterName Credential -Not)) {
            $Credential = $InputObject.Credential
        }
        
        if (-not $InputObject -or ($InputObject -and (Test-Bound -ParameterName ComputerName))) {
            foreach ($computer in $ComputerName) {
                $InputObject += Get-DbaPfDataCollectorCounter -ComputerName $computer -Credential $Credential -CollectorSet $CollectorSet -Collector $Collector -Counter $Counter
            }
        }
        
        if ($InputObject) {
            if (-not $InputObject.CounterObject) {
                Stop-Function -Message "InputObject is not of the right type. Please use Get-DbaPfDataCollectorCounter"
                return
            }
        }
        
        foreach ($object in $InputObject) {
            $computer = $InputObject.ComputerName
            $null = Test-ElevationRequirement -ComputerName $computer -Continue
            $setname = $InputObject.DataCollectorSet
            $collectorname = $InputObject.DataCollector
            
            $xml = [xml]($InputObject.DataCollectorSetXml)
            
            <#
            $newItemtoAdd = $xml.CreateElement('newItemtoAdd')
            $newItemtoAdd.PsBase.InnerText = '1900-01-01'
            $xml.Entity.App.AppendChild($newItemtoAdd) | Out-Null
            #>
            
            foreach ($countername in $counter) {
                $node = $xml.SelectSingleNode("//Name[.='$collectorname']").SelectSingleNode("//Counter[.='$countername']")
                $null = $node.ParentNode.RemoveChild($node)
                $node = $xml.SelectSingleNode("//Name[.='$collectorname']").SelectSingleNode("//CounterDisplayName[.='$countername']")
                $null = $node.ParentNode.RemoveChild($node)
            }
            
            $plainxml = $xml.OuterXml
            
            if ($Pscmdlet.ShouldProcess("$computer", "Remove $countername from $collectorname with the $setname collection set")) {
                try {
                    $results = Invoke-Command2 -ComputerName $computer -Credential $Credential -ScriptBlock $setscript -ArgumentList $setname, $plainxml -ErrorAction Stop
                    Write-Message -Level Verbose -Message " $results"
                    $null = [pscustomobject]@{
                        ComputerName                                         = $computer
                        DataCollectorSet                                     = $setname
                        DataCollector                                        = $collectorname
                        DataCollectorCounter                                 = $counterName
                        Status                                               = "Removed"
                    }
                }
                catch {
                    Stop-Function -Message "Failure importing $Countername to $computer" -ErrorRecord $_ -Target $computer -Continue
                }
            }
        }
    }
}