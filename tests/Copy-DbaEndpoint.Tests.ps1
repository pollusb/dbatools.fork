$CommandName = $MyInvocation.MyCommand.Name.Replace(".Tests.ps1", "")
Write-Host -Object "Running $PSCommandPath" -ForegroundColor Cyan
$global:TestConfig = Get-TestConfig

Describe "$CommandName Unit Tests" -Tag 'UnitTests' {
    Context "Validate parameters" {
        [object[]]$params = (Get-Command $CommandName).Parameters.Keys | Where-Object {$_ -notin ('whatif', 'confirm')}
        [object[]]$knownParameters = 'Source', 'SourceSqlCredential', 'Destination', 'DestinationSqlCredential', 'Endpoint', 'ExcludeEndpoint', 'Force', 'EnableException'
        $knownParameters += [System.Management.Automation.PSCmdlet]::CommonParameters
        It "Should only contain our specific parameters" {
            (@(Compare-Object -ReferenceObject ($knownParameters | Where-Object {$_}) -DifferenceObject $params).Count ) | Should Be 0
        }
    }
}
<#
    Integration test should appear below and are custom to the command you are writing.
    Read https://github.com/dataplat/dbatools/blob/development/contributing.md#tests
    for more guidence.
#>

Describe "$commandname Integration Tests" -Tag "IntegrationTests" {
    BeforeAll {
        Get-DbaEndpoint -SqlInstance $TestConfig.instance2 -Type DatabaseMirroring | Remove-DbaEndpoint -Confirm:$false
        New-DbaEndpoint -SqlInstance $TestConfig.instance2 -Name dbatoolsci_MirroringEndpoint -Type DatabaseMirroring -Port 5022 -Owner sa
        Get-DbaEndpoint -SqlInstance $TestConfig.instance3 -Type DatabaseMirroring | Remove-DbaEndpoint -Confirm:$false
    }
    AfterAll {
        Get-DbaEndpoint -SqlInstance $TestConfig.instance2 -Type DatabaseMirroring | Remove-DbaEndpoint -Confirm:$false
        New-DbaEndpoint -SqlInstance $TestConfig.instance2 -Name dbatoolsci_MirroringEndpoint -Type DatabaseMirroring -Port 5022 -Owner sa
        Get-DbaEndpoint -SqlInstance $TestConfig.instance3 -Type DatabaseMirroring | Remove-DbaEndpoint -Confirm:$false
        New-DbaEndpoint -SqlInstance $TestConfig.instance3 -Name dbatoolsci_MirroringEndpoint -Type DatabaseMirroring -Port 5023 -Owner sa
    }

    It "copies an endpoint" {
        $results = Copy-DbaEndpoint -Source $TestConfig.instance2 -Destination $TestConfig.instance3 -Endpoint dbatoolsci_MirroringEndpoint
        $results.DestinationServer | Should -Be  $TestConfig.instance3
        $results.Status | Should -Be 'Successful'
        $results.Name | Should -Be 'dbatoolsci_MirroringEndpoint'
    }
}
