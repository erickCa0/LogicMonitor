<#
.SYNOPSIS
Exports all LogicMonitor web checks from a customer portal to CSV.

.DESCRIPTION
Prompts for the LogicMonitor portal/account name, LM API access ID, and LM API
access key. The script authenticates with the LogicMonitor PowerShell module,
runs Get-LMWebsite for all web checks, and writes the returned objects to CSV.

.PARAMETER PortalName
LogicMonitor portal/account name. You can enter either the short customer name
such as "customer" or a portal host such as "customer.logicmonitor.com".

.PARAMETER OutputPath
CSV output path. Defaults to a timestamped file in the current directory.

.PARAMETER ModuleName
LogicMonitor PowerShell module to import/install. Defaults to Logic.Monitor.

.PARAMETER InstallModuleIfMissing
Installs the LogicMonitor PowerShell module from PowerShell Gallery if it is not
already available on the workstation.

.EXAMPLE
.\Export-LMWebChecks.ps1

.EXAMPLE
.\Export-LMWebChecks.ps1 -PortalName customer -OutputPath C:\Temp\customer-web-checks.csv
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$PortalName,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = (Join-Path -Path (Get-Location) -ChildPath ("LMWebChecks_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))),

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ModuleName = 'Logic.Monitor',

    [Parameter()]
    [switch]$InstallModuleIfMissing
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function ConvertTo-LMAccountName {
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Portal
    )

    $normalizedPortal = $Portal.Trim()
    $normalizedPortal = $normalizedPortal -replace '^https?://', ''
    $normalizedPortal = $normalizedPortal -replace '/.*$', ''
    $normalizedPortal = $normalizedPortal -replace '\.logicmonitor\.com$', ''

    if ([string]::IsNullOrWhiteSpace($normalizedPortal)) {
        throw 'Portal name cannot be empty.'
    }

    return $normalizedPortal
}

function Ensure-Module {
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [Parameter()]
        [switch]$InstallIfMissing
    )

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        if (-not $InstallIfMissing) {
            throw "PowerShell module '$Name' is not installed. Install it first or rerun this script with -InstallModuleIfMissing."
        }

        Write-Host "Installing module '$Name' from PowerShell Gallery..." -ForegroundColor Cyan
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module -Name $Name -ErrorAction Stop
}

function Connect-LogicMonitorPortal {
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$AccountName,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$AccessId,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [securestring]$AccessKey
    )

    $connectCommand = Get-Command -Name Connect-LMAccount -ErrorAction SilentlyContinue
    if (-not $connectCommand) {
        throw "The expected authentication cmdlet 'Connect-LMAccount' was not found after importing the module."
    }

    $plainAccessKey = [System.Net.NetworkCredential]::new('', $AccessKey).Password

    try {
        $parameters = @{}

        if ($connectCommand.Parameters.ContainsKey('AccountName')) {
            $parameters['AccountName'] = $AccountName
        }
        elseif ($connectCommand.Parameters.ContainsKey('Company')) {
            $parameters['Company'] = $AccountName
        }
        elseif ($connectCommand.Parameters.ContainsKey('Portal')) {
            $parameters['Portal'] = $AccountName
        }
        else {
            throw "Connect-LMAccount does not expose an AccountName, Company, or Portal parameter."
        }

        if ($connectCommand.Parameters.ContainsKey('AccessId')) {
            $parameters['AccessId'] = $AccessId
        }
        elseif ($connectCommand.Parameters.ContainsKey('AccessID')) {
            $parameters['AccessID'] = $AccessId
        }
        else {
            throw "Connect-LMAccount does not expose an AccessId or AccessID parameter."
        }

        $accessKeyParameterName = $null
        if ($connectCommand.Parameters.ContainsKey('AccessKey')) {
            $accessKeyParameterName = 'AccessKey'
        }
        elseif ($connectCommand.Parameters.ContainsKey('AccessKEY')) {
            $accessKeyParameterName = 'AccessKEY'
        }
        else {
            throw "Connect-LMAccount does not expose an AccessKey or AccessKEY parameter."
        }

        if ($connectCommand.Parameters[$accessKeyParameterName].ParameterType -eq [securestring]) {
            $parameters[$accessKeyParameterName] = $AccessKey
        }
        else {
            $parameters[$accessKeyParameterName] = $plainAccessKey
        }

        Connect-LMAccount @parameters | Out-Null
    }
    finally {
        $plainAccessKey = $null
    }
}

function Get-AllLMWebsites {
    $getWebsiteCommand = Get-Command -Name Get-LMWebsite -ErrorAction SilentlyContinue
    if (-not $getWebsiteCommand) {
        throw "The expected cmdlet 'Get-LMWebsite' was not found after importing the module."
    }

    $parameters = @{}
    if ($getWebsiteCommand.Parameters.ContainsKey('All')) {
        $parameters['All'] = $true
    }

    return @(Get-LMWebsite @parameters)
}

try {
    Ensure-Module -Name $ModuleName -InstallIfMissing:$InstallModuleIfMissing

    if (-not $PortalName) {
        $PortalName = Read-Host 'Enter the LogicMonitor portal name'
    }

    $accountName = ConvertTo-LMAccountName -Portal $PortalName
    $accessId = Read-Host 'Enter the LM API access ID'
    $accessKey = Read-Host 'Enter the LM API access key' -AsSecureString

    Write-Host "Authenticating to LogicMonitor portal '$accountName'..." -ForegroundColor Cyan
    Connect-LogicMonitorPortal -AccountName $accountName -AccessId $accessId -AccessKey $accessKey

    Write-Host 'Retrieving web checks with Get-LMWebsite...' -ForegroundColor Cyan
    $webChecks = Get-AllLMWebsites

    $outputDirectory = Split-Path -Path $OutputPath -Parent
    if ($outputDirectory -and -not (Test-Path -Path $outputDirectory)) {
        New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
    }

    $webChecks | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

    Write-Host ("Exported {0} web check(s) to {1}" -f $webChecks.Count, (Resolve-Path -Path $OutputPath).Path) -ForegroundColor Green
}
catch {
    Write-Error $_
    exit 1
}
