[CmdletBinding()] Param(
    [Parameter()]
    [string]$AzureEnvironment,
    [Parameter()]
    [string]$ExchangeEnvironment,
    [Parameter()]
    [datetime]$StartDate = [DateTime]::UtcNow.AddDays(-90),
    [Parameter()]
    [datetime]$EndDate = [datetime]::UtcNow,
    [Parameter()]
    [string]$ExportDir = (Join-Path ([System.Environment]::GetFolderPath("Desktop")) 'ExportDir'),
    [Parameter()]
    [switch]$No365 = $false)

    function Import-PSModules {
            [CmdletBinding()] Param(
            [Parameter(Mandatory = $true)]
            [string]$ExportDir)

            $ModuleArray = @("ExchangeOnlineManagement","AzureAD", "MSOnline")

            foreach($ReqModule in $ModuleArray) {
                If ($null -eq (Get-Module $ReqModule -ListAvailable -ErrorAction SilentlyContinue)) {
                    Write-Verbose "Required Module, $ReqModule, is not installed on the system!"
                    Write-Verbose "Installing $ReqModule from default repository!"
                    Install-Module -Name $ReqModule -Force
                    Write-Verbose "Importing $ReqModule"
                    Import-Module -Name $ReqModule
                } elseif ($null -eq (Get-Module $ReqModule -ErrorAction SilentlyContinue)){
                    Write-Verbose "Importing $ReqModule"
                    Import-Module -Name $ReqModule
                }
            }

            if (!(Test-Path $ExportDir)) {
                New-Item -Path $ExportDir -ItemType "Directory" -Force
            }
    }

    function  Get-AzureEnvironments() {
            [CmdletBinding()] Param(
            [Parameter()]
            [string]$AzureEnvironment,
            [Parameter()]
            [string]$ExchangeEnvironment)

            $AzureEnvironments = [Microsoft.Open.Azure.AD.CommonLibrary.AzureEnvironment]::PublicEnvironments.Keys
            while ($AzureEnvironments -cnotcontains $AzureEnvironment -or [string]::IsNullOrWhiteSpace($AzureEnvironment)) {
                Write-Host 'Azure Environments'
                Write-Host '-------------------'
                $AzureEnvironments | ForEach-Object {Write-Host $_}
                $AzureEnvironment = Read-Host 'Choose your Azure Environment [AzureCloud]'
                if ([string]::IsNullOrWhiteSpace($AzureEnvironment)) {$AzureEnvironment = 'AzureCloud'}
            }
            
            if ($No365 -eq $false) {
                $ExchangeEnvironments = [System.Enum]::GetNames([Microsoft.Exchange.Management.RestAPIClient.ExchangeEnvironment])
                while ($ExchangeEnvironments -cnotcontains $ExchangeEnvironment -or [string]::IsNullOrWhiteSpace($ExchangeEnvironment) -and $ExchangeEnvironment -ne "None") {
                    Write-Host 'Exchange Environments'
                    Write-Host '---------------------'
                    $ExchangeEnvironments | ForEach-Object {Write-Host $_}
                    Write-Host 'None'
                    $ExchangeEnvironment = Read-Host 'Choose your Exchange Environment [O365 Default]'
                    if ([string]::IsNullOrWhiteSpace($ExchangeEnvironment)) {$ExchangeEnvironment = 'O365Default'}
                }
            } else {
                $ExchangeEnvironment = "None"
            }
            Return ($AzureEnvironment, $ExchangeEnvironment)
    }

    function New-ExcelFromCSV() {
        [CmdletBinding()]Param(
            [string]$ExportDir
        )

        try {
            $Excel = New-Object -ComObject Excel.Application
        }
        catch {
            Write-Host 'Warning: Excel Not Found! - Skipping Combined File!'
            return 
        }

        $Excel.DisplayAlerts = $false
        $Workbook = $Excel.Workbooks.Add()
        $CSVs = Get-ChildItem -Path "${ExportDir}\*.csv" -Force
        $ToDeletes = $Workbook.Sheets | ForEach-Object -ExpandProperty Name
        foreach ($CSV in $CSVs) {
            $TempWorkbook = $Excel.Worksbooks.Open($CSV.FullName)
            $TempWorkbook.Sheets[1].Copy($Workbook.Sheets[1]. [Type]::Missing) | Out-Null
            $Workbook.Sheets[1].UsedRange.Columns.AutoFit() | Out-Null
            $Workbook.Sheets[1].Name = $CSV.BaseName -replace '_Operations_.*',''
        }

        foreach($ToDelete In $ToDeletes) {
            $Workbook.Activate()
            $Workbook.Sheets[$ToDelete].Activate()
            $Workbook.Sheets[$ToDelete].Delete()
        }

        $Workbook.SaveAs((Join-Path $ExportDir 'Summary_Export.xlsx'))
        $Excel.Quit()
    }

    function Get-UALData {
        [CmdletBinding()] param (
            [Parameter(Mandatory = $true)]
            [datetime] $StartDate,
            [Parameter(Mandatory = $true)]
            [datetime] $EndDate,
            [Parameter(Mandatory = $true)]
            [string] $AzureEnvironment,
            [Paramter(Mandatory = $true)]
            [string] $ExchangeEnvironment,
            [Parameter(Mandatory = $true)]
            [string] $ExportDir
        )

        Connect-ExchangeOnline -ExchangeEnvironment $ExchangeEnvironment
        $LicenseQuestion = Read-Host 'Do you have an Office365/Microsoft 365 E5/E6 License? Y/N'
        switch ($LicenseQuestion) {
            Y {$LicenseAnswer = "Yes"}
            N {$LicenseAnswer = "No"}
        }
        $AppIDQuestion = Read-Host 'Would you like to investigate a certain application? Y/N'
        switch ($AppIDQuestion) {
            Y {$AppIDInvestigation = "Yes"}
            N {$AppIDInvestigation = "No"}
        }
        If ($AppIDInvestigation -eq "Yes") {
            $SusAppID = Read-Host "Enter the application AppID to investigate!"
        } else {
            Write-Host "Skipping AppID Investigation!"
        }

        Write-Verbose "Searching for 'Set Domain Authentication' and 'Set federation settings on domian' operations in the UAL!"
        $DomainData = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -RecordType AzureActiveDirectory -Operations "Set Domain Authentication", "Set federation settings on domain!" -ResultSize 5000 | Select-Object -ExpandProperty AuditData | ConvertFrom-Json
        Export-UALData -ExportDir $ExportDir -UALInput $DomainData -CSVName "Domain_Operations_Export" -WorkloadType "AAD"
        Write-Verbose "Searching for 'Update Application' and 'Update Allication | Certificate and Secrets Management' in the UAL!"
        $AppData = Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -RecordType AzureActiveDirectory -Operations "Update Application", "Update Application | Certificates and Secret Management" -ResultSize 5000 | Select-Object -ExpandProperty AuditDate | ConvertFrom-Json

    }

