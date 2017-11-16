<#
TPAR Prerequisite Installer v2.0
Joe Beaulieu
ERP Technology Integration
Nov 2017

Contributions:  https://stackoverflow.com/questions/31712686/how-to-check-if-a-program-is-installed-and-install-it-if-it-is-not
                    
#>

####################
# GLOBAL VARIABLES #
####################

$tempdir = (Get-Location).toString()
$preReqFeatures = Get-Content $tempdir\WindowsFeatures.txt
$SQLSysClrTypes = '*Microsoft System CLR Types for SQL Server 2012 (x64)*'
$ReportViewer2012 = '*Microsoft Report Viewer 2012 Runtime*'
$sqlCLRmsi = $tempdir+'\SQLSysClrTypes.msi'
$RVmsi = $tempdir+'\ReportViewer.msi'
$msiArgs = '/qb'
$rebootNeeded = $false



#############
# FUNCTIONS #
#############


function Install-PrereqWindowsFeatures 
{
    param(
    [Parameter (Mandatory=$true)]
    [String[]] $features
    )

    foreach ($feature in $features) 
    {
        $wf = Get-WindowsFeature -Name $feature
        $displayName = $wf.DisplayName
        if (!$wf.Installed) 
        {
            if (!$wf) 
            {
                Write-Host
                Write-Host "Feature $feature not found." -ForegroundColor Yellow
            }
            else 
            {
                Write-Host
                Write-Host "Installing $displayName" -ForegroundColor White -BackgroundColor DarkGreen
                Install-WindowsFeature $feature
            } #end inner if-else
        }
        else{
            Write-Host
            Write-Host "The feature $displayName is already installed." -ForegroundColor DarkYellow
        } #end outer if-else
    } #end foreach
} #end function Install-PrereqWindowsFeatures



function Get-InstalledApps
{
    if ([IntPtr]::Size -eq 4) {
        $regpath = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    }
    else {
        $regpath = @(
            'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
            'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
        )
    } #end if-else

    Get-ItemProperty $regpath | .{process{if($_.DisplayName -and $_.UninstallString) { $_ } }} | Select DisplayName, Publisher, InstallDate, DisplayVersion, UninstallString |Sort DisplayName
} #end function Get-InstalledApps


function Install-PrereqMSIs
{
    $preReqSoftware = @(
        $SQLSysClrTypes
        $ReportViewer2012
    )
    
    foreach ($preReq in $preReqSoftware)
    {
        if ($preReq -like '*SQL*')
        {
            $msiFile = $sqlCLRmsi
            $displayName = 'Microsoft System CLR Types for SQL Server 2012 (x64)' 
        }
        else
        {
            $msiFile = $RVmsi
            $displayName = 'Microsoft Report Viewer 2012 Runtime'
        } #end if-else
        

        $result = Get-InstalledApps | where {$_.DisplayName -like $preReq}
        if ($result -eq $null) 
        {
            $exitCode = (Start-Process -FilePath $msiFile -ArgumentList $msiArgs -Wait -Passthru).ExitCode
            if ($exitCode -eq 0) 
            {
                Write-Host
                Write-Host "$displayName installed successfully." -ForegroundColor White -BackgroundColor DarkGreen
            }
            else 
            {
                Write-Host
                Write-Host "$displayName did not install successfully." -ForegroundColor White -BackgroundColor Red
                Write-Host "The exit code was $exitCode" -ForegroundColor White -BackgroundColor Red
            } #end inner if-else
        }

        else 
        {
            Write-Host
            Write-Host "$displayName is already installed."
        } #end outer if-else

    } #end foreach
    
} #end function Install-PrereqMSIs


function Enable-MSDTC
{
    #Enable MSDTC Inbound and Outbound transactions via the registry
    $msdtcPath = 'HKLM:\SOFTWARE\Microsoft\MSDTC\Security'
    $msdtcProps = @(
                'NetworkDtcAccess'
                'NetworkDtcAccessTransactions'
                'NetworkDtcAccessInbound'
                'NetworkDtcAccessOutbound'
                )
    
    Write-Host
    Write-Host "Enabling MSDTC Inbound and Outbound transactions..." -ForegroundColor White -BackgroundColor DarkGreen

    foreach ($msdtcProp in $msdtcProps) 
    {
        Set-ItemProperty -path $msdtcPath -Name $msdtcProp -Value 1
        Write-Host
        Write-Host "$msdtcPath\$msdtcProp = 1" -ForegroundColor White -BackgroundColor DarkGreen
    } #end foreach

    Restart-Service MSDTC
} #end function Enable-MSDTC



########
# MAIN #
########

Install-PrereqWindowsFeatures($preReqFeatures)
Install-PrereqMSIs
Enable-MSDTC