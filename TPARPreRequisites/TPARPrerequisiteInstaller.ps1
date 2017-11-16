<#
TPAR Prerequisite Installer v1.0
Joe Beaulieu
ERP Technology Integration
Oct 2017

Contributions:  https://stackoverflow.com/questions/31712686/how-to-check-if-a-program-is-installed-and-install-it-if-it-is-not
                    
#>


$tempdir = (Get-Location).toString()
$preReqFeatures = Get-Content $tempdir\WindowsFeatures.txt
$SQLSysClrTypes = '*Microsoft System CLR Types for SQL Server 2012 (x64)*'
$ReportViewer2012 = '*Microsoft Report Viewer 2012 Runtime*'
$msiFile = $tempdir+"\SQLSysClrTypes.msi"
$msiArgs = "/qb"


foreach ($feature in $preReqFeatures) {
    $wf = Get-WindowsFeature -Name $feature
    if (!$wf.Installed) {
        if (!$wf) {
            Write-Host "`nFeature" $feature "not found." -ForegroundColor Yellow
        }
        else {
            Write-Host "`nInstalling" $wf.DisplayName -ForegroundColor White -BackgroundColor DarkGreen
            Install-WindowsFeature $feature
        }
    }
    else{
        Write-Host "`nThe feature" $wf.DisplayName "is already installed." -ForegroundColor DarkYellow
    }
}

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
    }
    Get-ItemProperty $regpath | .{process{if($_.DisplayName -and $_.UninstallString) { $_ } }} | Select DisplayName, Publisher, InstallDate, DisplayVersion, UninstallString |Sort DisplayName
}

#Check for MS System CLR Types for SQL 2012; install if needed
$result = Get-InstalledApps | where {$_.DisplayName -like $SQLSysClrTypes}

if ($result -eq $null) {
    $exitCode = (Start-Process -FilePath $msiFile -ArgumentList $msiArgs -Wait -Passthru).ExitCode
    if ($exitCode -eq 0) {
        Write-Host "`nMicrosoft System CLR Types for SQL Server 2012 (x64) installed successfully." -ForegroundColor White -BackgroundColor DarkGreen
    }
    else {
        Write-Host "`nMicrosoft System CLR Types for SQL Server 2012 (x64) did not install successfully.`nThe exit code was" $exitCode -ForegroundColor White -BackgroundColor Red
    }
}

else {
    Write-Host "`nMicrosoft System CLR Types for SQL Server 2012 (x64) is already installed."
}

#Check for MS Report Viewer 2012 Runtime; install if needed
$msiFile = $tempdir+"\ReportViewer.msi"
$result = Get-InstalledApps | where {$_.DisplayName -like $ReportViewer2012}

if ($result -eq $null) {
    $exitCode = (Start-Process -FilePath $msiFile -ArgumentList $msiArgs -Wait -Passthru).ExitCode
    if ($exitCode -eq 0) {
        Write-Host "`nMicrosoft Report Viewer 2012 Runtime installed successfully.`n" -ForegroundColor White -BackgroundColor DarkGreen
    }
    else {
        Write-Host "`nMicrosoft Report Viewer 2012 Runtime did not install successfully.`nThe exit code was" $exitCode "`n" -ForegroundColor White -BackgroundColor Red
    }
}

else {
    Write-Host "`nMicrosoft Report Viewer 2012 Runtime is already installed."
}


#Enable MSDTC Inbound and Outbound transactions via the registry
$msdtcPath = 'HKLM:\SOFTWARE\Microsoft\MSDTC\Security'
$msdtcProps = @("NetworkDtcAccess"
                "NetworkDtcAccessTransactions"
                "NetworkDtcAccessInbound"
                "NetworkDtcAccessOutbound"
                )

Write-Host "`nEnabling MSDTC Inbound and Outbound transactions...`n" -ForegroundColor White -BackgroundColor DarkGreen
foreach ($msdtcProp in $msdtcProps) {
    Set-ItemProperty -path $msdtcPath -Name $msdtcProp -Value 1
    Write-Host $msdtcPath"\"$msdtcProp "= 1" -ForegroundColor White -BackgroundColor DarkGreen
}

Restart-Service MSDTC