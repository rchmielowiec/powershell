<#automaticly downloads latest version of VirtualBox 
and creates installation package in sccm#>

Clear-Host
################change according to your environment
$7zip_path = 'C:\Program Files\7-Zip\7z.exe'
$SiteCode = "Site_Code"
$sccm_apps_repo_path = "\\sccm_name\sources$\Applications"
$working_dir = 'C:\Temp\'
################

$url = 'https://www.virtualbox.org/wiki/Downloads'
$docs_url = "https://www.virtualbox.org/wiki/Technical_documentation"

if(Test-Path $7zip_path)
    {write-host "7-zip installed"}
else
    {Write-Host "Please install 7-zip before starting";exit}

if(Test-Path (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1))
    {write-host 'System Center Configuration Manager Console installed'}
else
    {write-host 'Please install System Center Configuration Manager Console';Start-Sleep -Seconds 30; exit}

#1. Check and download virtualbox

$download_path = "$working_dir$(Get-Date -Format "dd-MM-yyyy")\virtualbox"

$proxy = [System.Net.WebRequest]::GetSystemWebProxy()
$proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
$WebClient = New-object System.Net.WebClient
$WebClient.UseDefaultCredentials = $true ## Proxy credentials only
$WebClient.Proxy.Credentials = $WebClient.Credentials
$info = Invoke-WebRequest -Uri $url
$download_url = ($info.Links | where {$_.OuterText -match "Windows hosts"}).href
$file_name = $download_url.Split("/")[-1]

Write-Host "Latest version avaiable: $file_name. Continue to download?"
Write-Host "Continue? (enter Y and click Enter)"
$continue = Read-Host
if($continue -ne 'Y'){Write-Host "Aborted. Script will exit in 30s";start-sleep -Seconds 30 ;exit}


if(!(Test-Path $download_path)){New-Item -ItemType Directory -Path $download_path -Force | Out-Null}
Write-Host "Downloading files"
$WebClient.DownloadFile($download_url,"$download_path\$file_name")
if(Test-Path "$download_path\$file_name"){Start-Process "$download_path\$file_name" -ArgumentList ("--extract","-path","$download_path","--silent") -NoNewWindow -Wait}
Write-Host "Extracting files"
if(Test-Path "$download_path\common.cab"){Start-Process "expand" -ArgumentList ("-f:*.iso","$download_path\common.cab","$download_path") -NoNewWindow -Wait}
$iso = Get-ChildItem -Path $download_path -Include "*.iso" -Recurse
Set-Location $download_path
if($iso){&$7zip_path x $iso.FullName}

###get information about msi file
$msi_path = Get-ChildItem -Path $download_path -Include "VirtualBox*amd64*.msi" -Recurse 

$WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer

$MSIDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase","InvokeMethod",$Null,$WindowsInstaller,@($msi_path.FullName,0))

$Query = "SELECT * FROM Property"

$View = $MSIDatabase.GetType().InvokeMember("OpenView","InvokeMethod",$null,$MSIDatabase,($Query))

$View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)

$Record = $View.GetType().InvokeMember("Fetch","InvokeMethod",$null,$View,$null)

$msi_props = @{}
        while ($record -ne $null) {
            $msi_props[$record.GetType().InvokeMember("StringData", "GetProperty", $Null, $record, 1)] = $record.GetType().InvokeMember("StringData", "GetProperty", $Null, $record, 2)
            $record = $View.GetType().InvokeMember(
                "Fetch",
                "InvokeMethod",
                $Null,
                $View,
                $Null
            )
        }
######Commit database and close view
        $MSIDatabase.GetType().InvokeMember("Commit", "InvokeMethod", $null, $MSIDatabase, $null)
        $View.GetType().InvokeMember("Close", "InvokeMethod", $null, $View, $null)           
        $MSIDatabase = $null
        $View = $null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null
######
$msi_info = New-Object -TypeName pscustomobject -Property $msi_props

#$msi_info


########

#2. Prepare files
Write-Host "Moving required files to new directory"
$dir_name = $download_path+"\VirtualBox_" + $msi_info.ProductVersion

New-Item -ItemType Directory -Path $dir_name
if(Test-Path $dir_name){
    Copy-Item -Path $msi_path -Destination $dir_name
    Copy-Item -Path $download_path\common.cab -Destination $dir_name
    Get-ChildItem -Path $download_path\cert | Copy-Item -Destination $dir_name    
}

[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')  | Out-Null
[System.Drawing.Icon]::ExtractAssociatedIcon("$download_path\$file_name").ToBitmap().Save("$dir_name\app.ico")

Write-Host 'Prepare scripts'

$install_oracle_certs_script = @'
$certs = Get-ChildItem -Include "*.cer" -Recurse
foreach ($cert in $certs)
{
    $cert_obj = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
    $cert_obj.Import($cert.fullname)
    if(Get-ChildItem -Path Cert:\LocalMachine\TrustedPublisher | where {$_.Thumbprint -eq $cert_obj.Thumbprint})
        {
            
        }
    else
        {
            .\VBoxCertUtil.exe add-trusted-publisher $cert.FullName
        }
Remove-Variable cert_obj -Force
}
'@

$install_oracle_certs_script | Out-File "$dir_name\install_oracle_certs.ps1"

$uninstall_oracle_certs_script = @'
$certs = Get-ChildItem -Include "*.cer" -Recurse

foreach ($cert in $certs)
{
    $cert_obj = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
    $cert_obj.Import($cert.fullname)
    Get-ChildItem -Path Cert:\LocalMachine\TrustedPublisher | where {$_.Thumbprint -eq $cert_obj.Thumbprint} | Remove-Item
    
Remove-Variable cert_obj -Force

}
'@
$uninstall_oracle_certs_script | Out-File "$dir_name\uninstall_oracle_certs.ps1"


$thumbs = New-Object System.Collections.ArrayList
$certs = Get-ChildItem $dir_name -Include "*.cer" -Recurse

foreach ($cert in $certs)
{
    $cert_obj = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
    $cert_obj.Import($cert.fullname)

            $thumbs.Add($cert_obj.Thumbprint) | Out-Null
Remove-Variable cert_obj -Force
}

$add_to_script = '$thumbs=("'+($thumbs -join ",").replace(',','","')+'")'+"`r`n"

$scriptDetection = @'
foreach ($thumb in $thumbs)
{
    if(!(Get-ChildItem -Path Cert:\LocalMachine\TrustedPublisher | where {$_.Thumbprint -eq $thumb})){exit}   
}
Write-Host "Compliant"
'@

$scriptDetection = $add_to_script+$scriptDetection

$scriptDetection | Out-File "$dir_name\certificates_scipt_detection_for_sccm.txt"

Write-Host "Done. At this point you have complete set of files redy to deploy in sccm located in $dir_name folder. You can stop here or continue to automaticly create SCCM Application"
Write-Host "Continue? (enter Y and click Enter)"
$continue = Read-Host
if($continue -eq 'Y')
{
Write-Host "Starting to create SCCM Appliaction"
######3.create sccm appliction

Import-Module(Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1)

Set-Location($SiteCode + ":") 


$ApplicationName = $msi_info.ProductName.Replace(' ','_')
$Publisher = $msi_info.Manufacturer
$ApplicationVersion = $msi_info.ProductVersion
$LocalizedApplicationName = $msi_info.ProductName
$UserDocumentation = $docs_url
$LocalizedApplicationDescription = @"
VirtualBox is a powerful x86 and AMD64/Intel64 virtualization product for enterprise as well as home use. Not only is VirtualBox an extremely feature rich, high performance product for enterprise customers, it is also the only professional solution that is freely available as Open Source Software under the terms of the GNU General Public License (GPL) version 2. See "About VirtualBox" for an introduction.
Presently, VirtualBox runs on Windows, Linux, Macintosh, and Solaris hosts and supports a large number of guest operating systems including but not limited to Windows (NT 4.0, 2000, XP, Server 2003, Vista, Windows 7, Windows 8, Windows 10), DOS/Windows 3.x, Linux (2.4, 2.6, 3.x and 4.x), Solaris and OpenSolaris, OS/2, and OpenBSD.
VirtualBox is being actively developed with frequent releases and has an ever growing list of features, supported guest operating systems and platforms it runs on. VirtualBox is a community effort backed by a dedicated company: everyone is encouraged to contribute while Oracle ensures the product always meets professional quality criteria.
"@

New-CMApplication -Name $ApplicationName -Publisher $Publisher -AutoInstall $true -SoftwareVersion $ApplicationVersion -LocalizedApplicationName $LocalizedApplicationName -LocalizedApplicationDescription $LocalizedApplicationDescription -UserDocumentation $UserDocumentation


New-PSDrive -PSProvider FileSystem -Root $sccm_apps_repo_path -Name "sccm_apps_root"

Copy-Item -Path $dir_name -Recurse -Destination sccm_apps_root:\

$app_folder_name = $dir_name.Split("\")[-1]


#### CREATE MSI DEPLOYMENT TYPE

$msi_productProductCode = $msi_info.ProductCode
$msi_file_name = (Get-ChildItem -Path sccm_apps_root:\$app_folder_name\ -Include "*.msi" -Recurse).Name

$DeploymentTypeHash = @{                    
                    Applicationname = "$ApplicationName" #Application Name 
                    DeploymentTypeName = "Install_$ApplicationName"    #Name given to the Deployment Type
                    InstallationFileLocation = "$sccm_apps_repo_path\$app_folder_name\$msi_file_name"  # NAL path to the package
                    MsiInstaller = $true
                    InstallationBehaviorType = 'InstallForSystem'
                    AutoIdentifyFromInstallationFile = $true
                    ForceForUnknownPublisher = $true                    
                    }

Add-CMDeploymentType @DeploymentTypeHash


$DeploymentTypeHash = @{
                    
                    Applicationname = "$ApplicationName" #Application Name 
                    DeploymentTypeName = "Install_$ApplicationName"    #Name given to the Deployment Type
                    InstallationProgram ='msiexec /i '+ '"' + "$msi_file_name" + '"' + ' VBOX_INSTALLDESKTOPSHORTCUT=0 VBOX_INSTALLQUICKLAUNCHSHORTCUT=0 VBOX_START=0 /qn /norestart'  #Command line to Run for install
                    UninstallProgram ='msiexec /x '+ $msi_productProductCode +' /qn /norestart'  #Command line to Run for un-Install
                    RequiresUserInteraction = $false  #Don't let User interact with this
                    EstimatedInstallationTimeMinutes = '10'
                    MaximumAllowedRunTimeMinutes = '20'
                    LogonRequirementType = 'WhereOrNotUserLoggedOn'
                    InstallationProgramVisibility = 'Hidden'
                    MsiOrScriptInstaller = $true
                    }

Set-CMDeploymentType @DeploymentTypeHash

##create script deploymenttype for certifiactes installation

$DeploymentTypeHash = @{
                    ManualSpecifyDeploymentType = $true #Yes we are going to manually specify the Deployment type
                    Applicationname = "$ApplicationName" #Application Name 
                    DeploymentTypeName = "oracle_certs_install"    #Name given to the Deployment Type                    
                    DetectDeploymentTypeByCustomScript = $true # Yes deployment type will use a custom script to detect the presence of this 
                    ScriptInstaller = $true # Yes this is a Script Installer
                    ScriptType = 'PowerShell' # yep we will use PowerShell Script
                    ScriptContent =$scriptDetection  # Use the earlier defined here string
                    AdministratorComment = "This will install Oracle root certificates" 
                    ContentLocation = "$sccm_apps_repo_path\$app_folder_name"  # NAL path to the package
                    InstallationProgram ='powershell.exe -executionpolicy bypass  -file ".\install_oracle_certs.ps1"'  #Command line to Run for install
                    UninstallProgram ='powershell.exe -executionpolicy bypass  -file ".\uninstall_oracle_certs.ps1"'  #Command line to Run for un-Install
                    RequiresUserInteraction = $false  #Don't let User interact with this
                    InstallationBehaviorType = 'InstallForSystem' # Targeting Devices here
                    InstallationProgramVisibility = 'Hidden'  # Hide the PowerShell Console
                    EstimatedInstallationTimeMinutes = '5'
                    MaximumAllowedRunTimeMinutes = '15'
                    LogonRequirementType = 'WhereOrNotUserLoggedOn'
                    }

Add-CMDeploymentType @DeploymentTypeHash
Get-CMDeploymentType -ApplicationName $ApplicationName | select -Last 1 | Set-CMDeploymentType -DeploymentTypeName 'oracle_certs_install' #fix bug that creates deploymenttype with '- script' name
######

Get-CMDeploymentType -ApplicationName $ApplicationName | select -First 1 | New-CMDeploymentTypeDependencyGroup -GroupName 'Required' | Add-CMDeploymentTypeDependency -DeploymentTypeDependency `
(Get-CMDeploymentType -ApplicationName $ApplicationName | select -Last 1) -IsAutoInstall $true

##############################################
Remove-PSDrive sccm_apps_root
}

#cleanup
Set-Location -Path 'c:\'
Get-ChildItem -Path $download_path -Exclude $dir_name.Split("\")[-1] | Remove-Item -Recurse -Force


<#to do

Add-CMDeploymentTypeDependency

Adds a deployment type as a dependency to a dependency group. Required input is a deployment type object from Get-CMDeploymentType and a dependency group from [Get|New]-CMDeploymentTypeDependencyGroup.

Example

Get-CMDeploymentType -ApplicationName MyApp |
New-CMDeploymentTypeDependencyGroup -GroupName MyGroup |
Add-CMDeploymentTypeDependency -DeploymentTypeDependency `
(Get-CMDeploymentType -ApplicationName MyChildApp) `
-IsAutoInstall $true


#>