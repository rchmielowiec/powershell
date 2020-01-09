#used to create software update groups and packages and download updates in SCCM
#patch tuesday
#IsExpired                          : False
#IsSuperseded                       : False
#LocalizedCategoryInstanceNames, LocalizedDisplayName, LocalizedDescription
$site = 'YourSCCMSite:'
Import-Module -Name "$(split-path $Env:SMS_ADMIN_UI_PATH)\ConfigurationManager.psd1"
set-location -Path $site
$today_is = get-date -Format yyyy-MM-dd
$updates = Get-CMSoftwareUpdate -DatePostedMin $(get-date).AddDays(-1)
foreach ($update in $updates)
    {
        if (!($update.LocalizedCategoryInstanceNames).Contains('Windows Defender') -and !($update.LocalizedCategoryInstanceNames).Contains('Forefront Endpoint Protection 2010')){$update.LocalizedDisplayName}
    }

$Windows7_Updates = New-Object System.Collections.ArrayList
#$Windows7_Updates_SUGName = $today_is + '_PatchTuesday_Windows7'
$Windows8_Updates = New-Object System.Collections.ArrayList
#$Windows8_Updates_SUGName = $today_is + '_PatchTuesday_Windows8_81'
$Windows10LTSB_Updates = New-Object System.Collections.ArrayList
#$Windows10LTSB_Updates_SUGName = $today_is + '_PatchTuesday_Windows10_LTSB'
$Windows10_Updates = New-Object System.Collections.ArrayList
#$Windows10_Updates_SUGName = $today_is + '_PatchTuesday_Windows10'
$Office2013_Updates = New-Object System.Collections.ArrayList
#$Office2013_Updates_SUGName = $today_is + '_PatchTuesday_Office2013'
$Office2016_Updates = New-Object System.Collections.ArrayList
#$Office2016_Updates_SUGName = $today_is + '_PatchTuesday_Office2016'
$Other_Updates = New-Object System.Collections.ArrayList
#$Other_Updates_SUGName = $today_is + '_PatchTuesday_NotSupportedUpdates'

#Windows Defender
#Security Updates
#Office 2016
#Office 2013
#Windows 7
#Windows 8.1
#Windows 10
#Update Rollups
#Windows Server 2008
#Windows 8
#Windows 10 LTSB
#Windows Server 2012
#Windows Server 2012 R2
#Windows Server 2016
#Critical Updates
#Windows 10 Dynamic Update
#Windows Server 2008 R2

foreach ($update in $updates)
{
    
    if (!($update.LocalizedCategoryInstanceNames).Contains('Windows Defender') -and !($update.LocalizedCategoryInstanceNames).Contains('Forefront Endpoint Protection 2010'))
        {
            if (($update.LocalizedCategoryInstanceNames).Contains('Windows 10 LTSB')){$Windows10LTSB_Updates.Add($update) | Out-Null}
            elseif (!($update.LocalizedCategoryInstanceNames).Contains('Windows 10 LTSB') -and ($update.LocalizedCategoryInstanceNames).Contains('Windows 10') `
            -and ($update.LocalizedDisplayName -notmatch 'ARM64')){$Windows10_Updates.Add($update) | Out-Null}
            elseif (($update.LocalizedCategoryInstanceNames).Contains('Windows 8.1')){$Windows8_Updates.Add($update) | Out-Null}
            elseif (($update.LocalizedCategoryInstanceNames).Contains('Windows 8')){$Windows8_Updates.Add($update) | Out-Null}
            elseif (($update.LocalizedCategoryInstanceNames).Contains('Office 2013')){$Office2013_Updates.Add($update) | Out-Null}
            elseif (($update.LocalizedCategoryInstanceNames).Contains('Office 2016')){$Office2016_Updates.Add($update) | Out-Null}
            elseif (($update.LocalizedCategoryInstanceNames).Contains('Windows 7') -and ($update.LocalizedDisplayName -match 'x64')){$Windows7_Updates.Add($update) | Out-Null}
            else{$Other_Updates.Add($update) | Out-Null}
           
            #$update.LocalizedDisplayName
        }

}

#$updates.LocalizedCategoryInstanceNames
#$Windows7_Updates.LocalizedDisplayName
#$Windows8_Updates.LocalizedDisplayName
#$Windows10LTSB_Updates.LocalizedDisplayName
#$Windows10_Updates.LocalizedDisplayName
#$Office2013_Updates.LocalizedDisplayName
#$Office2016_Updates.LocalizedDisplayName
#$Other_Updates.LocalizedDisplayName
#$Updates | select LocalizedDisplayName,PlatformType

$sug_name = $today_is +"_PatchTuesday_Win7_Office2013_2016"
New-CMSoftwareUpdateGroup -Name $sug_name -UpdateID ($Windows7_Updates +  $Office2013_Updates + $Office2016_Updates).CI_ID | Out-Null
Get-CMSoftwareUpdateGroup -Name $sug_name | select LocalizedDisplayName,NumberOfUpdates

$sug_name = $today_is +"_PatchTuesday_Win10_Win8_81"
New-CMSoftwareUpdateGroup -Name $sug_name -UpdateID ($Windows10LTSB_Updates + $Windows10_Updates + $Windows8_Updates).CI_ID | Out-Null
Get-CMSoftwareUpdateGroup -Name $sug_name | select LocalizedDisplayName,NumberOfUpdates

$sug_name = $today_is +"_PatchTuesday_Others_Unsupported"
New-CMSoftwareUpdateGroup -Name $sug_name -UpdateID $Other_Updates.CI_ID  | Out-Null
Get-CMSoftwareUpdateGroup -Name $sug_name | select LocalizedDisplayName,NumberOfUpdates
