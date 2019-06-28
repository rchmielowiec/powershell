<#
Downloads contacts from active directory and add them as local contacts on Outlook.
Usefull when you want to have outlook contacts synced with your mobile phone. 
#>
#set your data here
$your_domain_name = 'contoso.com' #set your domain name
$users_container_path = 'OU=Office1,OU=Users,DC=contoso,DC=com' #set containers where users are available
$your_company_name = 'Contoso' # set your comany name

Clear-Host
####### Domain Contacts
if( Test-Connection $your_domain_name -Count 3 -Quiet){
Write-Host "Downloading active directory contacts"


$isp = [adsi] "LDAP://$users_container_path"
$searcher = New-Object System.DirectoryServices.DirectorySearcher $isp
$searcher.Filter = "(&(objectClass=user)(objectCategory=person)(!sAMAccountType=805306370)(!userAccountControl:1.2.840.113556.1.4.803:=2)(sn=*)(mail=*)(mobile=*))"
$domain_contacts = New-Object System.Collections.ArrayList


$searcher.FindAll() | % {

  $user = [adsi]$_.Properties.adspath[0]
  
  $u = New-Object -Type PSCustomObject -Property @{
                                                    
                                                    Email1Address         = $user.mail[0]
                                                    FirstName             = (Get-Culture).TextInfo.ToTitleCase((Get-Culture).TextInfo.ToLower($user.givenName[0]))
                                                    LastName             =  (Get-Culture).TextInfo.ToTitleCase((Get-Culture).TextInfo.ToLower($user.sn[0]))                                                   
                                                    #CompanyName          = $user.CompanyName[0]
                                                     CompanyName          = $your_company_name
                                                    #HomeAddress          = $user.physicalDeliveryOfficeName[0]
                                                    MobileTelephoneNumber = $user.Mobile[0] -replace ' ' 
                                                    PrimaryTelephoneNumber = $user.telephoneNumber[0] -replace ' ' 
                                                    Picture = $user.thumbnailPhoto        
                                                    
                                                  }
  $domain_contacts.Add($u) | Out-Null
 
}

$domain_contacts = $domain_contacts | Sort-Object -Property Email1Address
function get-userpicture($user){
$file_path = $env:TEMP+ "\" + $user.FirstName + "_" + $user.LastName + ".jpg"

$user | select -ExpandProperty  Picture | Set-Content -Path $file_path -Encoding Byte

$file_path

}

<#
foreach ($u in $domain_contacts)
{
    if($u.Picture){get-userpicture $u}
}
#>

######Outlook contacts
Write-Host "Downloading outlook contacts."

#Get-Process | where {$_.Name -match "outlook"} | Stop-Process

$Outlook=NEW-OBJECT –comobject Outlook.Application
$outlook_contacts=$Outlook.session.GetDefaultFolder(10).items | Sort-Object -Property Email1Address


$old_contacts = New-Object System.Collections.ArrayList

$outlook_contacts | foreach{
    $contact = [PSCustomObject]@{

                Email1Address = $_.Email1Address
                CompanyName = $_.CompanyName
                #HomeAddress = $_.HomeAddress
                #(Get-Culture).TextInfo.ToTitleCase((Get-Culture).TextInfo.ToLower($_.FirstName)
                FirstName = $_.FirstName
                #(Get-Culture).TextInfo.ToTitleCase((Get-Culture).TextInfo.ToLower($_.LastName)
                LastName = $_.LastName
                MobileTelephoneNumber = $_.MobileTelephoneNumber -replace ' '
                PrimaryTelephoneNumber = $_.PrimaryTelephoneNumber -replace ' '

        } 
        if($contact.Email1Address){ $old_contacts.add($contact) | Out-Null}
}

########compare contacts

function add-outlook_contact($user)
    {

    #$Outlook=NEW-OBJECT –comobject Outlook.Application
    Write-Host "Adding user: " $user.Email1Address $(get-date -format HH:mm:ss)
    $new_contact = $Outlook.session.GetDefaultFolder(10).Items.Add()

    $new_contact.Email1Address = $user.Email1Address
    $new_contact.CompanyName = $user.CompanyName
    #$new_contact.HomeAddress = $user.HomeAddress
    $new_contact.FirstName =  $user.FirstName
    $new_contact.LastName = $user.LastName
    $new_contact.MobileTelephoneNumber = $user.MobileTelephoneNumber
    $new_contact.PrimaryTelephoneNumber = $user.PrimaryTelephoneNumber
    
    if($user.Picture){$new_contact.addPicture($(get-userpicture $user)) }
    
    $new_contact.Save()

    }

function update-outlook_contact($user)
    {

    #$Outlook=NEW-OBJECT –comobject Outlook.Application
    Write-Host "Updating user: " $user.Email1Address $(get-date -format HH:mm:ss)
    $mail = $user.Email1Address
    $contact_to_update = $outlook_contacts | where {$_.Email1Address -eq $mail} | select -First 1
    $contact_to_update.CompanyName = $user.CompanyName
    #$contact_to_update.HomeAddress = $user.HomeAddress
    #(Get-Culture).TextInfo.ToTitleCase((Get-Culture).TextInfo.ToLower($_.FirstName)
    $contact_to_update.FirstName =  $user.FirstName
    $contact_to_update.LastName = $user.LastName
    $contact_to_update.MobileTelephoneNumber = $user.MobileTelephoneNumber
    $contact_to_update.PrimaryTelephoneNumber = $user.PrimaryTelephoneNumber
    $contact_to_update.Save()
    
    }

function compare-contacts ($old_contact, $new_contact)
{
    $equal = $true
    if($old_contact.CompanyName -ne $new_contact.CompanyName){$equal = $false}#;write-host "CompanyName" $old_contact.CompanyName " <> " $new_contact.CompanyName}
    #if($old_contact.HomeAddress -ne $new_contact.HomeAddress){$equal = $false;write-host "HomeAddress" "|" $old_contact.HomeAddress "|" " <> " "|" $new_contact.HomeAddress "|"}
    if($old_contact.FirstName -ne $new_contact.FirstName){$equal = $false}#;write-host $old_contact.FirstName " <> " $new_contact.FirstName}
    if($old_contact.LastName -ne $new_contact.LastName){$equal = $false}#;write-host $old_contact.LastName " <> " $new_contact.LastName}
    if($old_contact.MobileTelephoneNumber -ne $new_contact.MobileTelephoneNumber){$equal = $false}#;write-host "MobileTelephoneNumber" $old_contact.MobileTelephoneNumber " <> " $new_contact.MobileTelephoneNumber}
    if($old_contact.PrimaryTelephoneNumber -ne $new_contact.PrimaryTelephoneNumber){$equal = $false}#;write-host "PrimaryTelephoneNumber" $old_contact.PrimaryTelephoneNumber " <> " $new_contact.PrimaryTelephoneNumber}
    $equal
    #if(!$equal){$old_contact.Email1Address}
}

$contacts_to_update = New-Object System.Collections.ArrayList

$new_contacts = New-Object System.Collections.ArrayList



if($old_contacts){
        Write-Host "Checking existing contacts."


        Compare-Object $domain_contacts.Email1Address $old_contacts.Email1Address -IncludeEqual | ForEach-Object {

             $mail = $_.InputObject

             if($_.SideIndicator -eq '<='){$new_contacts.add($($domain_contacts | where {$_.Email1Address -eq $mail})) | Out-Null} #create list of contacts to add to outlook
 
             if($_.SideIndicator -eq '=='){ #create list of contacts that should be updated
                                
                                            $domain_contact = $domain_contacts | where {$_.Email1Address -eq $mail}
                                            $old_contact = $old_contacts | where {$_.Email1Address -eq $mail}
                                            if(!$(compare-contacts  $old_contact  $domain_contact)){$contacts_to_update.Add($domain_contact) | Out-Null}
                                            }
 
        }



        Write-Host "Number of contacts to be updated in outlook: " $contacts_to_update.Count

        Write-Host "Number of new contacts from active directory: " $new_contacts.Count

        
        $contacts_to_update | foreach {update-outlook_contact($_)}
        


}else{$new_contacts = $domain_contacts}

###############add new contacts
if($new_contacts){$new_contacts | foreach {add-outlook_contact($_)}}

#clean temp
Get-ChildItem -path "$env:TEMP\*.jpg" -File | Remove-Item -ErrorAction SilentlyContinue
###############
Write-Host "Total number of contacts in Outlook: " $($Outlook.session.GetDefaultFolder(10).items).Count
write-host "Finished at:" $(get-date -Format HH:mm:ss)
#$Outlook.session.GetDefaultFolder(10).items | sort-object Email1Address | select Email1Address,FirstName,LastName,MobileTelephoneNumber,PrimaryTelephoneNumber,CompanyName | FT -AutoSize
}else{Write-Host "Cannot connect to active directory. Check network connection" -BackgroundColor Red}