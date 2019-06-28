<#provides interface for user to set hist mobile pgone number in active directory
Also sends email with update to support team.#>

$support_team_mail = 'email@testmail.your'

Add-Type -AssemblyName System.DirectoryServices.AccountManagement
Add-Type -AssemblyName PresentationFramework


[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

function send-message ($p, $user_fullname, $user_mail, $to)
{
    $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)
    $Mail.To = $to
    $Mail.CC = $user_mail
    $Mail.Subject = 'My new mobile phone number'   
    $message_text=@"
<p>Hi,</p>
<p>I have new mobile phone number: <b>$p</b><br>Please, update my information.</p>
<p>Regards</p>
<p><b>$user_fullname</b></p>
<p>Tech<br>My addresss</p>
<p><small>----------<br>Message created automatically when script is used by user to update mobile phone number.</small></p>
"@

    $Mail.HTMLBody = $message_text
    $Mail.Send() | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    $Outlook = $null
}


function get-mobile ()
        {
            if(Test-Connection $env:USERDNSDOMAIN -Quiet -Count 2)
            {
                           
                $ctype = [System.DirectoryServices.AccountManagement.ContextType]::Domain
                $user = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($ctype, $env:USERNAME)
                $u = $user.GetUnderlyingObject()
                $u.mobile
            }
            else
            {
                [void][System.Windows.MessageBox]::Show("Connection to domain failed.`r`nPlease, make sure you have working network connection. ","Error",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Error)
                exit
            }

        }


function set-mobile ($phone_number)
        {
                    if ($phone_number -match '^(?:\+48)?[0-9]{9}$')
                        {
                            $ctype = [System.DirectoryServices.AccountManagement.ContextType]::Domain
                            $user = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($ctype, $env:USERNAME)
                            $u = $user.GetUnderlyingObject()
                            $u.mobile = $phone_number
                            $u.CommitChanges()
                            $Form.Close()
                            Start-Sleep -Seconds 1
                                $phone = get-mobile
                            if ($phone -eq $phone_number)
                                {
                                    $user_fn = (Get-Culture).TextInfo.ToTitleCase((Get-Culture).TextInfo.ToLower($u.givenName)) +" " + (Get-Culture).TextInfo.ToTitleCase((Get-Culture).TextInfo.ToLower($u.sn))
                                    [void][System.Windows.MessageBox]::Show("Phone succesfully set to: $phone")
                                    send-message -p $phone -user_fullname $user_fn -user_mail $($u.mail.ToString()) -to $support_team_mail
                                }
                            else
                                 {
                                    [void][System.Windows.MessageBox]::Show("Changing phone failed, please try again or contact support squad.","Warning",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Error)
                                 } 
                         } 
                      else
                         {
                            [void][System.Windows.MessageBox]::Show("Please provide number in format: +48123456789","Warning",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Error)
                         }            
                    
    
        }





$current_mobile = get-mobile

$Form = New-Object System.Windows.Forms.Form    
$Form.Size = New-Object System.Drawing.Size(310,160)  
$Form.StartPosition = "CenterScreen" #loads the window in the center of the screen
$Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow #modifies the window border
$Form.TopMost = $true
$Form.Text = "User: $env:USERDOMAIN\$env:USERNAME" #window description

$label = New-Object Windows.Forms.Label
$label.Location = New-Object System.Drawing.Size(20,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Font = New-Object System.Drawing.Font("Verdana",12,[System.Drawing.FontStyle]::Bold)
$label.text = "Your phone: " + $current_mobile
$Form.Controls.Add($label)


$InputBox_phone = New-Object System.Windows.Forms.TextBox 
$InputBox_phone.Location = New-Object System.Drawing.Size(130,50) 
$InputBox_phone.Size = New-Object System.Drawing.Size(150,50) 
$Form.Controls.Add($InputBox_phone) 

$label_c = New-Object Windows.Forms.Label
$label_c.Location = New-Object System.Drawing.Size(20,50)
$label_c.Size = New-Object System.Drawing.Size(100,20)
$label_c.Font = New-Object System.Drawing.Font("Verdana",12,[System.Drawing.FontStyle]::Bold)
$label_c.text = "Change to: "
$Form.Controls.Add($label_c)


$Tooltip = New-Object System.Windows.Forms.ToolTip


$Button = New-Object System.Windows.Forms.Button 
$Button.Location = New-Object System.Drawing.Size(20,80) 
$Button.Size = New-Object System.Drawing.Size(100,30) 
$Button.Text = "Change" 
$Button.Cursor = [System.Windows.Forms.Cursors]::Hand
$Button.Add_Click({set-mobile($InputBox_phone.Text)}) 
$Form.Controls.Add($Button) 
$Tooltip.SetToolTip($Button,"Change phone number.")


$Button_cancel = New-Object System.Windows.Forms.Button 
$Button_cancel.Location = New-Object System.Drawing.Size(180,80) 
$Button_cancel.Size = New-Object System.Drawing.Size(100,30) 
$Button_cancel.Text = "Cancel" 
$Button_cancel.Cursor = [System.Windows.Forms.Cursors]::Hand
$Button_cancel.Add_Click({[void]$Form.Close()}) 
$Form.Controls.Add($Button_cancel) 
$Tooltip.SetToolTip($Button_cancel,"Cancel")



$Form.Add_Shown({$Form.Activate()})
[void] $Form.ShowDialog()


