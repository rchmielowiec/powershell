#used as a task in MS TFS pipeline to unzip file that was uploaded to remote host via "Copy file to remote host" task
param(
[string]$admin_user,[string]$admin_password,[string]$destination_host,[string]$path
)
$secpasswd = ConvertTo-SecureString "$admin_password"-AsPlainText -Force
$mycreds = New-Object System.Management.Automation.PSCredential ($admin_user, $secpasswd)

$command = {
    param([string]$path)
    
    $shell = new-object -com shell.application
    $zip_files = gci $path -File  "*.zip"
    foreach($file in $zip_files)
        {
            $zip = $shell.NameSpace($file.FullName)
            foreach($item in $zip.items())
            {
             $shell.Namespace($path).copyhere($item)
            }
         }
    $zip_files | remove-item
    $files = gci -Path $path -Recurse -Depth 3 | select FullName | sort 
    $files
         }


$result = Invoke-Command -ComputerName $destination_host -Credential $mycreds -ScriptBlock $command -ArgumentList $path -UseSSL
$result
