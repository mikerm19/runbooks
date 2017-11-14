#This requires: Posh-SSH https://www.powershellgallery.com/packages/Posh-SSH/2.0.2

#Configuration
$date = Get-Date
$resourcegroup = "mag-e-linux-rg"           #Resource group where the Linux machines reside you want to update.
$azureconnectioninfo = "AzureRunAsConnection" #Connection asset to the Azure subscription
$linuxconnectioninfo = "linuxautomation@medadvgrp.com" #Automation connection name (authentication for each server)
$azuresubscriptionid = "1e7b4083-3d34-4740-bad5-bd3f947539bc" #Azure subscription where the hosts are
$infocommand = "apt list --upgradable"      #Command to run to show what packages are being upgraded
$upgradecommand = '' #'sudo apt-get update; sudo apt-get -o Dpkg::Options::="--force-confdef" -o Dpkg::Options::="--force-confold" dist-upgrade -y; sudo reboot' #Actual upgrade command.
$emailfrom = "passwords@medadvgrp.com"      #The From address for the email
$emailto = "mmason@medadvgrp.com"#,"asoyez@medadvgrp.com" #Can be a comma separated and quoted text list. E.X. "email1@medadvgrp.com","email2@medadvgrp.com"
$emailsubject = 'Monthly Linux Upgrades Run for $date'
$emailsmtp = "smtp.office365.com"           #The SMTP server address to send from
$emailport = 587                            #SMTP server port


#######################################################################################################
# Main Script
#######################################################################################################
$error = ""

#Function to send out the actual email, either on error or at the end of the run.
function Send-Email {
    param(
     [string]$body
     )
    foreach ($email in $emailto) { 
        Send-MailMessage -To $email -SmtpServer $emailsmtp -Credential $emailcred -UseSsl -Port $emailport -Subject $ExecutionContext.InvokeCommand.ExpandString($emailsubject) -Body $ExecutionContext.InvokeCommand.ExpandString($body) -From $emailfrom -BodyAsHtml
    }
}

#Get automation connection that will connect to the Linux machines. If there is a problem, send an error email and exit.
try{
    $azureconnectioncred = Get-AutomationConnection -Name $azureconnectioninfo
    $linuxconnectioncred = Get-AutomationPSCredential -Name $linuxconnectioninfo
    $emailcred = Get-AutomationPSCredential -Name $emailfrom
}
catch{
    $error = "Unable to get an automation connection. Stopping run. See job log for details.<br>"
    Send-Email -body $error
    throw $error
    Exit
}


#Log in to the Azure subscription
try {
    Add-AzureRmAccount `
            -ServicePrincipal `
            -TenantId $azureconnectioncred.TenantId `
            -ApplicationId $azureconnectioncred.ApplicationId `
            -CertificateThumbprint $azureconnectioncred.CertificateThumbprint 
}
catch {
    if (!$azureconnectioninfo)
    {
        $error = "Connection $azureconnectioninfo not found. Stopping run. See job log for details.<br>"
        Send-Email -body $error
        throw $error
    } else{
        Write-Error -Message $_.Exception
        Send-Email -body $_.Exception
        throw $_.Exception
    }
}

#Make sure the right subscriptiong is selected if specified.
try {
    If ($azuresubscriptionid) {
        Select-AzureRmSubscription -SubscriptionId $azuresubscriptionid
    } else {
        Write-Host "Continuing without specific subscription ID."
    }
}
catch {
    Write-Error -Message $_.Exception
    Send-Email -bosy $_.Exception
    throw $_.Exception
}

#Get all of the Linux machines in the resource group. If there is a problem, send an error email and exit.
try{
    $vminfo = @()
    $vms = Get-AzureRmVM -ResourceGroupName $resourcegroup | Select -Property Name, @{Label="NetInterfaceId";Expression={$_.networkProfile.networkInterfaces.id}}, @{Label="VmSize";Expression={$_.HardwareProfile.VmSize}}, @{Label="OsType";Expression={$_.StorageProfile.OsDisk.OsType}}
    foreach ($vm in $vms){
            $vm.NetInterfaceId = $vm.NetInterfaceId.Substring($vm.NetInterfaceId.LastIndexOf('/')+1)
            $ip = Get-AzureRmNetworkInterface -ResourceGroupName $resourcegroup -Name $vm.NetInterfaceId | Get-AzureRmNetworkInterfaceIpConfig | Select PrivateIPAddress
            $vminfo += New-Object psobject -Property @{
                        'Name' = $vm.Name
                        'IP' = $ip.PrivateIpAddress
                        }
    }
}
catch{
    $error = "There was an error retrieving a list of VM's and their private IPs. Stopping run. See job log for details.<br>"
    Send-Email -body $error
    throw $error
    Exit
}

#Log in to each machine and perform update. Also get what is getting updated to send in an email. If there is an error, include in the results email.
foreach ($machine in $vminfo){
    try{
        $vmstatus = Get-AzureRmVM -ResourceGroupName $resourcegroup -Status $vm | Select @{Label="Status";Expression={$_.Statuses[1].Code}}
        If ($vmstatus.status = "PowerState/running"){ 
            $session = New-SSHSession -ComputerName $machine.Name -Credential ($linuxconnectioncred) -AcceptKey
            Write-Host "Listing upgrades for $machine.Name ($machine.IP) ..."
            $upgradelist = Invoke-SSHCommandStream -SSHSession $session -Command $infocommand
            Write-Host "Upgrading $machine.Name ($machine.IP) ..."
            Invoke-SSHCommand -SSHSession $session -Command $upgradecommand
            Remove-SSHSession -Index 0
            $body += "<p><b>Upgraded: $machine.Name ($machine.IP) </b><br><br> $upgradelist </p>"
        } else {
            $error += "Warning: Skipped $machine.Name because power state was: $vmstatus.status <br>"
        }
    }
    catch{
        $error += "Error: Skipped $machine.Name: Unable to log in to $machine.Name <br>"
    }
}
Send-Email -body "$body $error"