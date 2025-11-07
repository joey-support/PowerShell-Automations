# HYPER-V SETUP
# Jeff Funk 2021
# Revised by Dan Ginnane 2025
#######################################

$RES_OS_VHDX    = "RESxxxx_OS.vhdx"
$RES_DATA_VHDX  = "RESxxxx_DATA.vhdx"

#######################################

$ScriptName     = "Hyper-V Setup"
$ScriptVer      = "4.0"

#######################################

#Grabs script path, extracts the folder it is in
Split-Path -Parent $Script:MyInvocation.MyCommand.Path 
Write-Debug (Split-Path -Parent $Script:MyInvocation.MyCommand.Path) - "Grabbing this path"

#Checks the user priv and elevates the session if need be
If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]'Administrator')) {
	Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
	Exit
}
Write-Debug "Running privlege checks"

# Get credentials from XML files in folder OR prompt user for them
function Set-CustomCreds {
	if (Test-Path ".\creds_domain.xml") {$script:domainacc = Import-CliXml -Path ".\creds_domain.xml"} #If we ever alter the creds, BE SURE TO CHANGE the names here
	if (Test-Path ".\creds_admin.xml") {$script:adminacc = Import-CliXml -Path ".\creds_admin.xml"}
	Write-Host "DomainAcc" $domainacc
	Write-Host "AdminAcc" $adminacc
	if ($script:domainacc -eq $null) {$script:domainacc = $host.ui.PromptForCredential($ScriptName, "Please enter JOEY Domain Admin credentials", "JOEY\admin$", ""); $domainacc | Export-CliXml -Path ".\creds_domain.xml"}
	if ($script:adminacc -eq $null) {$script:adminacc = $host.ui.PromptForCredential($ScriptName, "Please enter local Administrator credentials", "Administrator", ""); $adminacc | Export-CliXml -Path ".\creds_admin.xml"}
}
$script:userpassword = ConvertTo-SecureString -String 'user' -AsPlainText -Force

Write-Debug "Grabbing credentials from folder XML"

# Function to create switch name so we can use it later when speaking with the HV
function Set-SwitchNameHV {
    if ($script:switch -eq $null -or $script:switch -notlike '*[a-z0-9]*') {
        # Prompt user to enter switch name in CLI
        $script:switch = Read-Host "Please enter the name of the Hyper-V Switch"
        if (-not (Get-VMSwitch -Name $script:switch -ea SilentlyContinue)) {
            Write-Warning "Switch '$script:switch' does not exist in this HyperV."
        }
    }
}

# If creds weren't set
Set-CustomCreds
    If ($adminacc -eq $null) {Write-Output "Credentials not captured"; return}

# Set Hyper-V switch name
Set-SwitchNameHV
If ($script:switch -eq "<UNDEFINED>") {
    Write-Output "Switch undefined or incorrect - aborting"
    return
    }

$StartTime = $(Get-Date)
Write-Output = "Creating $RES_Name - Start Time: $StartTime"

# GET VM INFO - This is the point at which we extract VM configs from the .csv to push to the HyperV
$csv = Import-Csv VMinfo.csv | Where-Object {$_.VM_Name -like "RES*"}
$csv | foreach {

    # Extract required values from csv
	$RES_Name      = $csv.VM_Name
	$RES_IP        = $csv.IP_Address
	$RES_SM        = $csv.Subnet_Mask
	$RES_GW        = $csv.Gateway
	$RES_DNS1      = $csv.DNS1
	$RES_DNS2      = $csv.DNS2
	$RES_VLAN      = $csv.VLAN
	
    # Convert subnet mask to CIDR
	$mask = [ipaddress]$RES_SM
	$binary = [convert]::ToString($mask.Address, 2)
	$mask_length = ($binary -replace 0,$null).Length
	$RES_cidr = '{0}' -f $mask_length

    Write-Host "VM values stored within: $RES_Name | IP: $RES_IP | GW: $RES_GW | $RES_DNS1 | $RES_DNS2 | $RES_VLAN"
}

# CREATION OF THE VM - This is the actual creation of the VM and where we pass the values over

# Define VHDs paths
$Final_RES_OS = "D:\Hyper-V\Virtual Hard Disks\"+$RES_Name+"_OS.vhdx"
$Final_RES_DATA = "D:\Hyper-V\Virtual Hard Disks\"+$RES_Name+"_DATA.vhdx"
Write-Debug "The following are being written to"
Write-Debug $Final_RES_OS
Write-Debug $Final_RES_DATA

# Map Z: to network drives so we can pull from them
Write-Debug "Mapping to :Z drive"
New-PSDrive -Name Z -PSProvider FileSystem -Root "\\192.168.5.72\misutil$" -Persist

# Create a network path using the local Z: folder
$NetworkDir = "Z:\INSTALL FILES\ServerVMs\Setup\RES\VHDs"

#Convert-VHD -DestinationPath $Final_RES_OS -Path .\VHDs\$RES_OS_VHDX -VHDType Fixed
Write-Debug "Converting the VHDs from net drive to target local server D: drive"
Convert-VHD -Path "$NetworkDir\$RES_OS_VHDX" -DestinationPath $Final_RES_OS -VHDType Fixed #Converts from source to -> destination path
Convert-VHD -Path "$NetworkDir\$RES_DATA_VHDX" -DestinationPath $Final_RES_DATA -VHDType Fixed

# Create and config VM

Write-Debug "Creating new VM name"
New-VM -Name $RES_Name -MemoryStartupBytes 8GB -VHDPath $Final_RES_OS -Path "D:\Hyper-V\Virtual Machines" -Switch "External Switch" -Generation 2

Write-Debug "Adding VM hard disk drive"
Add-VMHardDiskDrive -VMName $RES_Name -Path $Final_RES_DATA

Write-Debug "Adding processors"
Set-VMProcessor -VMName $RES_Name -Count 4 -RelativeWeight 100

Write-Debug "Setting network adapter VLAN"
Set-VMNetworkAdapterVlan -VMName $RES_Name -Access -VlanId $RES_VLAN 

Write-Debug "Starting VM"
Start-VM $RES_Name

Write-Host "Finished! VM should be made, waiting 60 seconds before execution of VM commands"

Start-Sleep 60
#Create a VM just as a test, solve the problem of simply being able to create it don't assign a converted drive to it. Just to it all for test.

# GET VM INFO - This is the point at which we extract VM configs from the .csv to push to the HyperV
$csv = Import-Csv VMinfo.csv | Where-Object {$_.VM_Name -like "RES*"}
$csv | foreach {

    # Extract required values from csv
	$RES_Name      = $csv.VM_Name
	$RES_IP        = $csv.IP_Address
	$RES_SM        = $csv.Subnet_Mask
	$RES_GW        = $csv.Gateway
	$RES_DNS1      = $csv.DNS1
	$RES_DNS2      = $csv.DNS2
	$RES_VLAN      = $csv.VLAN
	
    # Convert subnet mask to CIDR
	$mask = [ipaddress]$RES_SM
	$binary = [convert]::ToString($mask.Address, 2)
	$mask_length = ($binary -replace 0,$null).Length
	$RES_cidr = '{0}' -f $mask_length

    Write-Host "VM values stored within: $RES_Name | IP: $RES_IP | GW: $RES_GW | $RES_DNS1 | $RES_DNS2 | $RES_VLAN"
}


function Set-CustomCreds {
	if (Test-Path ".\creds_domain.xml") {$script:domainacc = Import-CliXml -Path ".\creds_domain.xml"} #If we ever alter the creds, BE SURE TO CHANGE the names here
	if (Test-Path ".\creds_admin.xml") {$script:adminacc = Import-CliXml -Path ".\creds_admin.xml"}
	Write-Host "DomainAcc" $domainacc
	Write-Host "AdminAcc" $adminacc
}

$script:userpassword = ConvertTo-SecureString -String 'user' -AsPlainText -Force

$s = New-PSSession -VMName $RES_Name -Credential $adminacc #Change name to 414RES if it doesn't work
Invoke-Command -Session $s -ScriptBlock { 
    Get-NetAdapter | Get-NetIPAddress | Remove-NetIPAddress -Confirm:$false -ea SilentlyContinue
    Get-NetAdapter | Get-NetRoute | Remove-NetRoute -Confirm:$false -ea SilentlyContinue
    Get-NetIPConfiguration | New-NetIPAddress -IPAddress $Using:RES_IP -DefaultGateway $Using:RES_GW -PrefixLength $Using:RES_cidr
    Get-NetIPConfiguration | Set-DnsClientServerAddress -ServerAddresses $Using:RES_DNS1,$Using:RES_DNS2
    $job = 'D:\MICROS\Res\Pos\Etc\SetName.exe s'+$Using:RES_Name
    Invoke-Expression $job
    Sleep 30
    Write-Host $Using:RES_Name Rebooting... -ForegroundColor Green -BackgroundColor Black
    Restart-Computer -Force
    }
Write-Host "RES setup complete"
Remove-PSSession $s
