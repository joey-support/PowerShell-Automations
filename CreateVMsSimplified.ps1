#######################################

$RES_OS_VHDX    = "RESxxxx_OS.vhdx"
$RES_DATA_VHDX  = "RESxxxx_DATA.vhdx"
$LTSC_VHDX      = "LTSC_2019.vhdx"
$WKS_VHDX       = "PWKS-PLATINUM.vhdx"

#######################################

$ScriptName     = "Hyper-V Setup"
$ScriptVer      = "3.0"

#######################################

#Create WKS VMs
$domainacc = Get-Credential -Message "Domain Account"
$StoreCode = Read-Host "Enter the store code (ex 5044)"

$Switch = Read-Host "Enter the name of the hyper-v switch"
$StartTime = $(Get-Date)
Write-Host "Creating WKS VMs - Start Time: $StartTime"
$csv = Import-Csv VMinfo.csv | Where-Object {$_.VM_Name -like "*-*wks*"}
$csv | ForEach-Object {
	$name = $_.VM_Name    	
	Write-Host "Creating $name..."
	Convert-VHD -DestinationPath "D:\Hyper-V\Virtual Hard Disks\$name.vhdx" -Path .\VHDs\$WKS_VHDX -VHDType Fixed
	New-VM -Name $name -MemoryStartupBytes 2GB -VHDPath "D:\Hyper-V\Virtual Hard Disks\$name.vhdx" -Path "D:\Hyper-V\Virtual Machines" -Switch "$switch" -Generation 2
	Set-VMProcessor -VMName $name -Count 2 -RelativeWeight 125
	Set-VMNetworkAdapterVlan -VMName $name -Access -VlanId 10
	Start-VM $name
}

$elapsedTime = $(get-date) - $StartTime 
$totalTime = "{0:HH:mm:ss}" -f ([datetime]$elapsedTime.Ticks)
Write-Host "Completed WKS VMs - Total Time: $totalTime"

Write-Output "Joining to Domain..."
$csv | foreach-object {
	if ($_.VM_Name -like "*wks*") {
		$Name = $_.VM_Name
		Write-Host "Joining $Name to JOEY.local..."
		$username = ($Name + '\user')
		$userpassword = 'user' | convertto-securestring -AsPlainText -force
		$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username,$userpassword
		$s = New-PSSession -VMName $Name -Credential $creds
		Invoke-Command -Session $s -ScriptBlock {
			Add-Computer -DomainName joey.local -Credential $domainacc -Restart -Force
			}
		Remove-PSSession $s
	}
	else {
		$Name = $_.VM_Name
		Write-Host "Joining $Name to JOEY.local..."
		$s = New-PSSession -VMName $Name -Credential $adminacc
		Invoke-Command -Session $s -ScriptBlock {
			Add-Computer -DomainName joey.local -Credential $domainacc -Restart -Force
			}
		Remove-PSSession $s
	}
}

Write-Host "Installing ConnectWise Automate and Control"
		
New-PSDrive -Name "software" -PSProvider FileSystem -Root "\\joey-file\software$" -Credential $domainacc

$msi = $Storecode + ".msi"		
$CWautomate = (Get-ChildItem -Filter $msi -File -Path "\\joey-file\software$\ConnectWise\#\AUTOMATE").FullName
$CWcontrol = (Get-ChildItem -Filter $msi -File -Path "\\joey-file\software$\ConnectWise\#\CONTROL").FullName
Write-Output "$CWcontrol"
Write-Output "$CWautomate"
If ($CWautomate -eq $null -OR $CWcontrol -eq $null) {Write-Output "MSI files not found"; return}
mkdir C:\temp | Out-Null
Copy-Item -Path $CWcontrol -Destination C:\temp\CWcontrol.msi -Force
Copy-Item -Path $CWautomate -Destination C:\temp\CWautomate.msi -Force

$csv | foreach-object {
	if ($_.VM_Name -like "*wks*"){
		$Name = $_.VM_Name
		Write-Host "Installing on $Name..."
		$username = ($Name + '\user')
		$userpassword = 'user' | convertto-securestring -AsPlainText -force
		$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username,$userpassword
		$s = New-PSSession -VMName $Name -Credential $creds
		Copy-Item -ToSession $s -Path C:\temp\CWcontrol.msi -Destination C:\
		Invoke-Command -Session $s -ScriptBlock { 
			msiexec /i C:\CWcontrol.msi /quiet /norestart 
			}
		Remove-PSSession $s
	}
	else {
		$Name = (Get-VM $_.VMName).Name
		Write-Output "Installing on $Name..."
		$s = New-PSSession -VMName $Name -Credential $adminacc
		Copy-Item -ToSession $s -Path C:\temp\CWcontrol.msi -Destination C:\
		Copy-Item -ToSession $s -Path C:\temp\CWautomate.msi -Destination C:\
		Invoke-Command -Session $s -ScriptBlock { 
			msiexec /i C:\CWcontrol.msi /quiet /norestart 
			Start-Sleep 30
			msiexec /i C:\CWautomate.msi /quiet /norestart
			}
		Remove-PSSession $s
	}
}

net use "\\joey-file\software$" /d /y


#=================================================================================================================

# CAL VMs
# $StartTime = $(Get-Date)
# Write-Host "Running AutoCAL on WKS VMs - Start Time: $StartTime"

# $csv | foreach-object{	
# 	$VM = Get-VM $_.VM_Name	
# 	$VMdisplayname = $VM.VMName
# 	$username = ($VM.VMName + '\user')
# 	$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username,$userpassword
# 	$s = New-PSSession -VMName $_.VMName -Credential $creds
# 	$CSV_VM_Name = $null		
# 	$csv = Import-Csv VMinfo.csv | Where-Object {$VM.VM_Name -eq $VMdisplayname}
# 	$csv | foreach {
# 		$CSV_VM_Name = $csv.VM_Name
# 		$CSV_IP_Address = $csv.IP_Address
# 		$CSV_Subnet_Mask = $csv.Subnet_Mask
# 		$CSV_Gateway = $csv.Gateway
# 		$CSV_DNS1 = $csv.DNS1
# 		$CSV_DNS2 = $csv.DNS2
# 		$CSV_VLAN = $csv.VLAN
# 	}
					
# 	if ($CSV_VM_Name -eq $null) {return}

# 	Write-Host "Setting $CSV_VM_Name to AutoCAL..."
		
# 	$mask = [ipaddress]$CSV_Subnet_Mask
# 	$binary = [convert]::ToString($mask.Address, 2)
# 	$mask_length = ($binary -replace 0,$null).Length
# 	$cidr = '{0}' -f $mask_length
			
# 	Set-VMNetworkAdapterVlan -VMName $_.VMName -Access -VlanId $CSV_VLAN 
	
# 	Invoke-Command -Session $s -ScriptBlock { 
# 		Invoke-Expression "taskkill /f /t /im DbUpdateServer.exe" -ea SilentlyContinue
# 		Invoke-Expression "taskkill /f /t /im KDSController.exe" -ea SilentlyContinue
# 		Invoke-Expression "taskkill /f /t /im MDSHTTPService.exe" -ea SilentlyContinue
# 		Invoke-Expression "taskkill /f /t /im McrsCal.exe" -ea SilentlyContinue
# 		Invoke-Expression "taskkill /f /t /im Ops.exe" -ea SilentlyContinue
# 		Invoke-Expression "taskkill /f /t /im Periphs.exe" -ea SilentlyContinue
# 		Invoke-Expression "taskkill /f /t /im WaitForHostsFile.exe" -ea SilentlyContinue
# 		Invoke-Expression "taskkill /f /t /im Win7CALStart.exe" -ea SilentlyContinue
# 		Get-NetAdapter | Get-NetIPAddress | Remove-NetIPAddress -Confirm:$false -ea SilentlyContinue
# 		Get-NetAdapter | Get-NetRoute | Remove-NetRoute -Confirm:$false -ea SilentlyContinue
# 		Get-NetIPConfiguration | New-NetIPAddress -IPAddress $Using:CSV_IP_Address -DefaultGateway $Using:CSV_Gateway -PrefixLength $Using:cidr
# 		Get-NetIPConfiguration | Set-DnsClientServerAddress -ServerAddresses $Using:CSV_DNS1,$Using:CSV_DNS2
# 		rm C:\micros -r -fo -ea SilentlyContinue
# 		rm C:\CALTemp -r -fo -ea SilentlyContinue
# 		Remove-Item "HKLM:\SOFTWARE\Wow6432Node\Micros" -r -force -ea SilentlyContinue
# 		New-Item "HKLM:\SOFTWARE\Wow6432Node\Micros" -force -ea SilentlyContinue
# 		New-Item "HKLM:\SOFTWARE\Wow6432Node\Micros\CAL" -force -ea SilentlyContinue
# 		New-Item "HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config" -force -ea SilentlyContinue
# 		New-Item "HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Scripts" -force -ea SilentlyContinue
# 		New-Item "HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Scripts\McrsCAL" -force -ea SilentlyContinue
# 		New-Item "HKLM:\SOFTWARE\Wow6432Node\Micros\UserData" -force -ea SilentlyContinue
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL' -Name 'CALVersion' -Value '3.1.4.149' -PropertyType String -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL' -Name 'HwConfigured' -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL' -Name 'SHIDSilent' -Value '' -PropertyType String -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'ActiveHost' -Value $Using:RES_Name -PropertyType String -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'ActiveHostIpAddress' -Value $Using:RES_IP -PropertyType String -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'AutoStartApp' -Value '' -PropertyType String -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'AutoStartAppOld' -Value '' -PropertyType String -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'CALEnabled' -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'DeviceId' -Value $Using:CSV_VM_Name -PropertyType String -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'HostDiscoveryPort' -Value 7301 -PropertyType DWord -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'HostDiscoveryPort2' -Value 7302 -PropertyType DWord -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'HostPort' -Value 7300 -PropertyType DWord -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'POSFlags' -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'POSFlagsOld' -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'POSType' -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'PersistentStore' -Value 'C:\Windows\SysWOW64' -PropertyType String -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'PingOn' -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'PingTime' -Value 540000 -PropertyType DWord -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'ProductType' -Value 'WIN32RES' -PropertyType String -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'SecureHostPort' -Value 7303 -PropertyType DWord -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'ServiceHostId' -Value '' -PropertyType String -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'SocketPersistence' -Value 20 -PropertyType DWord -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'WSId' -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Scripts\McrsCAL' -Name 'Version' -Value 50398357 -PropertyType DWord -Force -ea SilentlyContinue;
# 		New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Scripts\McrsHAL' -Name 'Version' -Value 50398349 -PropertyType DWord -Force -ea SilentlyContinue;
# 		[System.Environment]::SetEnvironmentVariable('AppDataDir',$null,[System.EnvironmentVariableTarget]::Machine)		
# 		[System.Environment]::SetEnvironmentVariable('DocDir',$null,[System.EnvironmentVariableTarget]::Machine)
# 		[System.Environment]::SetEnvironmentVariable('DocDrive',$null,[System.EnvironmentVariableTarget]::Machine)
# 		[System.Environment]::SetEnvironmentVariable('MICROSDrive',$null,[System.EnvironmentVariableTarget]::Machine)
# 		[System.Environment]::SetEnvironmentVariable('MICROS_Current_Installation',$null,[System.EnvironmentVariableTarget]::Machine)
# 		[System.Environment]::SetEnvironmentVariable('REGEDITPATH',$null,[System.EnvironmentVariableTarget]::Machine)
# 		[System.Environment]::SetEnvironmentVariable('ServerDataDrive',$null,[System.EnvironmentVariableTarget]::Machine)
# 		[System.Environment]::SetEnvironmentVariable('ServerName',$null,[System.EnvironmentVariableTarget]::Machine)
# 		[System.Environment]::SetEnvironmentVariable('WinSys(x86)',$null,[System.EnvironmentVariableTarget]::Machine)
# 		[System.Environment]::SetEnvironmentVariable('AppRoot','C:',[System.EnvironmentVariableTarget]::Machine)
# 		Rename-Computer -NewName $Using:CSV_VM_Name -Force -ea SilentlyContinue
# 		Write-Host $Using:CSV_VM_Name Rebooting... -ForegroundColor Green -BackgroundColor Black
# 		Restart-Computer -Force
# 	}
# 	Remove-PSSession $s
# 	$output = & ({Write-Output "$CSV_VM_Name complete"}) | Out-String; $ConOut.AppendText($output)
# 	$elapsedTime = $(get-date) - $StartTime
# 	$totalTime = "{0:HH:mm:ss}" -f ([datetime]$elapsedTime.Ticks)
# 	$output = & ({Write-Output "Elapsed Time: $totalTime"}) | Out-String; $ConOut.AppendText($output)
# }	
# $elapsedTime = $(get-date) - $StartTime 
# $totalTime = "{0:HH:mm:ss}" -f ([datetime]$elapsedTime.Ticks)
# Write-Host "Completed AutoCAL on WKS VMs - TOTAL TIME: $totalTime"


#=================================================================================================================


# Write-Output "Changing VM network adapter to $script:switch"
				
# Get-VM * | sort | foreach {
# 	$Name = (Get-VM $_.VMName).Name
# 	$output = & ({Write-Output "Moving $Name to $script:switch..."}) | Out-String; $ConOut.AppendText($output)
# 	Get-VM $Name | Get-VMNetworkAdapter | Connect-VMNetworkAdapter -SwitchName "$script:switch"
# }

#=================================================================================================================

# $TZ = $null

# $TZform = New-Object System.Windows.Forms.Form
# $TZform.Text = $ScriptName
# $TZform.Size = New-Object System.Drawing.Size(300,200)
# $TZform.StartPosition = 'CenterScreen'

# $TZokb = New-Object System.Windows.Forms.Button
# $TZokb.Location = New-Object System.Drawing.Point(65,130)
# $TZokb.Size = New-Object System.Drawing.Size(75,25)
# $TZokb.Text = 'OK'
# $TZokb.DialogResult = [System.Windows.Forms.DialogResult]::OK
# $TZform.AcceptButton = $TZokb
# $TZform.Controls.Add($TZokb)

# $TZlist = New-Object System.Windows.Forms.ListBox
# $TZlist.Location = New-Object System.Drawing.Point(10,40)
# $TZlist.Size = New-Object System.Drawing.Size(255,20)
# $TZlist.Height = 75
# $TZlist.BorderStyle = "None"
# $TZlist.Font = $H3 
# $TZlist.Items.Add('Pacific Standard Time') | Out-Null
# $TZlist.Items.Add('Mountain Standard Time') | Out-Null
# $TZlist.Items.Add('Central Standard Time') | Out-Null
# $TZlist.Items.Add('Eastern Standard Time') | Out-Null
# $TZform.Controls.Add($TZlist)

# $TZcb = New-Object System.Windows.Forms.Button
# $TZcb.Location = New-Object System.Drawing.Point(150,130)
# $TZcb.Size = New-Object System.Drawing.Size(75,25)
# $TZcb.Text = 'Cancel'
# $TZcb.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
# $TZform.CancelButton = $TZcb
# $TZform.Controls.Add($TZcb)

# $TZlb1 = New-Object System.Windows.Forms.Label
# $TZlb1.Location = New-Object System.Drawing.Point(10,15)
# $TZlb1.Size = New-Object System.Drawing.Size(240,20)
# $TZlb1.Text = 'Timezone:'
# $TZform.Controls.Add($TZlb1)

# $TZform.Topmost = $true
# $TZrs = $TZform.ShowDialog()
# 	if ($TZrs -eq [System.Windows.Forms.DialogResult]::OK) {
# 		$script:TZ = $TZlist.SelectedItem
# 	}

# If ($TZ -eq $null) {Write-Output "Timezone invalid"}

# If ($TZ -eq $null) {return}

# Write-Host "Setting Timezone..."

# Get-VM * | Where-Object {$_.VMName -notlike "*wks*"} | sort | foreach {
# 	$Name = (Get-VM $_.VMName).Name
# 	Write-Host "Setting Timezone on $Name..."
# 	$s = New-PSSession -VMName $Name -Credential $adminacc
# 	Invoke-Command -Session $s -ScriptBlock {Set-TimeZone $Using:TZ}
# 	Remove-PSSession $s
# 	}
	
# Get-VM * | Where-Object {$_.VMName -like "*wks*"} | sort | foreach {
# 	$Name = (Get-VM $_.VMName).Name
# 	Write-Host "Setting Timezone on $Name..."
# 	$username = ($Name + '\user')
# 	$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username,$userpassword
# 	$s = New-PSSession -VMName $Name -Credential $creds
# 	Invoke-Command -Session $s -ScriptBlock {Set-TimeZone $Using:TZ}
# 	Remove-PSSession $s
# 	}				
# Write-Host "Setting Timezone on Hyper-V Host..."
# Set-TimeZone $TZ