# Create PWKS and CAL them

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
