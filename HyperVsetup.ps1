# HYPER-V SETUP
# Jeff Funk 2021

#######################################

$RES_OS_VHDX    = "RESxxxx_OS.vhdx"
$RES_DATA_VHDX  = "RESxxxx_DATA.vhdx"
$LTSC_VHDX      = "LTSC_2019.vhdx"
$WKS_VHDX       = "PWKS-PLATINUM.vhdx"

#######################################

$ScriptName     = "Hyper-V Setup"
$ScriptVer      = "3.0"

#######################################

Split-Path -Parent $Script:MyInvocation.MyCommand.Path
Add-Type -AssemblyName PresentationCore,PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]'Administrator')) {
	Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
	Exit
}

# GET CONSOLE PROCESS
$sig = '
[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
[DllImport("user32.dll")] public static extern int SetForegroundWindow(IntPtr hwnd);
[DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();'
$type = Add-Type -MemberDefinition $sig -Name WindowAPI -PassThru

# GET CREDENTIALS
function Set-CustomCreds {
	if (Test-Path ".\creds_domain.xml") {$script:domainacc = Import-CliXml -Path ".\creds_domain.xml"}
	if (Test-Path ".\creds_admin.xml") {$script:adminacc = Import-CliXml -Path ".\creds_admin.xml"}
	Write-Host "DomainAcc" $domainacc
	Write-Host "AdminAcc" $adminacc
	If ($script:domainacc -eq $null) {$script:domainacc = $host.ui.PromptForCredential($ScriptName, "Please enter JOEY Domain Admin credentials", "JOEY\admin$", ""); $domainacc | Export-CliXml -Path ".\creds_domain.xml"}
	If ($script:adminacc -eq $null) {$script:adminacc = $host.ui.PromptForCredential($ScriptName, "Please enter local Administrator credentials", "Administrator", ""); $adminacc | Export-CliXml -Path ".\creds_admin.xml"}
}
$script:userpassword = ConvertTo-SecureString -String 'user' -AsPlainText -Force

# GET SWITCH
function Set-SwitchNameHV {
	If ($script:switch -eq $null -OR $script:switch -notlike '*[a-z0-9]*') {
		
		Add-Type -AssemblyName System.Windows.Forms
		Add-Type -AssemblyName System.Drawing
		$promptform = New-Object System.Windows.Forms.Form
		$promptform.Text = $ScriptName
		$promptform.Size = New-Object System.Drawing.Size(300,200)
		$promptform.StartPosition = 'CenterScreen'
		$okb = New-Object System.Windows.Forms.Button
		$okb.Location = New-Object System.Drawing.Point(65,130)
		$okb.Size = New-Object System.Drawing.Size(75,25)
		$okb.Text = 'OK'
		$okb.DialogResult = [System.Windows.Forms.DialogResult]::OK
		$promptform.AcceptButton = $okb
		$promptform.Controls.Add($okb)
		$cb = New-Object System.Windows.Forms.Button
		$cb.Location = New-Object System.Drawing.Point(150,130)
		$cb.Size = New-Object System.Drawing.Size(75,25)
		$cb.Text = 'Cancel'
		$cb.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
		$promptform.CancelButton = $cb
		$promptform.Controls.Add($cb)
		$lb = New-Object System.Windows.Forms.Label
		$lb.Location = New-Object System.Drawing.Point(20,50)
		$lb.Size = New-Object System.Drawing.Size(240,20)
		$lb.Text = 'Please enter the name of the Hyper-V Switch:'
		$promptform.Controls.Add($lb)
		$tb = New-Object System.Windows.Forms.TextBox
		$tb.Location = New-Object System.Drawing.Point(20,80)
		$tb.Size = New-Object System.Drawing.Size(240,20)
		$promptform.Controls.Add($tb)
		$promptform.Topmost = $true
		$promptform.Add_Shown({$tb.Select()})
		$rs = $promptform.ShowDialog()
		if ($rs -eq [System.Windows.Forms.DialogResult]::OK) {
			$script:switch = $tb.Text
		}
	}
	if ($script:switch -notlike '*[a-z0-9]*') {$script:switch = "<UNDEFINED>"}
	if ($script:switch -eq $null) {$script:switch = "<UNDEFINED>"}
	Write-Host "Switch name: $script:switch" -ForegroundColor Black -BackgroundColor Yellow
	$Form.Controls.remove($Info_Switch); $Info_Switch.Text = "Switch: `n$script:switch"; $form.Refresh()
}

# GET VM INFO
$csv = Import-Csv VMinfo.csv | Where-Object {$_.VM_Name -like "RES*"}
$csv | foreach {
	$RES_Name      = $csv.VM_Name
	$RES_IP        = $csv.IP_Address
	$RES_SM        = $csv.Subnet_Mask
	$RES_GW        = $csv.Gateway
	$RES_DNS1      = $csv.DNS1
	$RES_DNS2      = $csv.DNS2
	$RES_VLAN      = $csv.VLAN
	
	$mask = [ipaddress]$RES_SM
	$binary = [convert]::ToString($mask.Address, 2)
	$mask_length = ($binary -replace 0,$null).Length
	$RES_cidr = '{0}' -f $mask_length
}
		
$csv = Import-Csv VMinfo.csv | Where-Object {$_.VM_Name -like "CSK*"}
$csv | foreach {
	$CSK_Name      = $csv.VM_Name
	$CSK_IP        = $csv.IP_Address
	$CSK_SM        = $csv.Subnet_Mask
	$CSK_GW        = $csv.Gateway
	$CSK_DNS1      = $csv.DNS1
	$CSK_DNS2      = $csv.DNS2
	$CSK_VLAN      = $csv.VLAN
	$CSK_MAC       = $csv.MAC
	
	$CSK_MAC = $CSK_MAC -replace '[^a-z0-9]', ''
		if ($CSK_MAC -notlike '*[a-z0-9]*') {
			$MACappend = (Get-Random -Minimum 0 -Maximum 99999).ToString('000000')
			$CSK_MAC = '00155D'+$MACappend
		}
		
	$mask = [ipaddress]$CSK_SM
	$binary = [convert]::ToString($mask.Address, 2)
	$mask_length = ($binary -replace 0,$null).Length
	$CSK_cidr = '{0}' -f $mask_length
}
		
$csv = Import-Csv VMinfo.csv | Where-Object {$_.VM_Name -like "HST*"}
$csv | foreach {
	$HST_Name      = $csv.VM_Name
	$HST_IP        = $csv.IP_Address
	$HST_SM        = $csv.Subnet_Mask
	$HST_GW        = $csv.Gateway
	$HST_DNS1      = $csv.DNS1
	$HST_DNS2      = $csv.DNS2
	$HST_VLAN      = $csv.VLAN
	$HST_MAC       = $csv.MAC
	
	$HST_MAC = $HST_MAC -replace '[^a-z0-9]', ''
		if ($HST_MAC -notlike '*[a-z0-9]*') {
			$MACappend = (Get-Random -Minimum 0 -Maximum 99999).ToString('000000')
			$HST_MAC = '00155D'+$MACappend
		}
	
	$mask = [ipaddress]$HST_SM
	$binary = [convert]::ToString($mask.Address, 2)
	$mask_length = ($binary -replace 0,$null).Length
	$HST_cidr = '{0}' -f $mask_length
}

# RESIZE CONSOLE
$Height = 30
$Width = 60
$console = $host.ui.rawui
$ConBuffer  = $console.BufferSize
$ConSize = $console.WindowSize
$currWidth = $ConSize.Width
$currHeight = $ConSize.Height
if ($Height -gt $host.UI.RawUI.MaxPhysicalWindowSize.Height) {$Height = $host.UI.RawUI.MaxPhysicalWindowSize.Height}
if ($Width -gt $host.UI.RawUI.MaxPhysicalWindowSize.Width) {$Width = $host.UI.RawUI.MaxPhysicalWindowSize.Width}
If ($ConBuffer.Width -gt $Width ) {$currWidth = $Width}
If ($ConBuffer.Height -gt $Height ) {$currHeight = $Height}
$host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.size($currWidth,$currHeight)
$host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.size($Width,2000)
$host.UI.RawUI.WindowSize = New-Object System.Management.Automation.Host.size($Width,$Height)
clear

# FORM STYLE
$Form                         = New-Object system.Windows.Forms.Form
$Form.text                    = $ScriptName + " " + $ScriptVer
$Form.StartPosition           = "CenterScreen"
$Form.TopMost                 = $false
$Form.BackColor               = [System.Drawing.ColorTranslator]::FromHtml("#353535")
$Form.ForeColor               = [System.Drawing.Color]::White
#$Form.AutoScaleDimensions     = '192,192'
$Form.AutoScaleMode           = "Dpi"
$Form.AutoSize                = $True
$Form.ClientSize              = '1024, 768'
$Form.FormBorderStyle         = 'FixedSingle'
$Form.Width                   = $objImage.Width
$Form.Height                  = $objImage.Height
$Form.MaximizeBox             = $False
$Form.MinimizeBox             = $False
$Form.ControlBox              = $False
$Form.BringToFront()

# ICON
$iconBase64                   = 'AAABAAEAMDAAAAAAAACoJQAAFgAAACgAAAAwAAAAYAAAAAEAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAvWgFYL1pBt8dnMffHZnJUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA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Vv8Otlb/DrZW/w62Vv8YqU//IZ9K7wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAETu9/wg6yv8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8XqlD/IJ9K/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEDq9/wg7y/8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Xq1D/IJ9K/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEDu+/wk7y/8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Xq1D/IKBL/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEDu//wg7y/8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Xq1H/H6BL/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADzu//wg7zP8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Xq1H/H6FL/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADzvA/wg7zP8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Wq1H/H6FL/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADzvB/wg7zP8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8WrFH/H6JL/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADjvB/wg7zP8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8WrFH/HqJL/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADjvC/wc7zf8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8BO9j/ATvY/wE72P8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8Otlb/DrZW/w62Vv8WrFH/HqJM/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADTvD/wc6zf8BO9j/ATvY/wE7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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///////8AAP///////wAA////////vd////w////U////8A///9j////gB///2P///4AB///Y////AAD//9j///4AAH//2P//+AAAH//Y///wAAAP/9j//+AAAAf/2P//wAAAA//Y//+AAAAB/9j//4AAAAH/2P//AAAAAP/Y//4AAAAAf1b//gAAAAB/Vv/8AAAAAD9W//wAAAAAP1b/+AAAAAAfVv/4AAAAAB9W//AAAAAAD1b/8AAAAAAPVv/wAAAAAA9W//AAAAAAD1b/8AAAAAAPVv/wAAAAAA9W//AAAAAAD0r/8AAAAAAPAADwAAAAAA8AAPAAAAAADwAA8AAAAAAPAADwAAAAAA8AAPAAAAAAD73/8AAAAAAP2P/wAAAAAA/Y//AAAAAAD9j/8AAAAAAP2P/wAAAAAA/Y//4AAAAAf9j///gAAB//2P///gAAf//Y////gAH//9j////gB///2P/////////Y/////////9j/////////2P8='
$iconBytes                       = [Convert]::FromBase64String($iconBase64)
$stream                          = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
$stream.Write($iconBytes, 0, $iconBytes.Length)
$Form.Icon                       = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())
 
# FONT STYLE
$H1                              = New-Object System.Drawing.Font('Arial',24,[System.Drawing.FontStyle]::Bold)
$H2                              = New-Object System.Drawing.Font('Arial',14)
$H3                              = New-Object System.Drawing.Font('Arial',8)
$ConsoleFont                     = New-Object System.Drawing.Font('Consolas',12,[System.Drawing.FontStyle]::Bold)
$FunctionFont                    = New-Object System.Drawing.Font('Arial',20,[System.Drawing.FontStyle]::Bold)
$InfoFont                        = New-Object System.Drawing.Font('Arial',12,[System.Drawing.FontStyle]::Bold)

# HORIZONTAL RULE
$HR                              = New-Object system.Windows.Forms.Label
$HR.text                         = ""
$HR.AutoSize                     = $false
$HR.width                        = 1040
$HR.height                       = 2
$HR.location                     = New-Object System.Drawing.Point(15,50)
$HR.BorderStyle                  = "Fixed3D"

# COLUMN 1
$Panel1                          = New-Object system.Windows.Forms.Panel
$Panel1.height                   = 640
$Panel1.width                    = 250
$Panel1.location                 = New-Object System.Drawing.Point(10,50)
$Panel1.AutoSize                 = $true

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "CREATE"
$Label1.AutoSize                 = $true
$Label1.width                    = 200
$Label1.height                   = 25
$Label1.location                 = New-Object System.Drawing.Point(10,10)
$Label1.Font                     = $H1
$Label1.ForeColor                = [System.Drawing.ColorTranslator]::FromHtml("#56B60E")

$RES_VM_Create                   = New-Object system.Windows.Forms.Button
$RES_VM_Create.text              = "Create RES VM"
$RES_VM_Create.width             = 200
$RES_VM_Create.height            = 100
$RES_VM_Create.location          = New-Object System.Drawing.Point(5,10)
$RES_VM_Create.Font              = $H2

$CSK_VM_Create                   = New-Object system.Windows.Forms.Button
$CSK_VM_Create.text              = "Create CSK VM"
$CSK_VM_Create.width             = 200
$CSK_VM_Create.height            = 100
$CSK_VM_Create.location          = New-Object System.Drawing.Point(5,120)
$CSK_VM_Create.Font              = $H2

$HST_VM_Create                   = New-Object system.Windows.Forms.Button
$HST_VM_Create.text              = "Create HST VM"
$HST_VM_Create.width             = 200
$HST_VM_Create.height            = 100
$HST_VM_Create.location          = New-Object System.Drawing.Point(5,230)
$HST_VM_Create.Font              = $H2

$WKS_VM_Create                   = New-Object system.Windows.Forms.Button
$WKS_VM_Create.text              = "Create WKS VMs"
$WKS_VM_Create.width             = 200
$WKS_VM_Create.height            = 100
$WKS_VM_Create.location          = New-Object System.Drawing.Point(5,340)
$WKS_VM_Create.Font              = $H2

# COLUMN 2
$Panel2                          = New-Object system.Windows.Forms.Panel
$Panel2.height                   = 640
$Panel2.width                    = 250
$Panel2.location                 = New-Object System.Drawing.Point(280,50)

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "CONFIG"
$Label2.AutoSize                 = $true
$Label2.width                    = 230
$Label2.height                   = 25
$Label2.location                 = New-Object System.Drawing.Point(280,10)
$Label2.Font                     = $H1
$Label2.ForeColor                = [System.Drawing.ColorTranslator]::FromHtml("#0078D4")

$AutoCAL                         = New-Object system.Windows.Forms.Button
$AutoCAL.text                    = "Auto CAL"
$AutoCAL.width                   = 200
$AutoCAL.height                  = 100
$AutoCAL.location                = New-Object System.Drawing.Point(5,10)
$AutoCAL.Font                    = $H2

$Enable_FPS                      = New-Object system.Windows.Forms.Button
$Enable_FPS.text                 = "Enable File && Print Sharing"
$Enable_FPS.width                = 200
$Enable_FPS.height               = 100
$Enable_FPS.location             = New-Object System.Drawing.Point(5,120)
$Enable_FPS.Font                 = $H2

$InstallCW                       = New-Object system.Windows.Forms.Button
$InstallCW.text                  = "Install CW"
$InstallCW.width                 = 200
$InstallCW.height                = 100
$InstallCW.location              = New-Object System.Drawing.Point(5,230)
$InstallCW.Font                  = $H2

$DomainJoin                      = New-Object system.Windows.Forms.Button
$DomainJoin.text                 = "Add to JOEY Domain"
$DomainJoin.width                = 200
$DomainJoin.height               = 100
$DomainJoin.location             = New-Object System.Drawing.Point(5,340)
$DomainJoin.Font                 = $H2

# COLUMN 3
$Panel3                          = New-Object system.Windows.Forms.Panel
$Panel3.height                   = 640
$Panel3.width                    = 250
$Panel3.location                 = New-Object System.Drawing.Point(550,50)

$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "TOOLS"
$Label3.AutoSize                 = $true
$Label3.width                    = 230
$Label3.height                   = 25
$Label3.location                 = New-Object System.Drawing.Point(550,10)
$Label3.Font                     = $H1
$Label3.ForeColor                = [System.Drawing.ColorTranslator]::FromHtml("#FFB900")

$CaptureCreds                    = New-Object system.Windows.Forms.Button
$CaptureCreds.text               = "Capture Domain && Admin credentials"
$CaptureCreds.width              = 200
$CaptureCreds.height             = 100
$CaptureCreds.location           = New-Object System.Drawing.Point(5,10)
$CaptureCreds.Font               = $H2

$ChangeSwitch                    = New-Object system.Windows.Forms.Button
$ChangeSwitch.text               = "Change Switch"
$ChangeSwitch.width              = 200
$ChangeSwitch.height             = 100
$ChangeSwitch.location           = New-Object System.Drawing.Point(5,120)
$ChangeSwitch.Font               = $H2

$DataPull                        = New-Object system.Windows.Forms.Button
$DataPull.text                   = "Data Pull"
$DataPull.width                  = 200
$DataPull.height                 = 100
$DataPull.location               = New-Object System.Drawing.Point(5,230)
$DataPull.Font                   = $H2

$Timezone                        = New-Object system.Windows.Forms.Button
$Timezone.text                   = "Timezone"
$Timezone.width                  = 200
$Timezone.height                 = 100
$Timezone.location               = New-Object System.Drawing.Point(5,340)
$Timezone.Font                   = $H2

# COLUMN 4
$Panel4                          = New-Object system.Windows.Forms.Panel
$Panel4.height                   = 640
$Panel4.width                    = 250
$Panel4.location                 = New-Object System.Drawing.Point(820,50)

$Label4                          = New-Object system.Windows.Forms.Label
$Label4.text                     = "INFO"
$Label4.AutoSize                 = $true
$Label4.width                    = 230
$Label4.height                   = 25
$Label4.location                 = New-Object System.Drawing.Point(820,10)
$Label4.Font                     = $H1
$Label4.ForeColor                = [System.Drawing.ColorTranslator]::FromHtml("#CC3232")

$Info_RES                        = New-Object system.Windows.Forms.Label
If ($RES_Name -ne $null) {$Info_RES.text = "$RES_Name `nIP: $RES_IP `nSM: $RES_SM `nGW: $RES_GW `nDNS1: $RES_DNS1 `nDNS2: $RES_DNS2"} else {$Info_RES.text = ""}
$Info_RES.AutoSize               = $true
$Info_RES.width                  = 230
$Info_RES.height                 = 25
$Info_RES.location               = New-Object System.Drawing.Point(10,10)
$Info_RES.Font                   = $InfoFont
$Info_RES.ForeColor              = [System.Drawing.ColorTranslator]::FromHtml("#A7A7A7")

$HR1                              = New-Object system.Windows.Forms.Label
$HR1.text                         = ""
$HR1.AutoSize                     = $false
$HR1.width                        = 228
$HR1.height                       = 2
$HR1.location                     = New-Object System.Drawing.Point(10,135)
$HR1.BorderStyle                  = "Fixed3D"

$Info_CSK                        = New-Object system.Windows.Forms.Label
If ($CSK_Name -ne $null) {$Info_CSK.text = "$CSK_Name `nIP: $CSK_IP `nSM: $CSK_SM `nGW: $CSK_GW `nDNS1: $CSK_DNS1 `nDNS2: $CSK_DNS2"} else {$Info_CSK.text = ""}
$Info_CSK.AutoSize               = $true
$Info_CSK.width                  = 230
$Info_CSK.height                 = 25
$Info_CSK.location               = New-Object System.Drawing.Point(10,145)
$Info_CSK.Font                   = $InfoFont
$Info_CSK.ForeColor              = [System.Drawing.ColorTranslator]::FromHtml("#A7A7A7")

$HR2                              = New-Object system.Windows.Forms.Label
$HR2.text                         = ""
$HR2.AutoSize                     = $false
$HR2.width                        = 228
$HR2.height                       = 2
$HR2.location                     = New-Object System.Drawing.Point(10,270)
$HR2.BorderStyle                  = "Fixed3D"

$Info_HST                        = New-Object system.Windows.Forms.Label
If ($HST_Name -ne $null) {$Info_HST.text = "$HST_Name `nIP: $HST_IP `nSM: $HST_SM `nGW: $HST_GW `nDNS1: $HST_DNS1 `nDNS2: $HST_DNS2"} else {$Info_HST.text = ""}
$Info_HST.AutoSize               = $true
$Info_HST.width                  = 230
$Info_HST.height                 = 25
$Info_HST.location               = New-Object System.Drawing.Point(10,280)
$Info_HST.Font                   = $InfoFont 
$Info_HST.ForeColor              = [System.Drawing.ColorTranslator]::FromHtml("#A7A7A7")

$HR3                              = New-Object system.Windows.Forms.Label
$HR3.text                         = ""
$HR3.AutoSize                     = $false
$HR3.width                        = 228
$HR3.height                       = 2
$HR3.location                     = New-Object System.Drawing.Point(10,405)
$HR3.BorderStyle                  = "Fixed3D"

$Info_Switch                     = New-Object system.Windows.Forms.Label
If ($script:switch -ne $null) {$Info_Switch.text = "Switch: `n$script:switchdisplayname"} else {$Info_Switch.text = "Switch: `n<UNDEFINED>"}
$Info_Switch.AutoSize            = $true
$Info_Switch.width               = 230
$Info_Switch.height              = 25
$Info_Switch.location            = New-Object System.Drawing.Point(10,415)
$Info_Switch.Font                = $InfoFont 
$Info_Switch.ForeColor           = [System.Drawing.ColorTranslator]::FromHtml("#A7A7A7")

$ClearConOut                     = New-Object system.Windows.Forms.Button
$ClearConOut.text                = "Clear"
$ClearConOut.width               = 100
$ClearConOut.height              = 50
$ClearConOut.location            = New-Object System.Drawing.Point(20,500)
$ClearConOut.Font                = $ConsoleFont
$ClearConOut.ForeColor           = [System.Drawing.Color]::White
$ClearConOut.BackColor           = [System.Drawing.ColorTranslator]::FromHtml("#012456")

$ShowConsole                     = New-Object system.Windows.Forms.Button
$ShowConsole.text                = "Show Console"
$ShowConsole.width               = 100
$ShowConsole.height              = 50
$ShowConsole.location            = New-Object System.Drawing.Point(120,500)
$ShowConsole.Font                = $ConsoleFont
$ShowConsole.ForeColor           = [System.Drawing.Color]::White
$ShowConsole.BackColor           = [System.Drawing.ColorTranslator]::FromHtml("#012456")

$Exit                            = New-Object system.Windows.Forms.Button
$Exit.text                       = "EXIT"
$Exit.width                      = 200
$Exit.height                     = 50
$Exit.location                   = New-Object System.Drawing.Point(20,570)
$Exit.Font                       = $FunctionFont
$Exit.ForeColor                  = [System.Drawing.Color]::White
$Exit.BackColor                  = [System.Drawing.ColorTranslator]::FromHtml("#CC3232")

# CONSOLE OUTPUT
$ConOut                            = New-Object system.Windows.Forms.TextBox 
$ConOut.text                       = ""
$ConOut.width                      = 800
$ConOut.height                     = 170
$ConOut.location                   = New-Object System.Drawing.Point(15,500)
$ConOut.Font                       = New-Object System.Drawing.Font('Consolas',10)
$ConOut.ForeColor                  = [System.Drawing.Color]::White
$ConOut.BackColor                  = [System.Drawing.ColorTranslator]::FromHtml("#012456")
$ConOut.Multiline                  = $true
$ConOut.AcceptsReturn              = $false
$ConOut.AcceptsTab                 = $false
$ConOut.WordWrap                   = $false
$ConOut.ScrollBars                 = "Vertical"

# ADD CONTROLS
$Form.controls.AddRange(@($ConOut,$HR,$Panel1,$Panel2,$Panel3,$Panel4,$Label1,$Label2,$Label3,$Label4))
$Panel1.controls.AddRange(@($RES_VM_Create,$CSK_VM_Create,$HST_VM_Create,$WKS_VM_Create))
$Panel2.controls.AddRange(@($AutoCAL,$Enable_FPS,$InstallCW,$DomainJoin))
$Panel3.controls.AddRange(@($CaptureCreds,$ChangeSwitch,$DataPull,$Timezone))
$Panel4.controls.AddRange(@($HR1,$HR2,$HR3,$Info_RES,$Info_CSK,$Info_HST,$Info_Switch,$ClearConOut,$ShowConsole,$Exit))

# COLUMN 1 ACTIONS
$RES_VM_Create.Add_Click({
$ButtonType = [System.Windows.MessageBoxButton]::YesNo
$MessageIcon = [System.Windows.MessageBoxImage]::Information
$MessageBody = "Create RES VM?"
$MessageTitle = "Confirm Operation"
$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)

	If ($Result -eq 'Yes') {
		Set-CustomCreds
			If ($adminacc -eq $null) {Write-Output "Credentials not captured"; return}
		Set-SwitchNameHV
		If ($script:switch -eq "<UNDEFINED>") {
			$output = & ({Write-Output "Switch undefined - aborting"}) 2>&1 | Out-String; $ConOut.AppendText($output)
			return
			}
		$StartTime = $(Get-Date)
		$output = & ({Write-Output "Creating $RES_Name - Start Time: $StartTime"}) | Out-String; $ConOut.AppendText($output)
		
		$output = & ({
			
			$csv = Import-Csv VMinfo.csv | Where-Object {$_.VM_Name -eq $RES_Name}
			$csv | ForEach-Object {
			$name = $_.VM_Name    	
			$Final_RES_OS = "D:\Hyper-V\Virtual Hard Disks\"+$RES_Name+"_OS.vhdx"
			$Final_RES_DATA = "D:\Hyper-V\Virtual Hard Disks\"+$RES_Name+"_DATA.vhdx"
			Convert-VHD -DestinationPath $Final_RES_OS -Path .\VHDs\$RES_OS_VHDX -VHDType Fixed
			Convert-VHD -DestinationPath $Final_RES_DATA -Path .\VHDs\$RES_DATA_VHDX -VHDType Fixed
			New-VM -Name $name -MemoryStartupBytes 8GB -VHDPath $Final_RES_OS -Path "D:\Hyper-V\Virtual Machines" -Switch "$script:switch" -Generation 2
			Add-VMHardDiskDrive -VMName $name -Path $Final_RES_DATA
			Set-VMProcessor -VMName $name -Count 4 -RelativeWeight 100
			Set-VMNetworkAdapterVlan -VMName $name -Access -VlanId $RES_VLAN 
			Start-VM $name
			
			Sleep 60
			
				$s = New-PSSession -VMName $name -Credential $adminacc
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
				Remove-PSSession $s
			
			}
		
		}) 2>&1 | Out-String; $ConOut.AppendText($output)
		
		$elapsedTime = $(get-date) - $StartTime 
		$totalTime = "{0:HH:mm:ss}" -f ([datetime]$elapsedTime.Ticks)
		$output = & ({Write-Output "Completed $RES_Name - Total Time: $totalTime"}) | Out-String; $ConOut.AppendText($output)
	}
})

$CSK_VM_Create.Add_Click({
$ButtonType = [System.Windows.MessageBoxButton]::YesNo
$MessageIcon = [System.Windows.MessageBoxImage]::Information
$MessageBody = "Create CSK VM?"
$MessageTitle = "Confirm Operation"
$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)

	If ($Result -eq 'Yes') {
		Set-CustomCreds
			If ($adminacc -eq $null) {Write-Output "Credentials not captured"; return}
		Set-SwitchNameHV
		If ($script:switch -eq "<UNDEFINED>") {
			$output = & ({Write-Output "Switch undefined - aborting"}) | Out-String; $ConOut.AppendText($output)
			return
			}
		$StartTime = $(Get-Date)
		$output = & ({Write-Output "Creating $CSK_Name - Start Time: $StartTime"}) | Out-String; $ConOut.AppendText($output)
		
		$output = & ({
		
			$csv = Import-Csv VMinfo.csv | Where-Object {$_.VM_Name -eq $CSK_Name}
			$csv | ForEach-Object {
			$name = $_.VM_Name    	
			Convert-VHD -DestinationPath "D:\Hyper-V\Virtual Hard Disks\$name.vhdx" -Path .\VHDs\$LTSC_VHDX -VHDType Fixed
			New-VM -Name $name -MemoryStartupBytes 4GB -VHDPath "D:\Hyper-V\Virtual Hard Disks\$name.vhdx" -Path "D:\Hyper-V\Virtual Machines" -Switch "$switch" -Generation 2
			Set-VMProcessor -VMName $name -Count 2 -RelativeWeight 100
			Set-VMNetworkAdapterVlan -VMName $name -Access -VlanId $CSK_VLAN
			Write-Output "Setting MAC Address: $CSK_MAC"
			Set-VMNetworkAdapter -VMName $name -StaticMacAddress $CSK_MAC
			Start-VM $name
			
			Sleep 60
			
				$s = New-PSSession -VMName $name -Credential $adminacc
				Invoke-Command -Session $s -ScriptBlock { 
					Get-NetAdapter | Get-NetIPAddress | Remove-NetIPAddress -Confirm:$false -ea SilentlyContinue
					Get-NetAdapter | Get-NetRoute | Remove-NetRoute -Confirm:$false -ea SilentlyContinue
					Get-NetIPConfiguration | New-NetIPAddress -IPAddress $Using:CSK_IP -DefaultGateway $Using:CSK_GW -PrefixLength $Using:CSK_cidr
					Get-NetIPConfiguration | Set-DnsClientServerAddress -ServerAddresses $Using:CSK_DNS1,$Using:CSK_DNS2
					Rename-Computer -NewName $Using:CSK_Name -Force -ea SilentlyContinue
					Write-Host $Using:CSK_Name Rebooting... -ForegroundColor Green -BackgroundColor Black
					Restart-Computer -Force
					}
				Remove-PSSession $s
				
			}
			
		}) 2>&1 | Out-String; $ConOut.AppendText($output)
		
		$elapsedTime = $(get-date) - $StartTime 
		$totalTime = "{0:HH:mm:ss}" -f ([datetime]$elapsedTime.Ticks)
		$output = & ({Write-Output "Completed $CSK_Name - Total Time: $totalTime"}) | Out-String; $ConOut.AppendText($output)
	}
})

$HST_VM_Create.Add_Click({
$ButtonType = [System.Windows.MessageBoxButton]::YesNo
$MessageIcon = [System.Windows.MessageBoxImage]::Information
$MessageBody = "Create HST VM?"
$MessageTitle = "Confirm Operation"
$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)

	If ($Result -eq 'Yes') {
		Set-CustomCreds
			If ($adminacc -eq $null) {Write-Output "Credentials not captured"; return}
		Set-SwitchNameHV
		If ($script:switch -eq "<UNDEFINED>") {
			$output = & ({Write-Output "Switch undefined - aborting"}) | Out-String; $ConOut.AppendText($output)
			return
			}
		$StartTime = $(Get-Date)
		$output = & ({Write-Output "Creating $HST_Name - Start Time: $StartTime"}) | Out-String; $ConOut.AppendText($output)
		
		$output = & ({
		
			$csv = Import-Csv VMinfo.csv | Where-Object {$_.VM_Name -eq $HST_Name}
			$csv | ForEach-Object {
			$name = $_.VM_Name  
			Convert-VHD -DestinationPath "D:\Hyper-V\Virtual Hard Disks\$name.vhdx" -Path .\VHDs\$LTSC_VHDX -VHDType Fixed
			New-VM -Name $name -MemoryStartupBytes 4GB -VHDPath "D:\Hyper-V\Virtual Hard Disks\$name.vhdx" -Path "D:\Hyper-V\Virtual Machines" -Switch "$switch" -Generation 2
			Set-VMProcessor -VMName $name -Count 2 -RelativeWeight 100
			Set-VMNetworkAdapterVlan -VMName $name -Access -VlanId $HST_VLAN 
			Write-Output "Setting MAC Address: $HST_MAC"
			Set-VMNetworkAdapter -VMName $name -StaticMacAddress $HST_MAC
			Start-VM $name
			
			Sleep 60
			
				$s = New-PSSession -VMName $name -Credential $adminacc
				Invoke-Command -Session $s -ScriptBlock { 
					Get-NetAdapter | Get-NetIPAddress | Remove-NetIPAddress -Confirm:$false -ea SilentlyContinue
					Get-NetAdapter | Get-NetRoute | Remove-NetRoute -Confirm:$false -ea SilentlyContinue
					Get-NetIPConfiguration | New-NetIPAddress -IPAddress $Using:HST_IP -DefaultGateway $Using:HST_GW -PrefixLength $Using:HST_cidr
					Get-NetIPConfiguration | Set-DnsClientServerAddress -ServerAddresses $Using:HST_DNS1,$Using:HST_DNS2
					Rename-Computer -NewName $Using:HST_Name -Force -ea SilentlyContinue
					Write-Host $Using:HST_Name Rebooting... -ForegroundColor Green -BackgroundColor Black
					Restart-Computer -Force
					}
				Remove-PSSession $s
				
			}
			
		}) 2>&1 | Out-String; $ConOut.AppendText($output)
			
		$elapsedTime = $(get-date) - $StartTime 
		$totalTime = "{0:HH:mm:ss}" -f ([datetime]$elapsedTime.Ticks)
		$output = & ({Write-Output "Completed $HST_Name - Total Time: $totalTime"}) | Out-String; $ConOut.AppendText($output)
	}
})

$WKS_VM_Create.Add_Click({
$ButtonType = [System.Windows.MessageBoxButton]::YesNo
$MessageIcon = [System.Windows.MessageBoxImage]::Information
$MessageBody = "Create WKS VMs?"
$MessageTitle = "Confirm Operation"
$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
	
	If ($Result -eq 'Yes') {
		Set-SwitchNameHV
		If ($switch -eq "<UNDEFINED>") {
			$output = & ({Write-Output "Switch undefined - aborting"}) | Out-String; $ConOut.AppendText($output)
			return
			}
		$StartTime = $(Get-Date)
		$output = & ({Write-Output "Creating WKS VMs - Start Time: $StartTime"}) | Out-String; $ConOut.AppendText($output)
		
		$output = & ({
			
			$csv = Import-Csv VMinfo.csv | Where-Object {$_.VM_Name -like "*-*wks*"}
			$csv | ForEach-Object {
			$name = $_.VM_Name    	
			$output = & ({Write-Output "Creating $name..."}) | Out-String; $ConOut.AppendText($output)
			Convert-VHD -DestinationPath "D:\Hyper-V\Virtual Hard Disks\$name.vhdx" -Path .\VHDs\$WKS_VHDX -VHDType Fixed
			New-VM -Name $name -MemoryStartupBytes 2GB -VHDPath "D:\Hyper-V\Virtual Hard Disks\$name.vhdx" -Path "D:\Hyper-V\Virtual Machines" -Switch "$switch" -Generation 2
			Set-VMProcessor -VMName $name -Count 2 -RelativeWeight 125
			Start-VM $name
			}
			
		}) 2>&1 | Out-String; $ConOut.AppendText($output)
		
		$elapsedTime = $(get-date) - $StartTime 
		$totalTime = "{0:HH:mm:ss}" -f ([datetime]$elapsedTime.Ticks)
		$output = & ({Write-Output "Completed WKS VMs - Total Time: $totalTime"}) | Out-String; $ConOut.AppendText($output)
	}
})

# COLUMN 2 ACTIONS
$AutoCAL.Add_Click({
$ButtonType = [System.Windows.MessageBoxButton]::YesNo
$MessageIcon = [System.Windows.MessageBoxImage]::Information
$MessageBody = "This will reset ALL currently running WKS VMs - Do you wish to continue?"
$MessageTitle = "Confirm Operation"
$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)

	If ($Result -eq 'Yes') {
		
		$StartTime = $(Get-Date)
		$output = & ({Write-Output "Running AutoCAL on WKS VMs - Start Time: $StartTime"}) | Out-String; $ConOut.AppendText($output)
		
		$output = & ({
				
			Get-VM *-*wks* | sort | foreach {
				$VMdisplayname = $_.VMName
				$username = ($_.VMName + '\user')
				$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username,$userpassword
				$s = New-PSSession -VMName $_.VMName -Credential $creds
				$CSV_VM_Name = $null
					
					$csv = Import-Csv VMinfo.csv | Where-Object {$_.VM_Name -eq $VMdisplayname}
					$csv | foreach {
					$CSV_VM_Name = $csv.VM_Name
					$CSV_IP_Address = $csv.IP_Address
					$CSV_Subnet_Mask = $csv.Subnet_Mask
					$CSV_Gateway = $csv.Gateway
					$CSV_DNS1 = $csv.DNS1
					$CSV_DNS2 = $csv.DNS2
					$CSV_VLAN = $csv.VLAN
					}
					
				if ($CSV_VM_Name -eq $null) {return}
	
				$output = & ({Write-Output "Setting $CSV_VM_Name to AutoCAL..."}) | Out-String; $ConOut.AppendText($output)
					
				$mask = [ipaddress]$CSV_Subnet_Mask
				$binary = [convert]::ToString($mask.Address, 2)
				$mask_length = ($binary -replace 0,$null).Length
				$cidr = '{0}' -f $mask_length
			
				Set-VMNetworkAdapterVlan -VMName $_.VMName -Access -VlanId $CSV_VLAN 
				
				Invoke-Command -Session $s -ScriptBlock { 
					Invoke-Expression "taskkill /f /t /im DbUpdateServer.exe" -ea SilentlyContinue
					Invoke-Expression "taskkill /f /t /im KDSController.exe" -ea SilentlyContinue
					Invoke-Expression "taskkill /f /t /im MDSHTTPService.exe" -ea SilentlyContinue
					Invoke-Expression "taskkill /f /t /im McrsCal.exe" -ea SilentlyContinue
					Invoke-Expression "taskkill /f /t /im Ops.exe" -ea SilentlyContinue
					Invoke-Expression "taskkill /f /t /im Periphs.exe" -ea SilentlyContinue
					Invoke-Expression "taskkill /f /t /im WaitForHostsFile.exe" -ea SilentlyContinue
					Invoke-Expression "taskkill /f /t /im Win7CALStart.exe" -ea SilentlyContinue
					Get-NetAdapter | Get-NetIPAddress | Remove-NetIPAddress -Confirm:$false -ea SilentlyContinue
					Get-NetAdapter | Get-NetRoute | Remove-NetRoute -Confirm:$false -ea SilentlyContinue
					Get-NetIPConfiguration | New-NetIPAddress -IPAddress $Using:CSV_IP_Address -DefaultGateway $Using:CSV_Gateway -PrefixLength $Using:cidr
					Get-NetIPConfiguration | Set-DnsClientServerAddress -ServerAddresses $Using:CSV_DNS1,$Using:CSV_DNS2
					rm C:\micros -r -fo -ea SilentlyContinue
					rm C:\CALTemp -r -fo -ea SilentlyContinue
					Remove-Item "HKLM:\SOFTWARE\Wow6432Node\Micros" -r -force -ea SilentlyContinue
					New-Item "HKLM:\SOFTWARE\Wow6432Node\Micros" -force -ea SilentlyContinue
					New-Item "HKLM:\SOFTWARE\Wow6432Node\Micros\CAL" -force -ea SilentlyContinue
					New-Item "HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config" -force -ea SilentlyContinue
					New-Item "HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Scripts" -force -ea SilentlyContinue
					New-Item "HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Scripts\McrsCAL" -force -ea SilentlyContinue
					New-Item "HKLM:\SOFTWARE\Wow6432Node\Micros\UserData" -force -ea SilentlyContinue
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL' -Name 'CALVersion' -Value '3.1.4.149' -PropertyType String -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL' -Name 'HwConfigured' -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL' -Name 'SHIDSilent' -Value '' -PropertyType String -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'ActiveHost' -Value $Using:RES_Name -PropertyType String -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'ActiveHostIpAddress' -Value $Using:RES_IP -PropertyType String -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'AutoStartApp' -Value '' -PropertyType String -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'AutoStartAppOld' -Value '' -PropertyType String -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'CALEnabled' -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'DeviceId' -Value $Using:CSV_VM_Name -PropertyType String -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'HostDiscoveryPort' -Value 7301 -PropertyType DWord -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'HostDiscoveryPort2' -Value 7302 -PropertyType DWord -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'HostPort' -Value 7300 -PropertyType DWord -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'POSFlags' -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'POSFlagsOld' -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'POSType' -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'PersistentStore' -Value 'C:\Windows\SysWOW64' -PropertyType String -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'PingOn' -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'PingTime' -Value 540000 -PropertyType DWord -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'ProductType' -Value 'WIN32RES' -PropertyType String -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'SecureHostPort' -Value 7303 -PropertyType DWord -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'ServiceHostId' -Value '' -PropertyType String -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'SocketPersistence' -Value 20 -PropertyType DWord -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Config' -Name 'WSId' -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Scripts\McrsCAL' -Name 'Version' -Value 50398357 -PropertyType DWord -Force -ea SilentlyContinue;
					New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Wow6432Node\Micros\CAL\Scripts\McrsHAL' -Name 'Version' -Value 50398349 -PropertyType DWord -Force -ea SilentlyContinue;
					[System.Environment]::SetEnvironmentVariable('AppDataDir',$null,[System.EnvironmentVariableTarget]::Machine)		
					[System.Environment]::SetEnvironmentVariable('DocDir',$null,[System.EnvironmentVariableTarget]::Machine)
					[System.Environment]::SetEnvironmentVariable('DocDrive',$null,[System.EnvironmentVariableTarget]::Machine)
					[System.Environment]::SetEnvironmentVariable('MICROSDrive',$null,[System.EnvironmentVariableTarget]::Machine)
					[System.Environment]::SetEnvironmentVariable('MICROS_Current_Installation',$null,[System.EnvironmentVariableTarget]::Machine)
					[System.Environment]::SetEnvironmentVariable('REGEDITPATH',$null,[System.EnvironmentVariableTarget]::Machine)
					[System.Environment]::SetEnvironmentVariable('ServerDataDrive',$null,[System.EnvironmentVariableTarget]::Machine)
					[System.Environment]::SetEnvironmentVariable('ServerName',$null,[System.EnvironmentVariableTarget]::Machine)
					[System.Environment]::SetEnvironmentVariable('WinSys(x86)',$null,[System.EnvironmentVariableTarget]::Machine)
					[System.Environment]::SetEnvironmentVariable('AppRoot','C:',[System.EnvironmentVariableTarget]::Machine)
					Rename-Computer -NewName $Using:CSV_VM_Name -Force -ea SilentlyContinue
					Write-Host $Using:CSV_VM_Name Rebooting... -ForegroundColor Green -BackgroundColor Black
					Restart-Computer -Force
				}
				Remove-PSSession $s
				$output = & ({Write-Output "$CSV_VM_Name complete"}) | Out-String; $ConOut.AppendText($output)
				$elapsedTime = $(get-date) - $StartTime
				$totalTime = "{0:HH:mm:ss}" -f ([datetime]$elapsedTime.Ticks)
				$output = & ({Write-Output "Elapsed Time: $totalTime"}) | Out-String; $ConOut.AppendText($output)
			}
		
		}) 2>&1 | Out-String; $ConOut.AppendText($output)
		
		$elapsedTime = $(get-date) - $StartTime 
		$totalTime = "{0:HH:mm:ss}" -f ([datetime]$elapsedTime.Ticks)
		$output = & ({Write-Output "Completed AutoCAL on WKS VMs - TOTAL TIME: $totalTime"}) | Out-String; $ConOut.AppendText($output)

	}  Else {	
		$output = & ({Write-Output "Operation Cancelled"}) | Out-String; $ConOut.AppendText($output)
	} 
})

$Enable_FPS.Add_Click({
$ButtonType = [System.Windows.MessageBoxButton]::YesNo
$MessageIcon = [System.Windows.MessageBoxImage]::Information
$MessageBody = "Enable File & Print Sharing on all VMs?"
$MessageTitle = "Confirm Operation"
$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
	
	If ($Result -eq 'Yes') {
			
		$output = & ({
			Set-CustomCreds
			If ($adminacc -eq $null) {Write-Output "Credentials not captured"; return}
			
			Get-VM * | Where-Object {$_.VMName -notlike "*wks*"} | sort | foreach {
				$Name = (Get-VM $_.VMName).Name
				$output = & ({Write-Output "Enabling File & Print Sharing on $Name..."}) | Out-String; $ConOut.AppendText($output)
				$s = New-PSSession -VMName $Name -Credential $adminacc
				Invoke-Command -Session $s -ScriptBlock { 
					Set-NetFirewallRule -DisplayGroup "File and Printer Sharing" -Enabled True -Profile Any
					}
				Remove-PSSession $s
				}
				
			Get-VM * | Where-Object {$_.VMName -like "*wks*"} | sort | foreach {
				$Name = (Get-VM $_.VMName).Name
				$output = & ({Write-Output "Enabling File & Print Sharing on $Name..."}) | Out-String; $ConOut.AppendText($output)
				$username = ($Name + '\user')
				$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username,$userpassword
				$s = New-PSSession -VMName $Name -Credential $creds
				Invoke-Command -Session $s -ScriptBlock { 
					Set-NetFirewallRule -DisplayGroup "File and Printer Sharing" -Enabled True -Profile Any
					}
				Remove-PSSession $s
				}
				
		}) | Out-String; $ConOut.AppendText($output)
	}

})

$InstallCW.Add_Click({
$ButtonType = [System.Windows.MessageBoxButton]::YesNo
$MessageIcon = [System.Windows.MessageBoxImage]::Question
$MessageBody = "Would you like to install ConnectWise software on all VMs?"
$MessageTitle = "Confirm Operation"
$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)

	If ($Result -eq 'Yes') {
		
		$output = & ({
		Set-CustomCreds
		If ($adminacc -eq $null -OR $domainacc -eq $null) {Write-Output "Credentials not captured"; return}
		}) | Out-String; $ConOut.AppendText($output)
				
		$output = & ({Write-Output "Installing ConnectWise Automate and Control"}) | Out-String; $ConOut.AppendText($output)
		
		$output = & ({New-PSDrive -Name "software" -PSProvider FileSystem -Root "\\joey-file\software$" -Credential $domainacc}) | Out-String; $ConOut.AppendText($output)
		
		$output = & ({
			If ($RES_Name -eq $null) {Write-Output "Server unknown"; return}
			$msi = ($RES_Name -replace "[^0-9]" , '') + ".msi"		
			$CWautomate = (Get-ChildItem -Filter $msi -File -Path "\\joey-file\software$\ConnectWise\#\AUTOMATE").FullName
			$CWcontrol = (Get-ChildItem -Filter $msi -File -Path "\\joey-file\software$\ConnectWise\#\CONTROL").FullName
			Write-Output "$CWcontrol"
			Write-Output "$CWautomate"
			If ($CWautomate -eq $null -OR $CWcontrol -eq $null) {Write-Output "MSI files not found"; return}
			mkdir C:\temp | Out-Null
			Copy-Item -Path $CWcontrol -Destination C:\temp\CWcontrol.msi -Force
			Copy-Item -Path $CWautomate -Destination C:\temp\CWautomate.msi -Force
		}) | Out-String; $ConOut.AppendText($output)
		
		$output = & ({
			Get-VM * | Where-Object {$_.VMName -notlike "*wks*"} | sort | foreach {
				$Name = (Get-VM $_.VMName).Name
				$output = & ({Write-Output "Installing on $Name..."}) | Out-String; $ConOut.AppendText($output)
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
				
			Get-VM * | Where-Object {$_.VMName -like "*wks*"} | sort | foreach {
				$Name = (Get-VM $_.VMName).Name
				$output = & ({Write-Output "Installing on $Name..."}) | Out-String; $ConOut.AppendText($output)
				$username = ($Name + '\user')
				$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username,$userpassword
				$s = New-PSSession -VMName $Name -Credential $creds
				Copy-Item -ToSession $s -Path C:\temp\CWcontrol.msi -Destination C:\
				Invoke-Command -Session $s -ScriptBlock { 
					msiexec /i C:\CWcontrol.msi /quiet /norestart 
					}
				Remove-PSSession $s
				}				
		}) | Out-String; $ConOut.AppendText($output)
		
		$output = & ({net use "\\joey-file\software$" /d /y}) | Out-String; $ConOut.AppendText($output)
	}
})

$DomainJoin.Add_Click({
$ButtonType = [System.Windows.MessageBoxButton]::YesNo
$MessageIcon = [System.Windows.MessageBoxImage]::Information
$MessageBody = "Add ALL VMs to JOEY Domain? - VMs will be restarted"
$MessageTitle = "Confirm Operation"
$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
	
	If ($Result -eq 'Yes') {
		
		$output = & ({
		Set-CustomCreds
		If ($adminacc -eq $null -OR $domainacc -eq $null) {Write-Output "Credentials not captured"; return}
		}) | Out-String; $ConOut.AppendText($output)
		
		$output = & ({Write-Output "Joining to Domain..."}) | Out-String; $ConOut.AppendText($output)
			
		$output = & ({
			Get-VM * | Where-Object {$_.VMName -notlike "*wks*"} | sort | foreach {
				$Name = (Get-VM $_.VMName).Name
				$output = & ({Write-Output "Joining $Name to JOEY.local..."}) | Out-String; $ConOut.AppendText($output)
				$s = New-PSSession -VMName $Name -Credential $adminacc
				Invoke-Command -Session $s -ScriptBlock { 
					Add-Computer -DomainName joey.local -Credential $Using:domainacc -Restart -Force
					}
				Remove-PSSession $s
				}
				
			Get-VM * | Where-Object {$_.VMName -like "*wks*"} | sort | foreach {
				$Name = (Get-VM $_.VMName).Name
				$output = & ({Write-Output "Joining $Name to JOEY.local..."}) | Out-String; $ConOut.AppendText($output)
				$username = ($Name + '\user')
				$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username,$userpassword
				$s = New-PSSession -VMName $Name -Credential $creds
				Invoke-Command -Session $s -ScriptBlock { 
					Add-Computer -DomainName joey.local -Credential $Using:domainacc -Restart -Force
					}
				Remove-PSSession $s
				}				
		}) | Out-String; $ConOut.AppendText($output)
	}

})

# COLUMN 3 ACTIONS
$CaptureCreds.Add_Click({
	rm -force ".\creds_domain.xml" | Out-Null
	rm -force ".\creds_admin.xml" | Out-Null
	$script:domainacc = $null
	$script:adminacc = $null
    Set-CustomCreds
})

$ChangeSwitch.Add_Click({
$script:switch = $null
Set-SwitchNameHV
If ($script:switch -eq "<UNDEFINED>") {Write-Host "Switch not defined"; return}

$ButtonType = [System.Windows.MessageBoxButton]::YesNo
$MessageIcon = [System.Windows.MessageBoxImage]::Question
$MessageBody = "Would you like to change to switch '$script:switch' on all VMs?"
$MessageTitle = "Confirm Operation"
$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)

	If ($Result -eq 'Yes') {
		
		$output = & ({Write-Output "Changing VM network adapter to $script:switch"}) | Out-String; $ConOut.AppendText($output)
		
		$output = & ({
						
			Get-VM * | sort | foreach {
				$Name = (Get-VM $_.VMName).Name
				$output = & ({Write-Output "Moving $Name to $script:switch..."}) | Out-String; $ConOut.AppendText($output)
				Get-VM $Name | Get-VMNetworkAdapter | Connect-VMNetworkAdapter -SwitchName "$script:switch"
			}
			
		}) | Out-String; $ConOut.AppendText($output)
		
		$output = & ({Write-Output "DONE"}) | Out-String; $ConOut.AppendText($output)
	}
})

$DataPull.Add_Click({
	
	$output = & ({
		$DPpromptform = New-Object System.Windows.Forms.Form
		$DPpromptform.Text = $ScriptName
		$DPpromptform.Size = New-Object System.Drawing.Size(300,200)
		$DPpromptform.StartPosition = 'CenterScreen'
		
		$DPokb = New-Object System.Windows.Forms.Button
		$DPokb.Location = New-Object System.Drawing.Point(65,130)
		$DPokb.Size = New-Object System.Drawing.Size(75,25)
		$DPokb.Text = 'OK'
		$DPokb.DialogResult = [System.Windows.Forms.DialogResult]::OK
		$DPpromptform.AcceptButton = $DPokb
		$DPpromptform.Controls.Add($DPokb)
		
		$DPcb = New-Object System.Windows.Forms.Button
		$DPcb.Location = New-Object System.Drawing.Point(150,130)
		$DPcb.Size = New-Object System.Drawing.Size(75,25)
		$DPcb.Text = 'Cancel'
		$DPcb.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
		$DPpromptform.CancelButton = $DPcb
		$DPpromptform.Controls.Add($DPcb)
		
		$DPlb1 = New-Object System.Windows.Forms.Label
		$DPlb1.Location = New-Object System.Drawing.Point(20,30)
		$DPlb1.Size = New-Object System.Drawing.Size(240,20)
		$DPlb1.Text = 'IP address of current RES server:'
		$DPpromptform.Controls.Add($DPlb1)
		
		$DPlb2 = New-Object System.Windows.Forms.Label
		$DPlb2.Location = New-Object System.Drawing.Point(20,80)
		$DPlb2.Size = New-Object System.Drawing.Size(240,20)
		$DPlb2.Text = 'IP address of current CSK server:'
		$DPpromptform.Controls.Add($DPlb2)
		
		$DPtb1 = New-Object System.Windows.Forms.TextBox
		$DPtb1.Location = New-Object System.Drawing.Point(20,50)
		$DPtb1.Size = New-Object System.Drawing.Size(240,20)
		$DPtb1.Text = $RES_IP
		$DPpromptform.Controls.Add($DPtb1)
		
		$DPtb2 = New-Object System.Windows.Forms.TextBox
		$DPtb2.Location = New-Object System.Drawing.Point(20,100)
		$DPtb2.Size = New-Object System.Drawing.Size(240,20)
		$DPtb2.Text = $CSK_IP
		$DPpromptform.Controls.Add($DPtb2)
		
		$DPpromptform.Topmost = $true
		$DPrs = $DPpromptform.ShowDialog()
			if ($DPrs -eq [System.Windows.Forms.DialogResult]::OK) {
				$script:DP_RES_IP = $DPtb1.Text
				$script:DP_CSK_IP = $DPtb2.Text
			}
	}) | Out-String; $ConOut.AppendText($output)

		$output = & ({
		if ($DP_RES_IP -eq $null) {Write-Output "IP Address invalid"; return}
		$TestRES = Test-Path "\\$DP_RES_IP\d$\MICROS"
		if ($TestRES -eq $false) {Write-Output "Network path invalid"; return}
		}) | Out-String; $ConOut.AppendText($output)
		
		$output = & ({
			Write-Output "Pulling data may take some time, please be patient..."
			Write-Output "RES IP: $DP_RES_IP"
			Write-Output "CSK IP: $DP_CSK_IP"
		}) | Out-String; $ConOut.AppendText($output)
		
		$output = & ({
			$script:datestamp = Get-Date -Format "yyyy-MM-dd_HH.mm"
			$script:BackupFolder = "D:\TEMP\"+$RES_Name+"_$datestamp"
			Write-Output "Backing up RES data to $BackupFolder"
		}) | Out-String; $ConOut.AppendText($output)
		
		Start-Sleep 5
		
		New-Item -Type Directory -Force "$BackupFolder\D_drive\MICROS\Res\CAL\Win32\Files\Micros\Res\Pos"
		New-Item -Type Directory -Force "$BackupFolder\D_drive\MICROS\Common\Etc"
		New-Item -Type Directory -Force "$BackupFolder\D_drive\MICROS\Database\Data\Backup\Archive"
		New-Item -Type Directory -Force "$BackupFolder\D_drive\MICROS\ProfessionalServices\StoredValueCard"
		New-Item -Type Directory -Force "$BackupFolder\D_drive\MICROS\Res\CAL\Win32\Files\Micros\Res\Pos"
		New-Item -Type Directory -Force "$BackupFolder\D_drive\MICROS\Res\CAL\WS5\Files\CF\Micros"
		New-Item -Type Directory -Force "$BackupFolder\D_drive\MICROS\Res\CAL\WS5A\Files\CF\Micros"
		New-Item -Type Directory -Force "$BackupFolder\D_drive\MICROS\Res\Pos\Etc"
		New-Item -Type Directory -Force "$BackupFolder\D_drive\MICROS\Res\Pos\Temp"
		New-Item -Type Directory -Force "$BackupFolder\D_drive\MICROS\Res\Pos\Reports\40 Column"
		New-Item -Type Directory -Force "$BackupFolder\C_Drive\Salescash\archive"
		New-Item -Type Directory -Force "$BackupFolder\C_Drive\Salescash\DW\archive"
		
		$output = & ({ Write-Output "Copying Common data..." }) | Out-String; $ConOut.AppendText($output)
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Common\Etc\MDSHosts.xml" -Destination "$BackupFolder\D_drive\MICROS\Common\Etc"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Common\Etc\MDSPrinters.xml" -Destination "$BackupFolder\D_drive\MICROS\Common\Etc"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Common\Cert" -Destination "$BackupFolder\D_drive\MICROS\Common"

		$output = & ({ Write-Output "Copying ProfessionalServices data..." }) | Out-String; $ConOut.AppendText($output)
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\ProfessionalServices\StoredValueCard\svcServer.cfg" -Destination "$BackupFolder\D_drive\MICROS\ProfessionalServices\StoredValueCard"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\ProfessionalServices\StoredValueCard\SVS.svcha.config" -Destination "$BackupFolder\D_drive\MICROS\ProfessionalServices\StoredValueCard"

		$output = & ({ Write-Output "Copying EM data..." }) | Out-String; $ConOut.AppendText($output)
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Res\EM3" -Destination "$BackupFolder\D_drive\MICROS\Res"

		$output = & ({ Write-Output "Copying CAL data..." }) | Out-String; $ConOut.AppendText($output)
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Res\CAL\Win32\Files\Micros\Res\Pos\Bin" -Destination "$BackupFolder\D_drive\MICROS\Res\CAL\Win32\Files\Micros\Res\Pos"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Res\CAL\Win32\Files\Micros\Res\Pos\Bitmaps" -Destination "$BackupFolder\D_drive\MICROS\Res\CAL\Win32\Files\Micros\Res\Pos"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Res\CAL\Win32\Files\Micros\Res\Pos\Etc" -Destination "$BackupFolder\D_drive\MICROS\Res\CAL\Win32\Files\Micros\Res\Pos"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Res\CAL\Win32\Files\Micros\Res\Pos\Scripts" -Destination "$BackupFolder\D_drive\MICROS\Res\CAL\Win32\Files\Micros\Res\Pos"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Res\CAL\Win32\Files\Micros\Res\Pos\Temp" -Destination "$BackupFolder\D_drive\MICROS\Res\CAL\Win32\Files\Micros\Res\Pos"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Res\CAL\WS5\Files\CF\Micros\Bitmaps" -Destination "$BackupFolder\D_drive\MICROS\Res\CAL\WS5\Files\CF\Micros"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Res\CAL\WS5\Files\CF\Micros\ETC" -Destination "$BackupFolder\D_drive\MICROS\Res\CAL\WS5\Files\CF\Micros"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Res\CAL\WS5\Files\CF\Micros\Scripts" -Destination "$BackupFolder\D_drive\MICROS\Res\CAL\WS5\Files\CF\Micros"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Res\CAL\WS5A\Files\CF\Micros\Bitmaps" -Destination "$BackupFolder\D_drive\MICROS\Res\CAL\WS5A\Files\CF\Micros"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Res\CAL\WS5A\Files\CF\Micros\ETC" -Destination "$BackupFolder\D_drive\MICROS\Res\CAL\WS5A\Files\CF\Micros"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Res\CAL\WS5A\Files\CF\Micros\Scripts" -Destination "$BackupFolder\D_drive\MICROS\Res\CAL\WS5A\Files\CF\Micros"

		$output = & ({ Write-Output "Copying POS data..." }) | Out-String; $ConOut.AppendText($output)
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Res\Pos\Bitmaps" -Destination "$BackupFolder\D_drive\MICROS\Res\Pos"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Res\Pos\Journals" -Destination "$BackupFolder\D_drive\MICROS\Res\Pos"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Res\Pos\Scripts" -Destination "$BackupFolder\D_drive\MICROS\Res\Pos"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Res\Pos\Temp" -Destination "$BackupFolder\D_drive\MICROS\Res\Pos"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Res\Pos\Etc\*.isl" -Destination "$BackupFolder\D_drive\MICROS\Res\Pos\Etc"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Res\Pos\Etc\QsrMicrosTSConnect.dll" -Destination "$BackupFolder\D_drive\MICROS\Res\Pos\Etc"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Res\Pos\Etc\QsrSettings.xml" -Destination "$BackupFolder\D_drive\MICROS\Res\Pos\Etc"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Res\Pos\Etc\DineTimeTableBoss.xml" -Destination "$BackupFolder\D_drive\MICROS\Res\Pos\Etc"			
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Res\Pos\Reports\Custom" -Destination "$BackupFolder\D_drive\MICROS\Res\Pos\Reports"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\MICROS\Res\Pos\Reports\40 Column\Custom" -Destination "$BackupFolder\D_drive\MICROS\Res\Pos\Reports\40 Column"
			
		$output = & ({ Write-Output "Copying MICROS database..." }) | Out-String; $ConOut.AppendText($output)
			$LatestDB = (Get-ChildItem "\\$DP_RES_IP\d$\MICROS\Database\Data\Backup\Archive" | Where-Object {$_.LastWriteTime -gt (Get-Date -Format "yyyy-MM-dd")}).FullName
			Copy-Item -Recurse -Force -Path $LatestDB -Destination "$BackupFolder\D_drive\MICROS\Database\Data\Backup\Archive"
			
		$output = & ({ Write-Output "Copying XPO data..." }) | Out-String; $ConOut.AppendText($output)
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\xpoDoordashAgent" -Destination "$BackupFolder\D_drive"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\xpoMicrosAgent" -Destination "$BackupFolder\D_drive"
			
		$output = & ({ Write-Output "Copying Misc data..." }) | Out-String; $ConOut.AppendText($output)
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\emkeytransfer.bat" -Destination "$BackupFolder\D_drive"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\emreceive.ps1" -Destination "$BackupFolder\D_drive"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\Micros Check Close Tool.exe" -Destination "$BackupFolder\D_drive"
			Copy-Item -Recurse -Force -Path "\\$DP_RES_IP\d$\vncviewer.exe" -Destination "$BackupFolder\D_drive"

		$output = & ({ Write-Output "Copying Salescash data..." }) | Out-String; $ConOut.AppendText($output)
			$RecentSalesCash = (Get-ChildItem "\\$DP_RES_IP\c$\Salescash\archive" | Where-Object {$_.LastWriteTime -gt '12/31/2020'}).FullName
			Copy-Item -Recurse -Force -Path $RecentSalesCash -Destination "$BackupFolder\C_Drive\Salescash\archive"
			
		$output = & ({ Write-Output "RES DATA COMPLETE" }) | Out-String; $ConOut.AppendText($output)
		
		$output = & ({
		if ($DP_CSK_IP -eq $null) {Write-Output "IP Address invalid"; return}
		$TestCSK = Test-Path "\\$DP_CSK_IP\c$\ProgramData\QSR Automations"
		if ($TestCSK -eq $false) {Write-Output "Network path invalid"; return}
		}) | Out-String; $ConOut.AppendText($output)
		
		$output = & ({
			$script:datestamp = Get-Date -Format "yyyy-MM-dd_HH.mm"
			$script:BackupFolder = "D:\TEMP\"+$CSK_Name+"_BACKUP_$datestamp"
			Write-Output "Backing up CSK data to $BackupFolder"
		}) | Out-String; $ConOut.AppendText($output)
		
		Start-Sleep 5
		
		New-Item -Type Directory -Force "$BackupFolder\ProgramData\QSR Automations\ConnectSmart\Common"
		New-Item -Type Directory -Force "$BackupFolder\ProgramData\QSR Automations\ConnectSmart\ControlPointServer"
		New-Item -Type Directory -Force "$BackupFolder\ProgramData\QSR Automations\ConnectSmart\KitchenServer"
		New-Item -Type Directory -Force "$BackupFolder\MicrosTSConnect_data"
		
		$output = & ({ Write-Output "Copying Common data..." }) | Out-String; $ConOut.AppendText($output)
			Copy-Item -Recurse -Force -Path "\\$DP_CSK_IP\c$\ProgramData\QSR Automations\ConnectSmart\Common\Data" -Destination "$BackupFolder\ProgramData\QSR Automations\ConnectSmart\Common"
			Copy-Item -Recurse -Force -Path "\\$DP_CSK_IP\c$\ProgramData\QSR Automations\ConnectSmart\Common\License" -Destination "$BackupFolder\ProgramData\QSR Automations\ConnectSmart\Common"

		$output = & ({ Write-Output "Copying ControlPointServer data..." }) | Out-String; $ConOut.AppendText($output)
			Copy-Item -Recurse -Force -Path "\\$DP_CSK_IP\c$\ProgramData\QSR Automations\ConnectSmart\ControlPointServer\Data" -Destination "$BackupFolder\ProgramData\QSR Automations\ConnectSmart\ControlPointServer"
			Copy-Item -Recurse -Force -Path "\\$DP_CSK_IP\c$\ProgramData\QSR Automations\ConnectSmart\ControlPointServer\Templates" -Destination "$BackupFolder\ProgramData\QSR Automations\ConnectSmart\ControlPointServer"

		$output = & ({ Write-Output "Copying KitchenServer data..." }) | Out-String; $ConOut.AppendText($output)
			Copy-Item -Recurse -Force -Path "\\$DP_CSK_IP\c$\ProgramData\QSR Automations\ConnectSmart\KitchenServer\Data" -Destination "$BackupFolder\ProgramData\QSR Automations\ConnectSmart\KitchenServer"
			Copy-Item -Recurse -Force -Path "\\$DP_CSK_IP\c$\ProgramData\QSR Automations\ConnectSmart\KitchenServer\Images" -Destination "$BackupFolder\ProgramData\QSR Automations\ConnectSmart\KitchenServer"
			Copy-Item -Recurse -Force -Path "\\$DP_CSK_IP\c$\ProgramData\QSR Automations\ConnectSmart\KitchenServer\SpeedOfService" -Destination "$BackupFolder\ProgramData\QSR Automations\ConnectSmart\KitchenServer"
		
		$output = & ({ Write-Output "Copying MicrosTSConnect data..." }) | Out-String; $ConOut.AppendText($output)
			Copy-Item -Recurse -Force -Path "\\$DP_CSK_IP\c$\Program Files (x86)\QSR Automations\MicrosTSConnect\Data\*" -Destination "$BackupFolder\MicrosTSConnect_data"

		$output = & ({ Write-Output "CSK DATA COMPLETE" }) | Out-String; $ConOut.AppendText($output)
		
		$output = & ({ Write-Output ""; Write-Output "!!-- DATA PULL COMPLETE --!!" }) | Out-String; $ConOut.AppendText($output)
})

$Timezone.Add_Click({
$ButtonType = [System.Windows.MessageBoxButton]::YesNo
$MessageIcon = [System.Windows.MessageBoxImage]::Information
$MessageBody = "Change timezone on all VMs and Hyper-V host?"
$MessageTitle = "Confirm Operation"
$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
	
	If ($Result -eq 'Yes') {
		
		$output = & ({			
			Set-CustomCreds
			If ($adminacc -eq $null) {Write-Output "Credentials not captured"; return}			
		}) | Out-String; $ConOut.AppendText($output)
		
		$output = & ({
			$TZ = $null
			
			$TZform = New-Object System.Windows.Forms.Form
			$TZform.Text = $ScriptName
			$TZform.Size = New-Object System.Drawing.Size(300,200)
			$TZform.StartPosition = 'CenterScreen'
			
			$TZokb = New-Object System.Windows.Forms.Button
			$TZokb.Location = New-Object System.Drawing.Point(65,130)
			$TZokb.Size = New-Object System.Drawing.Size(75,25)
			$TZokb.Text = 'OK'
			$TZokb.DialogResult = [System.Windows.Forms.DialogResult]::OK
			$TZform.AcceptButton = $TZokb
			$TZform.Controls.Add($TZokb)
			
			$TZlist = New-Object System.Windows.Forms.ListBox
			$TZlist.Location = New-Object System.Drawing.Point(10,40)
			$TZlist.Size = New-Object System.Drawing.Size(255,20)
			$TZlist.Height = 75
			$TZlist.BorderStyle = "None"
			$TZlist.Font = $H3 
			$TZlist.Items.Add('Pacific Standard Time') | Out-Null
			$TZlist.Items.Add('Mountain Standard Time') | Out-Null
			$TZlist.Items.Add('Central Standard Time') | Out-Null
			$TZlist.Items.Add('Eastern Standard Time') | Out-Null
			$TZform.Controls.Add($TZlist)
			
			$TZcb = New-Object System.Windows.Forms.Button
			$TZcb.Location = New-Object System.Drawing.Point(150,130)
			$TZcb.Size = New-Object System.Drawing.Size(75,25)
			$TZcb.Text = 'Cancel'
			$TZcb.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
			$TZform.CancelButton = $TZcb
			$TZform.Controls.Add($TZcb)
			
			$TZlb1 = New-Object System.Windows.Forms.Label
			$TZlb1.Location = New-Object System.Drawing.Point(10,15)
			$TZlb1.Size = New-Object System.Drawing.Size(240,20)
			$TZlb1.Text = 'Timezone:'
			$TZform.Controls.Add($TZlb1)
			
			$TZform.Topmost = $true
			$TZrs = $TZform.ShowDialog()
				if ($TZrs -eq [System.Windows.Forms.DialogResult]::OK) {
					$script:TZ = $TZlist.SelectedItem
				}

			If ($TZ -eq $null) {Write-Output "Timezone invalid"}
				
		}) | Out-String; $ConOut.AppendText($output)
		
		If ($TZ -eq $null) {return}
				
		$output = & ({Write-Output "Setting Timezone..."}) | Out-String; $ConOut.AppendText($output)
			
		$output = & ({
			Get-VM * | Where-Object {$_.VMName -notlike "*wks*"} | sort | foreach {
				$Name = (Get-VM $_.VMName).Name
				$output = & ({Write-Output "Setting Timezone on $Name..."}) | Out-String; $ConOut.AppendText($output)
				$s = New-PSSession -VMName $Name -Credential $adminacc
				Invoke-Command -Session $s -ScriptBlock {Set-TimeZone $Using:TZ}
				Remove-PSSession $s
				}
				
			Get-VM * | Where-Object {$_.VMName -like "*wks*"} | sort | foreach {
				$Name = (Get-VM $_.VMName).Name
				$output = & ({Write-Output "Setting Timezone on $Name..."}) | Out-String; $ConOut.AppendText($output)
				$username = ($Name + '\user')
				$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username,$userpassword
				$s = New-PSSession -VMName $Name -Credential $creds
				Invoke-Command -Session $s -ScriptBlock {Set-TimeZone $Using:TZ}
				Remove-PSSession $s
				}				
		}) | Out-String; $ConOut.AppendText($output)
		
		$output = & ({
			Write-Output "Setting Timezone on Hyper-V Host..."
			Set-TimeZone $TZ
		}) | Out-String; $ConOut.AppendText($output)
		
		$output = & ({ Write-Output "DONE" }) | Out-String; $ConOut.AppendText($output)

	}
		
})

# COLUMN 4 ACTIONS
$ClearConOut.Add_Click({
		$ConOut.Text = $null
})

$ShowConsole.Add_Click({
    
    $handleConsole = $type::GetConsoleWindow()
    $null = $type::ShowWindowAsync($handleConsole, 4); $null = $type::SetForegroundWindow($handleConsole)
    
  })

$Exit.Add_Click({
$ButtonType = [System.Windows.MessageBoxButton]::YesNo
$MessageIcon = [System.Windows.MessageBoxImage]::Question
$MessageBody = "Exit " + $ScriptName + "?"
$MessageTitle = "Confirm Operation"
$Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)

	If ($Result -eq 'Yes') {
		$Form.close()
	}
})

$null = $Form.ShowDialog()