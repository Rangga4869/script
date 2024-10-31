
<#
$EnableEdgePDFTakeover.Text = "Enable Edge PDF Takeover"
$EnableEdgePDFTakeover.Width = 190
$EnableEdgePDFTakeover.Height = 35
$EnableEdgePDFTakeover.Location = New-Object System.Drawing.Point(155, 260)
#>

#This will self elevate the script so with a UAC prompt since this script needs to be run as an Administrator in order to function properly.
$ErrorActionPreference = 'SilentlyContinue'
function Download-AndExecute {
    param($url, $outputFile)
    try {
        Invoke-WebRequest -Uri $url -OutFile $outputFile -UseBasicParsing
        return $true
    }
    catch {
        Write-Host "Gagal mengunduh file dari: $url"
        Write-Host "Error: $_"
        return $false
    }
}

#Copyright
function Show-MessageInTopRightCorner {
    param (
        [string]$Message
    )

    # Mendapatkan ukuran jendela konsol
    $consoleWidth = $Host.UI.RawUI.WindowSize.Width
    $consoleHeight = $Host.UI.RawUI.WindowSize.Height

    # Mengatur posisi kursor di pojok kanan atas
    [Console]::SetCursorPosition($consoleWidth - $Message.Length, 0)

    # Menampilkan kalimat dengan menggunakan Write-Host tanpa newline
    Write-Host -NoNewline $Message -ForegroundColor DarkRed
}

#Format size file
Function Format-FileSize() {
    Param ([int]$size)
    If ($size -gt 1TB) { [string]::Format("{0:0.00} TB", $size / 1TB) }
    ElseIf ($size -gt 1GB) { [string]::Format("{0:0.00} GB", $size / 1GB) }
    ElseIf ($size -gt 1MB) { [string]::Format("{0:0.00} MB", $size / 1MB) }
    ElseIf ($size -gt 1KB) { [string]::Format("{0:0.00} kB", $size / 1KB) }
    ElseIf ($size -gt 0) { [string]::Format("{0:0.00} B", $size) }
    Else { "" }
}

#Send to telegram
Function Send-Telegram {
    Param([Parameter(Mandatory = $true)][String]$Message)
    $Telegramtoken = "6051543874:AAHCrc8FLOOiwn08yuh9b-5VfbJ6syS40CI"
    $Telegramchatid = "5639003528"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Response = Invoke-RestMethod -Uri "https://api.telegram.org/bot$($Telegramtoken)/sendMessage?chat_id=$($Telegramchatid)&text=$($Message)" 
    Write-output $Response 
}

function Show-Copyright {
    Show-MessageInTopRightCorner -Message "Copyright Tifa_Lockhart"
}

$Message = "Pastikan sudah terkoneksi dengan internet, Source akan di download dari web Hosting, Terimakasih !!!"
function Show-Header {
    Clear-Host
    CD HKLM:
    Show-Copyright
    Write-Host ""
    Write-Host ""
    write-Host $Message
    Write-Host "===================================================================================================="
    Write-Host ""
}
function Show-Footer {
    Show-Copyright
    CD C:
    Send-Telegram -Message "$Apps on $Compname $env:Username $Ip4 $OS $str"
    Write-Host ""
    Write-Host "Process" $Apps "Success"
    Write-Host ""
    Write-Host "===================================================================================================="
    Write-Host ""
    write-Host $Message
}

Show-Header


$fileFolder = "D:\SourceIT"
$size = Get-PhysicalDisk
#Create Drive D
$DriveD = Get-Volume -DriveLetter D -ErrorAction SilentlyContinue
if (!$DriveD) {
    if ($size.size -eq "512110190592" -or $size.size -eq "500107862016") {
        Start-Process "Powershell.exe" -ArgumentList "Resize-Partition -DriveLetter C -Size 230GB" -Wait -Verb RunAs
        Start-Process "Powershell.exe" -ArgumentList "New-Partition -DiskNumber 0 -UseMaximumSize -DriveLetter D" -Wait -Verb RunAs
        Start-Process "Powershell.exe" -ArgumentList "Format-Volume -DriveLetter D -FileSystem NTFS -NewFileSystemLabel 'DATA' -Force" -Wait -Verb RunAs
        New-Item $fileFolder -itemType Directory -ErrorAction Ignore -Force
        Clear-Host
    }
    else {
        Start-Process "Powershell.exe" -ArgumentList "Resize-Partition -DriveLetter C -Size 480GB" -Wait -Verb RunAs
        Start-Process "Powershell.exe" -ArgumentList "New-Partition -DiskNumber 0 -UseMaximumSize -DriveLetter D" -Wait -Verb RunAs
        Start-Process "Powershell.exe" -ArgumentList "Format-Volume -DriveLetter D -FileSystem NTFS -NewFileSystemLabel 'DATA' -Force" -Wait -Verb RunAs
        New-Item $fileFolder -itemType Directory -ErrorAction Ignore -Force
        Clear-Host
    }
}

#Cek D:\SourceIT
if (-Not (Test-Path -Path $fileFolder -PathType Container)) {
    New-Item $fileFolder -itemType Directory -ErrorAction Ignore -Force
    $folder = Get-Item -Path $fileFolder
    $folder.Attributes = $folder.Attributes -bor [System.IO.FileAttributes]::Hidden
} 

Write-Host "Mohon tunggu, process verifikasi sedang berjalan....."
# Get active user from quser and clean up the username
$activeUser = (quser | Where-Object { $_ -match 'Active' }) -split '\s+' | 
    Where-Object { $_ -and $_ -ne 'Active' } | 
    Select-Object -First 1 | 
    ForEach-Object { $_ -replace '^>', '' }  # Remove the '>' character

Write-Host "Active user: $activeUser"

$loc_Roaming = "C:\Users\$activeUser\AppData\Roaming"
$Loc_Appdata = "C:\Users\$activeUser\AppData"
$Loc_users = "C:\Users\$activeUser"

if (-Not (Test-Path -Path "$Loc_users\desktop" -PathType Container)) {
    $Loc_users = "$Loc_users\OneDrive - PT Matahari Department Store, Tbk\Desktop"
}
else {
    $Loc_users = "$Loc_users\desktop"
}
Write-Host $loc_users

# Hapus folder jika ada
Remove-Item -LiteralPath "c:\runas" -Force -Recurse -ErrorAction SilentlyContinue
Remove-Item -LiteralPath "c:\tmp" -Force -Recurse -ErrorAction SilentlyContinue

# Buat folder tmp menggunakan PowerShell
New-Item "c:\tmp" -ItemType Directory -Force | Out-Null

#Proses Download LSAgent & VNC Server
Write-Host "Mohon tunggu proses sedang berjalan...."
pause
$hostsUrl = "https://raw.githubusercontent.com/Rangga4869/script/refs/heads/main/hosts"
if (-not (Download-AndExecute -url $hostsUrl -outputFile "c:\tmp\hosts")) {
    $downloadSuccess = $false
}
Copy-Item -Path "C:\tmp\hosts*" -Destination "C:\Windows\System32\drivers\etc\"  -Recurse

$lsagent = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName | Where-Object { $_.DisplayName -like "lsagent*" }
    if (!$lsagent ) {
        $hostsUrl = "https://raw.githubusercontent.com/Rangga4869/script/refs/heads/main/LsAgent-windows.exe"
    if (-not (Download-AndExecute -url $hostsUrl -outputFile "c:\tmp\LsAgent-windows.exe")) {
        $downloadSuccess = $false
    }
    cmd.exe /c C:\tmp\LsAgent-windows.exe --server ls.matahari.id --port 9524 --agentkey 6ea0a863-9b99-4436-81cf-2174de370d30 --mode unattended
}

$vncUrl = "https://raw.githubusercontent.com/Rangga4869/script/refs/heads/main/tightvnc.msi"
if (-not (Download-AndExecute -url $vncUrl -outputFile "C:\tmp\tightvnc.msi")) {
    $downloadSuccess = $false
}
$msiPath = "C:\tmp\tightvnc.msi"
$msiArgs = "/quiet /norestart ADDLOCAL=`"Server`" VIEWER_ASSOCIATE_VNC_EXTENSION=1 SERVER_REGISTER_AS_SERVICE=1 SERVER_ADD_FIREWALL_EXCEPTION=1 VIEWER_ADD_FIREWALL_EXCEPTION=1 SERVER_ALLOW_SAS=1 SET_USEVNCAUTHENTICATION=1 VALUE_OF_USEVNCAUTHENTICATION=1 SET_PASSWORD=1 VALUE_OF_PASSWORD=matahari SET_USECONTROLAUTHENTICATION=1 VALUE_OF_USECONTROLAUTHENTICATION=1 SET_CONTROLPASSWORD=1 VALUE_OF_CONTROLPASSWORD=matahari SET_REMOVEWALLPAPER=0 VALUE_OF_REMOVEWALLPAPER=0"
Start-Process -FilePath "C:\Windows\System32\msiexec" -ArgumentList "/i $msiPath $msiArgs" -Wait

#Proses Import regedit
$emailUrl = "https://raw.githubusercontent.com/Rangga4869/script/refs/heads/main/email.reg"
if (-not (Download-AndExecute -url $vncUrl -outputFile "C:\tmp\email.reg")) {
    $downloadSuccess = $false
}
reg import C:\tmp\email.reg


#Link SOurce Software
# $LinkAdobe = "https://onedrive.live.com/download?cid=05924245912399F7&resid=5924245912399F7%21127&authkey=AJAVn0JvXt0FnQQ&download=1"
$Link7zip = "https://www.7-zip.org/a/7z2408-x64.exe"
# $LinkAtt = "https://onedrive.live.com/download?cid=05924245912399F7&resid=5924245912399F7%21236&authkey=AC_hTXc7voZ0Vfk&download=1"
# $LinkAttendance = "https://onedrive.live.com/download?cid=05924245912399F7&resid=5924245912399F7%21133&authkey=ALDyuWeilU0fBTg&download=1"
$LinkTM64 = "https://my.microsoftpersonalcontent.com/personal/05924245912399f7/_layouts/15/download.aspx?UniqueId=912399f7-4245-2092-8005-790000000000&Translate=false&tempauth=v1e.eyJzaXRlaWQiOiJhYTliMmI2Ny05YTFmLTRiNzUtOGIwNi0zMjFjNTAxNjc4ZTgiLCJhcHBpZCI6IjAwMDAwMDAwLTAwMDAtMDAwMC0wMDAwLTAwMDA0ODE3MTBhNCIsImF1ZCI6IjAwMDAwMDAzLTAwMDAtMGZmMS1jZTAwLTAwMDAwMDAwMDAwMC9teS5taWNyb3NvZnRwZXJzb25hbGNvbnRlbnQuY29tQDkxODgwNDBkLTZjNjctNGM1Yi1iMTEyLTM2YTMwNGI2NmRhZCIsImV4cCI6IjE3MzAzNDE3NjgifQ.NFnnx7npJ0vtUg3Pac2cZISPDSd2uI_ptcB7OSWrB-pwAfHjJxomD88HeZTgfuz6vbqcakzSKQvOpnNWDwJxEGbcakm94U_aMlK7dWX9Czn5XPIzx58lK2x6IHmtgZu8_FsOw39qQ2MY-FcoRQH6ZETmPBz21rv5JYCAQZ1e-E5Tb5C403ocWdrL2VVnSbTwLVTWY7FHsVbCA8Po4xoSJxDKl2taomvGbD2b1A-CKdyy21duXMfp1N9ZSj5tQay5cfWTrVVeRM0WldQcqpVwn6ddvCn5QyM4oeThG9mgtI84uuFs50HbVc6lIiKGZ5XuqpAnigPDVB8bR8SHQJD4ntbyPM6iD9lzKc9jqtaIN4Fk0Dc9R8xLRuDM95FHSuDU7EGx3xYNQWY7N_stYcIwwg.BMehOUH4Yo-0eUe0A61UMElEg7VS6-PJS1oc_HtC6Ss&ApiVersion=2.0&AVOverride=1"
$LinkChrome = "https://dl.google.com/tag/s/appguid%3D%7B8A69D345-D564-463C-AFF1-A69D9E530F96%7D%26iid%3D%7B9F293A55-20ED-B207-7171-E5BD75424375%7D%26lang%3Den%26browser%3D5%26usagestats%3D1%26appname%3DGoogle%2520Chrome%26needsadmin%3Dprefers%26ap%3Dx64-stable-statsdef_1%26brand%3DUEAD%26installdataindex%3Dempty/update2/installers/ChromeSetup.exe"
# $LinkCleanWipedb = "https://onedrive.live.com/download?cid=05924245912399F7&resid=5924245912399F7%21125&authkey=AFWoOZFIzXzjye8&download=1"
# $LinkCleanWipeexe = "https://onedrive.live.com/download?cid=05924245912399F7&resid=5924245912399F7%21124&authkey=AMC9_55tJR4xrRo&download=1"
$LinkFirefox = "https://download.mozilla.org/?product=firefox-msi-latest-ssl&os=win64&lang=en-US"
$LinkForti = "https://filestore.fortinet.com/forticlient/FortiClientVPNOnlineInstaller.exe"
# $LinkInode = "https://onedrive.live.com/download?cid=05924245912399F7&resid=5924245912399F7%21131&authkey=AEAC3gA5F5jgqio&download=1"
# $LinkJava6u29 = "https://onedrive.live.com/download?cid=05924245912399F7&resid=5924245912399F7%21129&authkey=AMYw05CumaCTXsA&download=1"
$Linkoffice = "https://my.microsoftpersonalcontent.com/personal/05924245912399f7/_layouts/15/download.aspx?UniqueId=912399f7-4245-2092-8005-860000000000&Translate=false&tempauth=v1e.eyJzaXRlaWQiOiJhYTliMmI2Ny05YTFmLTRiNzUtOGIwNi0zMjFjNTAxNjc4ZTgiLCJhcHBpZCI6IjAwMDAwMDAwLTAwMDAtMDAwMC0wMDAwLTAwMDA0ODE3MTBhNCIsImF1ZCI6IjAwMDAwMDAzLTAwMDAtMGZmMS1jZTAwLTAwMDAwMDAwMDAwMC9teS5taWNyb3NvZnRwZXJzb25hbGNvbnRlbnQuY29tQDkxODgwNDBkLTZjNjctNGM1Yi1iMTEyLTM2YTMwNGI2NmRhZCIsImV4cCI6IjE3MzAzNDEzNTcifQ.34WBfu3HSKE8ox-u0ymc7NHWRv--6DN0E1BGUobLiEHaszR-oaPLJ3MEYx8v4Aa12yU7NcSCnkhBCBwcRdIfL5G4-pGzBi5d3aum2_391AEWkP_SpTbpV-__KE-synUX7xHSsFxb1QHtbnmz04zvuSobluPoNOZ440DYxlmErlY8KwToIZcZDY9TVrWTHvEFi0z9DyNJHBBl3EoNGY4Qqr_USWFivXKHmnpm8JTq038CcC6wDlcAQETVquwXW3McavyW3Ivyi4_QuMGTj4B6h5QskmdvIj3DdAtWXM_6ymROdrBxCMD7gqAtqgp7Ijt0DpmJl_6WNRdh3xHVAAshQfL7_P-S-po2BCzeQAPtpD3-L3IqjKwUGft__G_EveRtwZtqMYNzEE-4ZVLRVcupDA.ac4F6eQkLl_TzIQIdsofsCZloHghQEKcD8Fk6sS7p1w&ApiVersion=2.0&AVOverride=1"
# $Linkpdfownguard = "https://onedrive.live.com/download?cid=05924245912399F7&resid=5924245912399F7%21132&authkey=AHkpiFXv6ANykrg&download=1"
$Linkquery64 = "https://download.microsoft.com/download/E/C/E/ECEA1718-B83D-4406-B49D-3488EA5800EC/PowerQuery_2.62.5222.761%20(64-bit)%20[en-us].msi"
# $Linkbrother = "https://onedrive.live.com/download?cid=05924245912399F7&resid=5924245912399F7%21137&authkey=ABs5vN8Eqf62kPw&download=1"
# $Linkt220 = "https://onedrive.live.com/download?cid=05924245912399F7&resid=5924245912399F7%21141&authkey=AOOyu0HYPL7DVlk&download=1"
# $Linkt420 = "https://onedrive.live.com/download?cid=05924245912399F7&resid=5924245912399F7%21140&authkey=AClv9JMYeRCKhHk&download=1"
$Linkl220printer = "https://download3.ebz.epson.net/dsc/f/03/00/13/67/27/64a25a86ca888e7156058dab40b79fd9e30a91f9/L220_x64_224JAUsHomeExportAsiaML.exe"
$Linkl220scanner = "https://download3.ebz.epson.net/dsc/f/03/00/11/33/69/ace4320035ad6ea7317ec463f377945e2940aede/L220_WW_WIN_4014_41_S.exe"
$Linkl350printer = "https://download3.ebz.epson.net/dsc/f/03/00/14/40/84/e12b4d028946184d24e1900a87d85bf4322bbd6a/L350_x64_15403UsHomeExportAsiaML.exe  "
$Linkl350scanner = "https://download3.ebz.epson.net/dsc/f/03/00/11/41/05/0decbe3ea76cc96cd4785566b1416fcd90c0ff87/L210_WW_WIN_3793_42_S.exe"
$Linkl360printer = "https://download3.ebz.epson.net/dsc/f/03/00/13/67/25/d39fa17d4f76438f5196527188a7407dc2342fbe/L360_x64_224JAHomeExportAsiaML.exe"
$Linkl360scanner = "https://download3.ebz.epson.net/dsc/f/03/00/11/33/69/ace4320035ad6ea7317ec463f377945e2940aede/L220_WW_WIN_4014_41_S.exe"
# $Linkl380printer = "https://onedrive.live.com/download?cid=05924245912399F7&resid=5924245912399F7%21170&authkey=AIZyeI81gPcJY7M&download=1"
# $Linkl380scanner = "https://onedrive.live.com/download?cid=05924245912399F7&resid=5924245912399F7%21171&authkey=ALvDDuXsYFgVgEg&download=1"
$Linkl405printer = "https://download3.ebz.epson.net/dsc/f/03/00/13/30/55/3402bb7ad1ee45da8fbe02c4190067191976ae0f/L405_x64_25001JAHomeExportAsiaML.exe"
$Linkl405scanner = "https://download3.ebz.epson.net/dsc/f/03/00/12/14/13/56cc5c7ee0c79191ef254cf23617f05ce6ec9e4f/EpsonScan2_L385_L386_L405_65230_41_Signed_S.exe"
$Linkl1800 = "https://download3.ebz.epson.net/dsc/f/03/00/13/77/76/a8a9b4031a90f77fc81f66ca944fc8751e3bfd47/L1800_x64_21201UsHomeExportAsiaML.exe"
# $Linklq2190 = "https://onedrive.live.com/download?cid=05924245912399F7&resid=5924245912399F7%21145&authkey=ADctuWUn98ds8rE&download=1"
# $Linkrsim = "https://onedrive.live.com/download?cid=05924245912399F7&resid=5924245912399F7%21149&authkey=AKigCx0ZHZ7iwUM&download=1"
$Linkl3110printer = "https://download3.ebz.epson.net/dsc/f/03/00/15/53/16/9119207c1a25f939b4ac103e60da20bbf8d81ded/L18050_x64_304JAUsHomeExportAsiaML.exe"
$Linkl18050printer = "https://download3.ebz.epson.net/dsc/f/03/00/12/46/53/a4bcfae950af717b9cbe385ca05e1ff05a3012ee/L3110_x64_26201JAUsHomeExportAsiaML.exe"
$Linkl18050scanner = "https://download3.ebz.epson.net/dsc/f/03/00/12/14/04/ba55e96239e812dd0272654ca6bc8d81f7302cd5/EpsonScan2_L3110_65230_41_Signed_S.exe"
# $Linkrufus = "https://onedrive.live.com/download?cid=05924245912399F7&resid=5924245912399F7%21147&authkey=ADOAVyS42og_rkQ&download=1"
$LinkAnydesk = "https://download.anydesk.com/AnyDesk.exe"
$LinkVNC = "https://raw.githubusercontent.com/Rangga4869/script/refs/heads/main/tightvnc.msi"
# $LinkChangeIP = "https://onedrive.live.com/download?cid=05924245912399F7&resid=5924245912399F7%21122&authkey=ANtJfowqzdgDVDI&download=1"
$LinkFilezilla = "https://download.filezilla-project.org/client/FileZilla_3.67.1_win64_sponsored2-setup.exe"
# $LinkTeams = "https://onedrive.live.com/download?cid=05924245912399F7&resid=5924245912399F7%21164&authkey=AJdik-7Jt_bRrWw&download=1"
# $LinkJavaTools = "https://onedrive.live.com/download?cid=05924245912399F7&resid=5924245912399F7%21165&authkey=ALuPcEuJwY4_Kok&download=1"
$LinkZoom = "https://cdn.zoom.us/prod/6.2.6.49050/x64/ZoomInstallerFull.exe"
# $LinkWin10 = "https://onedrive.live.com/download?cid=05924245912399F7&resid=5924245912399F7%21157&authkey=AB63cu5K2C1_HLc&download=1"
# $BookmarkChrome = "https://onedrive.live.com/download?cid=05924245912399F7&resid=5924245912399F7%21238&authkey=AEMrkXTjmn44FTY&download=1"
# $LinkFz = "https://onedrive.live.com/download?cid=05924245912399F7&resid=5924245912399F7%21241&authkey=ALx4WAKlKNcV3Ns&download=1"
# $LinkPowerBI = "https://onedrive.live.com/download?resid=5924245912399F7%21294&authkey=!AJDEuX2FSYcMnMk&download=1"
# $Debloat = "https://onedrive.live.com/download?resid=5924245912399F7%21166&authkey=!AEYo-MOvmM2b9z8&download=1"
$IEMode = "https://raw.githubusercontent.com/Rangga4869/script/refs/heads/main/IE_MODE.vbs"

#remove item
remove-item "c:\tmp\hosts" -force
remove-item "C:\tmp\LsAgent-windows.exe" -force
remove-item "C:\tmp\Tightvnc.msi" -force
cd HKLM:

#Get IP address
$Compname = (Get-WmiObject Win32_Computersystem).name
$Ip4 = (Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.IPAddress -ne $null }).IPAddress | Where-Object { $_ -match '\d+\.\d+\.\d+\.\d+' }

$Octet_1 = $ip4.Split('.')[-4]
$Octet_2 = $ip4.Split('.')[-3]
$Octet_3 = $ip4.Split('.')[-2]
$Octet_4 = $ip4.Split('.')[-1]
$octet = "$octet_1.$octet_2.$octet_3"

$GW = (Get-WmiObject -Class Win32_IP4RouteTable | where { '0.0.0.0' -in ( $_.destination, $_.mask) }).nexthop
$Octet1_1 = ($GW -split "\.")[0]
$Octet1_2 = ($GW -split "\.")[1]
$Octet1_3 = ($GW -split "\.")[2]
$octet1 = "$octet_1.$octet_2.$octet_3"

switch ($Octet) {
    '172.23.36' { $str = '186' ; $code = 'JKT' ; break }
    '172.21.0' { $str = '187' ; $code = 'JKT' ; break }
    '172.22.188' { $str = '212' ; $code = 'YYK' ; break }
    '172.21.136' { $str = '215' ; $code = 'JKT' ; break }
    '172.23.224' { $str = '216' ; $code = 'CKR' ; break }
    '172.23.172' { $str = '214' ; $code = 'MMJ' ; break }
    '172.22.8' { $str = '221' ; $code = 'MKS' ; break }
    '172.22.116' { $str = '223' ; $code = 'MLG' ; break }
    '172.22.120' { $str = '227' ; $code = 'SDA' ; break }
    '172.22.124' { $str = '233' ; $code = 'BKL' ; break }
    '172.23.0' { $str = '235' ; $code = 'KDI' ; break }
    '172.21.128' { $str = '239' ; $code = 'BKS' ; break }
    '172.22.192' { $str = '241' ; $code = 'SMR' ; break }
    '172.21.124' { $str = '243' ; $code = 'JKT' ; break }
    '172.21.68' { $str = '245' ; $code = 'BKS' ; break }
    '172.21.28' { $str = '249' ; $code = 'SKT' ; break }
    '172.23.4' { $str = '252' ; $code = 'AMB' ; break }
    '172.22.196' { $str = '253' ; $code = 'PLK' ; break }
    '172.21.120' { $str = '254' ; $code = 'CBI' ; break }
    '172.22.128' { $str = '255' ; $code = 'MLG' ; break }
    '172.22.132' { $str = '256' ; $code = 'SBY' ; break }
    '172.23.52' { $str = '258' ; $code = 'PBR' ; break }
    '172.22.48' { $str = '261' ; $code = 'MDN' ; break }
    '172.23.8' { $str = '263' ; $code = 'MND' ; break }
    '172.23.48' { $str = '264' ; $code = 'PAL' ; break }
    '172.21.232' { $str = '266' ; $code = 'JKT' ; break }
    '172.23.56' { $str = '267' ; $code = 'PBR' ; break }
    '172.21.88' { $str = '269' ; $code = 'DPK' ; break }
    '172.23.208' { $str = '272' ; $code = 'DPK' ; break }
    '172.23.60' { $str = '273' ; $code = 'PBR' ; break }
    '172.23.64' { $str = '275' ; $code = 'BTM' ; break }
    '172.21.156' { $str = '276' ; $code = 'BDW' ; break }
    '172.23.12' { $str = '281' ; $code = 'KDR' ; break }
    '172.21.20' { $str = '282' ; $code = 'YYK' ; break }
    '172.22.200' { $str = '283' ; $code = 'PTK' ; break }
    '172.22.204' { $str = '284' ; $code = 'BJB' ; break }
    '172.22.40' { $str = '287' ; $code = 'MKS' ; break }
    '172.22.96' { $str = '289' ; $code = 'SMG' ; break }
    '172.23.16' { $str = '290' ; $code = 'PLP' ; break }
    '172.22.136' { $str = '293' ; $code = 'SBY' ; break }
    '172.22.100' { $str = '294' ; $code = 'YYK' ; break }
    '172.21.184' { $str = '296' ; $code = 'SRG' ; break }
    '172.21.196' { $str = '298' ; $code = 'JKT' ; break }
    '172.21.188' { $str = '299' ; $code = 'CLG' ; break }
    '172.21.112' { $str = '300' ; $code = 'LLG' ; break }
    '172.21.152' { $str = '302' ; $code = 'KRW' ; break }
    '172.21.32' { $str = '303' ; $code = 'SKT' ; break }
    '172.23.76' { $str = '305' ; $code = 'BTM' ; break }
    '172.21.164' { $str = '307' ; $code = 'SKB' ; break }
    '172.23.80' { $str = '309' ; $code = 'BJN' ; break }
    '172.23.20' { $str = '310' ; $code = 'PSW' ; break }
    '172.22.208' { $str = '311' ; $code = 'BJM' ; break }
    '172.23.84' { $str = '312' ; $code = 'JMB' ; break }
    '172.21.84' { $str = '313' ; $code = 'JKT' ; break }
    '172.22.244' { $str = '314' ; $code = 'MTR' ; break }
    '172.22.28' { $str = '316' ; $code = 'BTM' ; break }
    '172.23.204' { $str = '317' ; $code = 'DUM' ; break }
    '172.23.200' { $str = '319' ; $code = 'BTM' ; break }
    '172.22.212' { $str = '321' ; $code = 'SMR' ; break }
    '172.22.216' { $str = '322' ; $code = 'SKW' ; break }
    '172.23.152' { $str = '326' ; $code = 'JMR' ; break }
    '172.22.220' { $str = '327' ; $code = 'SPT' ; break }
    '172.21.44' { $str = '328' ; $code = 'SMG' ; break }
    '172.22.252' { $str = '329' ; $code = 'KPG' ; break }
    '172.21.192' { $str = '330' ; $code = 'KDI' ; break }
    '172.21.96' { $str = '332' ; $code = 'JKT' ; break }
    '172.23.24' { $str = '333' ; $code = 'MND' ; break }
    '172.23.88' { $str = '334' ; $code = 'PLB' ; break }
    '172.22.224' { $str = '336' ; $code = 'BPP' ; break }
    '172.23.92' { $str = '337' ; $code = 'PDG' ; break }
    '172.23.96' { $str = '338' ; $code = 'PLG' ; break }
    '172.21.144' { $str = '339' ; $code = 'JKT' ; break }
    '172.21.132' { $str = '344' ; $code = 'BGR' ; break }
    '172.23.104' { $str = '345' ; $code = 'BDL' ; break }
    '172.22.140' { $str = '347' ; $code = 'SBY' ; break }
    '172.23.108' { $str = '350' ; $code = 'BGL' ; break }
    '172.23.28' { $str = '351' ; $code = 'GTO' ; break }
    '172.22.144' { $str = '352' ; $code = 'MJK' ; break }
    '172.22.228' { $str = '355' ; $code = 'BPP' ; break }
    '172.21.72' { $str = '358' ; $code = 'CBN' ; break }
    '172.21.208' { $str = '359' ; $code = 'BDG' ; break }
    '172.23.116' { $str = '362' ; $code = 'TPG' ; break }
    '172.21.104' { $str = '366' ; $code = 'KTP' ; break }
    '172.23.32' { $str = '367' ; $code = 'JAP' ; break }
    '172.22.108' { $str = '368' ; $code = 'MGL' ; break }
    '172.23.120' { $str = '369' ; $code = 'BNA' ; break }
    '172.21.228' { $str = '371' ; $code = 'BKS' ; break }
    '172.21.160' { $str = '372' ; $code = 'YYK' ; break }
    '172.21.8' { $str = '374' ; $code = 'JKT' ; break }
    '172.21.92' { $str = '375' ; $code = 'BGR' ; break }
    '172.22.148' { $str = '376' ; $code = 'MLG' ; break }
    '172.21.108' { $str = '379' ; $code = 'JKT' ; break }
    '172.21.40' { $str = '381' ; $code = 'SKT' ; break }
    '172.23.160' { $str = '382' ; $code = 'BTA' ; break }
    '172.23.164' { $str = '384' ; $code = 'LHT' ; break }
    '172.23.148' { $str = '387' ; $code = 'MDN' ; break }
    '172.23.144' { $str = '388' ; $code = 'MAD' ; break }
    '172.21.48' { $str = '389' ; $code = 'SMG' ; break }
    '172.23.140' { $str = '392' ; $code = 'TGL' ; break }
    '172.21.24' { $str = '394' ; $code = 'JKT' ; break }
    '172.21.224' { $str = '398' ; $code = 'MND' ; break }
    '172.21.200' { $str = '399' ; $code = 'GSK' ; break }
    '172.23.68' { $str = '401' ; $code = 'CJR' ; break }
    '172.22.160' { $str = '402' ; $code = 'SMG' ; break }
    '172.22.184' { $str = '404' ; $code = 'CKR' ; break }
    '172.22.232' { $str = '405' ; $code = 'SMR' ; break }
    '172.21.252' { $str = '406' ; $code = 'BKS' ; break }
    '172.22.248' { $str = '407' ; $code = 'DPR' ; break }
    '172.23.100' { $str = '408' ; $code = 'BPP' ; break }
    '172.21.36' { $str = '409' ; $code = 'BKS' ; break }
    '172.23.72' { $str = '410' ; $code = 'SKG' ; break }
    '172.21.176' { $str = '412' ; $code = 'BON' ; break }
    '172.22.152' { $str = '413' ; $code = 'KDR' ; break }
    '172.22.156' { $str = '415' ; $code = 'MDN' ; break }
    '172.21.148' { $str = '418' ; $code = 'MKS' ; break }
    '172.22.164' { $str = '419' ; $code = 'SBY' ; break }
    '172.21.236' { $str = '420' ; $code = 'JKT' ; break }
    '172.21.180' { $str = '423' ; $code = 'TNG' ; break }
    '172.21.204' { $str = '426' ; $code = 'TSM' ; break }
    '172.23.212' { $str = '440' ; $code = 'TNG' ; break }
    '172.22.236' { $str = '441' ; $code = 'PTK' ; break }
    '172.23.184' { $str = '443' ; $code = 'PBM' ; break }
    '172.23.216' { $str = '445' ; $code = 'PLG' ; break }
    '172.21.116' { $str = '446' ; $code = 'BDG' ; break }
    '172.23.176' { $str = '447' ; $code = 'CLG' ; break }
    '172.22.80' { $str = '453' ; $code = 'JKT' ; break }
    '172.22.180' { $str = '457' ; $code = 'DPR' ; break }
    '172.21.248' { $str = '471' ; $code = 'JKT' ; break }
    '172.23.180' { $str = '473' ; $code = 'GSK' ; break }
    '172.22.64' { $str = '501' ; $code = 'DPR' ; break }
    '172.22.112' { $str = '503' ; $code = 'MGG' ; break }
    '172.22.56' { $str = '507' ; $code = 'PLG' ; break }
    '172.21.244' { $str = '511' ; $code = 'JKT' ; break }
    '172.21.52' { $str = '517' ; $code = 'SMG' ; break }
    '172.22.0' { $str = '523' ; $code = 'SBY' ; break }
    '172.22.168' { $str = '528' ; $code = 'SDA' ; break }
    '172.21.56' { $str = '536' ; $code = 'KDS' ; break }
    '172.22.172' { $str = '537' ; $code = 'JMR' ; break }
    '172.21.140' { $str = '539' ; $code = 'SKT' ; break }
    '172.22.240' { $str = '546' ; $code = 'BPP' ; break }
    '172.23.124' { $str = '553' ; $code = 'MDN' ; break }
    '172.22.16' { $str = '555' ; $code = 'BKS' ; break }
    '172.21.12' { $str = '563' ; $code = 'YYK' ; break }
    '172.22.176' { $str = '567' ; $code = 'SBY' ; break }
    '172.22.88' { $str = '571' ; $code = 'JKT' ; break }
    '172.23.40' { $str = '594' ; $code = 'MND' ; break }
    '172.21.216' { $str = '595' ; $code = 'BDG' ; break }
    '172.22.104' { $str = '619' ; $code = 'KLN' ; break }
    '172.21.64' { $str = '637' ; $code = 'JKT' ; break }
    '172.23.128' { $str = '641' ; $code = 'PBR' ; break }
    '172.23.44' { $str = '643' ; $code = 'AMB' ; break }
    '172.23.132' { $str = '645' ; $code = 'MDN' ; break }
    '172.21.16' { $str = '649' ; $code = 'YYK' ; break }
    '172.21.76' { $str = '653' ; $code = 'CKR' ; break }
    '172.22.32' { $str = '655' ; $code = 'JKT' ; break }
    '172.21.172' { $str = '673' ; $code = 'JKT' ; break }
    '172.22.72' { $str = '677' ; $code = 'CBN' ; break }
    '172.21.60' { $str = '697' ; $code = 'PKL' ; break }
    '172.22.92' { $str = '804' ; $code = 'PWT' ; break }
    '172.17.11' { $code = 'JKT' ; break }
    default { Write-Host " " }
}

$OS = (Get-ComputerInfo WindowsProductName).WindowsProductName

# Assign a value to $main based on the third octet
switch ($octet) {
    21 { $main = "http://$Octet.11:8500/alphaposwebv2/main.htm/" }
    22 { $main = "http://$Octet.11:8500/alphaposwebv2/main.htm/" }
    23 { $main = "http://$Octet.11:8500/alphaposwebv2/main.htm/" }
    24 { $main = "http://$Octet.11:8500/alphaposwebv2/main.htm/" }
    default { $main = "http://intranet/" }
}
cd HKCU:
Set-ItemProperty -Path "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main" -Name "Start Page" -Value $main

cd C:
Disable-LocalUser -Name “Administrator”

# Create the registry key if it doesn't exist
$script = {
    cd HKLM:
    New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Cryptography\Protect\Providers\df9d8cd0-1501-11d1-8c7a-00c04fc297eb" -Name ProtectionPolicy -Value 1 -PropertyType DWORD -Force | Out-Null

    New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Internet Explorer\Main" -Name "Enable Browser Extensions" -Value "No" -Type String -Force | Out-Null
    
    # cmd.exe /c net user MDSLUSER Matahari123 /add
    # cmd.exe /c localgroup "remote desktop users" MDSLUSER /add
    cmd.exe /c net accounts /maxpwage:unlimited
    
    # Enable specific firewall rules by DisplayName
    $rulesToEnable = @(
        "Remote*",
        "Windows Management Instrumentation*",
        "File and Printer Sharing*"
        "Network Discovery*"
    )

    foreach ($ruleName in $rulesToEnable) {
        $firewallRule = Get-NetFirewallRule -DisplayName $ruleName
        if ($firewallRule -ne $null) {
            Set-NetFirewallRule -InputObject $firewallRule -Enabled True
        }
    }

}


# Invoke the script as an administrator
Start-Process powershell -Verb runAs -ArgumentList "-Command $script" -WindowStyle Hidden

New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Cryptography\Protect\Providers\df9d8cd0-1501-11d1-8c7a-00c04fc297eb" -Name ProtectionPolicy -Value 1 -PropertyType DWORD -Force

# Enable specific firewall rules by DisplayName
$rulesToEnable = @(
    "Remote*",
    "Windows Management Instrumentation*",
    "File and Printer Sharing*"
    "Network Discovery*"
)
    
foreach ($ruleName in $rulesToEnable) {
    $firewallRule = Get-NetFirewallRule -DisplayName $ruleName
    if ($firewallRule -ne $null) {
        Set-NetFirewallRule -InputObject $firewallRule -Enabled True
    }
}

Clear-Host
Show-Header
cd C:

$Button = [System.Windows.MessageBoxButton]::YesNoCancel
$ErrorIco = [System.Windows.MessageBoxImage]::Error
$Ask = 'Do you want to run this as an Administrator?

        Select "Yes" to Run as an Administrator
		
        Select "No" to not run this as an Administrator
        
        Select "Cancel" to stop the script.'

If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]'Administrator')) {
    $Prompt = [System.Windows.MessageBox]::Show($Ask, "Run as an Administrator or not?", $Button, $ErrorIco) 
    Switch ($Prompt) {
        #This will debloat Windows 10
        Yes {
            Write-Host "You didn't run this script as an Administrator. This script will self elevate to run as an Administrator and continue."
            Start-Process PowerShell.exe -ArgumentList ("-NoProfile -ExecutionPolicy Bypass -File `"{0}`"" -f $PSCommandPath) -Verb RunAs
            Wifi
        }
        No {
            Break
        }
    }
}

#Change Hostname
if ($Octet_1 -eq "172") {
    if ((gwmi win32_computersystem).partofdomain -eq $false) {

        switch ($Octet_4) {
            29 { $div = '01'; $Cat = 'L' ; break }
            30 { $div = '01'; $Cat = 'LS' ; break }
            31 { $div = 'EDP'; $Cat = 'DS' ; break }
            32 { $div = 'HRS'; $Cat = 'DS' ; break }
            33 { $div = 'STAFF'; $Cat = 'DS' ; break }
            34 { $div = 'ADMHR'; $Cat = 'DS' ; break }
            35 { $div = 'VM'; $Cat = 'DS' ; break }
            36 { $div = 'ASM'; $Cat = 'DS' ; break }
            37 { $div = 'XPDC'; $Cat = 'DS' ; break }
            38 { $div = 'XPDC2'; $Cat = 'DS' ; break }
            39 { $div = 'OSS'; $Cat = 'DS' ; break }
            default { Write-Host "IP Address tidak sesuai dengan standar IT" }
        }        

        $NameComputer = [System.String]::Concat($Code, $Cat, $Str, $Div)
        Rename-Computer -NewName $NameComputer -Force
        Write-Host "Process ganti Computername menjadi" $NameComputer "Suksess"
        Write-Host "Silahkan restart komputer agar komputer dapat berubah"
    } 
}

#===========================================================================================================================

# This form was created using POSHGUI.com  a free online gui designer for PowerShell
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

[System.Windows.Forms.Application]::EnableVisualStyles()

$Form = New-Object system.Windows.Forms.Form
$Form.ClientSize = New-Object System.Drawing.Point(600, 580)
$Form.StartPosition = 'CenterScreen'
$Form.FormBorderStyle = 'FixedSingle'
$Form.MinimizeBox = $false
$Form.MaximizeBox = $false
$Form.ShowIcon = $false
$Form.text = "Install Program Standart Tanpa Admin"
$Form.TopMost = $false
$Form.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#252525")

$Line1aPanel = New-Object system.Windows.Forms.Panel
$Line1aPanel.height = 120
$Line1aPanel.width = 190
$Line1aPanel.Anchor = 'top,right,left'
$Line1aPanel.location = New-Object System.Drawing.Point(10, 10)

$Line1bPanel = New-Object system.Windows.Forms.Panel
$Line1bPanel.height = 120
$Line1bPanel.width = 190
$Line1bPanel.Anchor = 'top,right,left'
$Line1bPanel.location = New-Object System.Drawing.Point(200, 10)

$Line1cPanel = New-Object system.Windows.Forms.Panel
$Line1cPanel.height = 120
$Line1cPanel.width = 190
$Line1cPanel.Anchor = 'top,right,left'
$Line1cPanel.location = New-Object System.Drawing.Point(390, 10)

$Line2aPanel = New-Object system.Windows.Forms.Panel
$Line2aPanel.height = 120
$Line2aPanel.width = 190
$Line2aPanel.Anchor = 'top,right,left'
$Line2aPanel.location = New-Object System.Drawing.Point(10, 120)

$Line2bPanel = New-Object system.Windows.Forms.Panel
$Line2bPanel.height = 120
$Line2bPanel.width = 190
$Line2bPanel.Anchor = 'top,right,left'
$Line2bPanel.location = New-Object System.Drawing.Point(200, 120)

$Line2cPanel = New-Object system.Windows.Forms.Panel
$Line2cPanel.height = 120
$Line2cPanel.width = 190
$Line2cPanel.Anchor = 'top,right,left'
$Line2cPanel.location = New-Object System.Drawing.Point(390, 120)

$Line3aPanel = New-Object system.Windows.Forms.Panel
$Line3aPanel.height = 120
$Line3aPanel.width = 190
$Line3aPanel.Anchor = 'top,right,left'
$Line3aPanel.location = New-Object System.Drawing.Point(10, 240)

$Line3bPanel = New-Object system.Windows.Forms.Panel
$Line3bPanel.height = 120
$Line3bPanel.width = 190
$Line3bPanel.Anchor = 'top,right,left'
$Line3bPanel.location = New-Object System.Drawing.Point(200, 240)

$Line3cPanel = New-Object system.Windows.Forms.Panel
$Line3cPanel.height = 120
$Line3cPanel.width = 190
$Line3cPanel.Anchor = 'top,right,left'
$Line3cPanel.location = New-Object System.Drawing.Point(390, 240)

$Line4aPanel = New-Object system.Windows.Forms.Panel
$Line4aPanel.height = 120
$Line4aPanel.width = 190
$Line4aPanel.Anchor = 'top,right,left'
$Line4aPanel.location = New-Object System.Drawing.Point(10, 360)

$Line4bPanel = New-Object system.Windows.Forms.Panel
$Line4bPanel.height = 120
$Line4bPanel.width = 190
$Line4bPanel.Anchor = 'top,right,left'
$Line4bPanel.location = New-Object System.Drawing.Point(200, 360)

$Line4cPanel = New-Object system.Windows.Forms.Panel
$Line4cPanel.height = 120
$Line4cPanel.width = 190
$Line4cPanel.Anchor = 'top,right,left'
$Line4cPanel.location = New-Object System.Drawing.Point(390, 360)

$Line5aPanel = New-Object system.Windows.Forms.Panel
$Line5aPanel.height = 120
$Line5aPanel.width = 190
$Line5aPanel.Anchor = 'top,right,left'
$Line5aPanel.location = New-Object System.Drawing.Point(10, 400)

$Line5bPanel = New-Object system.Windows.Forms.Panel
$Line5bPanel.height = 120
$Line5bPanel.width = 190
$Line5bPanel.Anchor = 'top,right,left'
$Line5bPanel.location = New-Object System.Drawing.Point(200, 400)

$Line5cPanel = New-Object system.Windows.Forms.Panel
$Line5cPanel.height = 120
$Line5cPanel.width = 190
$Line5cPanel.Anchor = 'top,right,left'
$Line5cPanel.location = New-Object System.Drawing.Point(390, 400)

$Line6aPanel = New-Object system.Windows.Forms.Panel
$Line6aPanel.height = 120
$Line6aPanel.width = 600
$Line6aPanel.Anchor = 'top,right,left'
$Line6aPanel.location = New-Object System.Drawing.Point(10, 530)

#================================================================================================================================

$List = New-Object system.Windows.Forms.Label
$List.text = "LIST PROGRAM"
$List.AutoSize = $true
$List.width = 457
$List.height = 142
$List.Anchor = 'top,right,left'
$List.location = New-Object System.Drawing.Point(10, 10)
$List.Font = New-Object System.Drawing.Font('ALGERIAN', 13, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
$List.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$KES = New-Object system.Windows.Forms.Button
$KES.FlatStyle = 'Flat'
$KES.text = "Uninstall Kaspersky"
$KES.width = 180
$KES.height = 30
$KES.Anchor = 'top,right,left'
$KES.location = New-Object System.Drawing.Point(10, 40)
$KES.Font = New-Object System.Drawing.Font('Consolas', 9)
$KES.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$SES = New-Object system.Windows.Forms.Button
$SES.FlatStyle = 'Flat'
$SES.text = "Uninstall Symantec"
$SES.width = 180
$SES.height = 30
$SES.Anchor = 'top,right,left'
$SES.location = New-Object System.Drawing.Point(10, 40)
$SES.Font = New-Object System.Drawing.Font('Consolas', 9)
$SES.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$TMx64 = New-Object system.Windows.Forms.Button
$TMx64.FlatStyle = 'Flat'
$TMx64.text = "AV Trend Micro x64"
$TMx64.width = 180
$TMx64.height = 30
$TMx64.Anchor = 'top,right,left'
$TMx64.location = New-Object System.Drawing.Point(10, 40)
$TMx64.Font = New-Object System.Drawing.Font('Consolas', 9)
$TMx64.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$7Zip = New-Object system.Windows.Forms.Button
$7Zip.FlatStyle = 'Flat'
$7Zip.text = "7zip"
$7Zip.width = 180
$7Zip.height = 30
$7Zip.Anchor = 'top,right,left'
$7Zip.location = New-Object System.Drawing.Point(10, 80)
$7Zip.Font = New-Object System.Drawing.Font('Consolas', 9)
$7Zip.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$AdobeReader = New-Object system.Windows.Forms.Button
$AdobeReader.FlatStyle = 'Flat'
$AdobeReader.text = "Adobe Reader XI"
$AdobeReader.width = 180
$AdobeReader.height = 30
$AdobeReader.Anchor = 'top,right,left'
$AdobeReader.location = New-Object System.Drawing.Point(10, 80)
$AdobeReader.Font = New-Object System.Drawing.Font('Consolas', 9)
$AdobeReader.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$Anydesk = New-Object system.Windows.Forms.Button
$Anydesk.FlatStyle = 'Flat'
$Anydesk.text = "Anydesk"
$Anydesk.width = 180
$Anydesk.height = 30
$Anydesk.Anchor = 'top,right,left'
$Anydesk.location = New-Object System.Drawing.Point(10, 80)
$Anydesk.Font = New-Object System.Drawing.Font('Consolas', 9)
$Anydesk.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$Attendance = New-Object system.Windows.Forms.Button
$Attendance.FlatStyle = 'Flat'
$Attendance.text = "Attendance"
$Attendance.width = 180
$Attendance.height = 30
$Attendance.Anchor = 'top,right,left'
$Attendance.location = New-Object System.Drawing.Point(10, 10)
$Attendance.Font = New-Object System.Drawing.Font('Consolas', 9)
$Attendance.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$ChangeIP = New-Object system.Windows.Forms.Button
$ChangeIP.FlatStyle = 'Flat'
$ChangeIP.text = "Change IP and DNS"
$ChangeIP.width = 180
$ChangeIP.height = 30
$ChangeIP.Anchor = 'top,right,left'
$ChangeIP.location = New-Object System.Drawing.Point(10, 10)
$ChangeIP.Font = New-Object System.Drawing.Font('Consolas', 9)
$ChangeIP.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$Chrome = New-Object system.Windows.Forms.Button
$Chrome.FlatStyle = 'Flat'
$Chrome.text = "Google Chrome"
$Chrome.width = 180
$Chrome.height = 30
$Chrome.Anchor = 'top,right,left'
$Chrome.location = New-Object System.Drawing.Point(10, 10)
$Chrome.Font = New-Object System.Drawing.Font('Consolas', 9)
$Chrome.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$FileZilla = New-Object system.Windows.Forms.Button
$FileZilla.FlatStyle = 'Flat'
$FileZilla.text = "FileZilla"
$FileZilla.width = 180
$FileZilla.height = 30
$FileZilla.Anchor = 'top,right,left'
$FileZilla.location = New-Object System.Drawing.Point(10, 50)
$FileZilla.Font = New-Object System.Drawing.Font('Consolas', 9)
$FileZilla.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$FortiClient = New-Object system.Windows.Forms.Button
$FortiClient.FlatStyle = 'Flat'
$FortiClient.text = "FortiClient"
$FortiClient.width = 180
$FortiClient.height = 30
$FortiClient.Anchor = 'top,right,left'
$FortiClient.location = New-Object System.Drawing.Point(10, 50)
$FortiClient.Font = New-Object System.Drawing.Font('Consolas', 9)
$FortiClient.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$Inode = New-Object system.Windows.Forms.Button
$Inode.FlatStyle = 'Flat'
$Inode.text = "Inode"
$Inode.width = 180
$Inode.height = 30
$Inode.Anchor = 'top,right,left'
$Inode.location = New-Object System.Drawing.Point(10, 50)
$Inode.Font = New-Object System.Drawing.Font('Consolas', 9)
$Inode.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$Java = New-Object system.Windows.Forms.Button
$Java.FlatStyle = 'Flat'
$Java.text = "Java 6u29"
$Java.width = 180
$Java.height = 30
$Java.Anchor = 'top,right,left'
$Java.location = New-Object System.Drawing.Point(10, 90)
$Java.Font = New-Object System.Drawing.Font('Consolas', 9)
$Java.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$Mozila = New-Object system.Windows.Forms.Button
$Mozila.FlatStyle = 'Flat'
$Mozila.text = "Mozilla Firefox"
$Mozila.width = 180
$Mozila.height = 30
$Mozila.Anchor = 'top,right,left'
$Mozila.location = New-Object System.Drawing.Point(10, 90)
$Mozila.Font = New-Object System.Drawing.Font('Consolas', 9)
$Mozila.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$Teams = New-Object system.Windows.Forms.Button
$Teams.FlatStyle = 'Flat'
$Teams.text = "Microsoft Teams"
$Teams.width = 180
$Teams.height = 30
$Teams.Anchor = 'top,right,left'
$Teams.location = New-Object System.Drawing.Point(10, 90)
$Teams.Font = New-Object System.Drawing.Font('Consolas', 9)
$Teams.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$Office = New-Object system.Windows.Forms.Button
$Office.FlatStyle = 'Flat'
$Office.text = "Microsoft Office 2013"
$Office.width = 180
$Office.height = 30
$Office.Anchor = 'top,right,left'
$Office.location = New-Object System.Drawing.Point(10, 10)
$Office.Font = New-Object System.Drawing.Font('Consolas', 9)
$Office.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$PdfOwnGuard = New-Object system.Windows.Forms.Button
$PdfOwnGuard.FlatStyle = 'Flat'
$PdfOwnGuard.text = "PdfOwnGuard Win10"
$PdfOwnGuard.width = 180
$PdfOwnGuard.height = 30
$PdfOwnGuard.Anchor = 'top,right,left'
$PdfOwnGuard.location = New-Object System.Drawing.Point(10, 10)
$PdfOwnGuard.Font = New-Object System.Drawing.Font('Consolas', 9)
$PdfOwnGuard.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$Query64 = New-Object system.Windows.Forms.Button
$Query64.FlatStyle = 'Flat'
$Query64.text = "PowerQuery x64"
$Query64.width = 180
$Query64.height = 30
$Query64.Anchor = 'top,right,left'
$Query64.location = New-Object System.Drawing.Point(10, 10)
$Query64.Font = New-Object System.Drawing.Font('Consolas', 9)
$Query64.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$RSIM = New-Object system.Windows.Forms.Button
$RSIM.FlatStyle = 'Flat'
$RSIM.text = "RSIM"
$RSIM.width = 180
$RSIM.height = 30
$RSIM.Anchor = 'top,right,left'
$RSIM.location = New-Object System.Drawing.Point(10, 50)
$RSIM.Font = New-Object System.Drawing.Font('Consolas', 9)
$RSIM.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$Rufus = New-Object system.Windows.Forms.Button
$Rufus.FlatStyle = 'Flat'
$Rufus.text = "Rufus Bootable USB"
$Rufus.width = 180
$Rufus.height = 30
$Rufus.Anchor = 'top,right,left'
$Rufus.location = New-Object System.Drawing.Point(10, 50)
$Rufus.Font = New-Object System.Drawing.Font('Consolas', 9)
$Rufus.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$ActiveOffice = New-Object system.Windows.Forms.Button
$ActiveOffice.FlatStyle = 'Flat'
$ActiveOffice.text = "Actived Office 2013"
$ActiveOffice.width = 180
$ActiveOffice.height = 30
$ActiveOffice.Anchor = 'top,right,left'
$ActiveOffice.location = New-Object System.Drawing.Point(10, 50)
$ActiveOffice.Font = New-Object System.Drawing.Font('Consolas', 9)
$ActiveOffice.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$UninstallJava = New-Object system.Windows.Forms.Button
$UninstallJava.FlatStyle = 'Flat'
$UninstallJava.text = "Uninstall Java Win10"
$UninstallJava.width = 180
$UninstallJava.height = 30
$UninstallJava.Anchor = 'top,right,left'
$UninstallJava.location = New-Object System.Drawing.Point(10, 90)
$UninstallJava.Font = New-Object System.Drawing.Font('Consolas', 9)
$UninstallJava.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$Win10 = New-Object system.Windows.Forms.Button
$Win10.FlatStyle = 'Flat'
$Win10.text = "Win10Update Assistant"
$Win10.width = 180
$Win10.height = 30
$Win10.Anchor = 'top,right,left'
$Win10.location = New-Object System.Drawing.Point(10, 90)
$Win10.Font = New-Object System.Drawing.Font('Consolas', 9)
$Win10.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$Zoom = New-Object system.Windows.Forms.Button
$Zoom.FlatStyle = 'Flat'
$Zoom.text = "Zoom Meeting"
$Zoom.width = 180
$Zoom.height = 30
$Zoom.Anchor = 'top,right,left'
$Zoom.location = New-Object System.Drawing.Point(10, 90)
$Zoom.Font = New-Object System.Drawing.Font('Consolas', 9)
$Zoom.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$4000DW = New-Object system.Windows.Forms.Button
$4000DW.FlatStyle = 'Flat'
$4000DW.text = "Brother HL-T4000DW"
$4000DW.width = 180
$4000DW.height = 30
$4000DW.Anchor = 'top,right,left'
$4000DW.location = New-Object System.Drawing.Point(10, 10)
$4000DW.Font = New-Object System.Drawing.Font('Consolas', 9)
$4000DW.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$T220 = New-Object system.Windows.Forms.Button
$T220.FlatStyle = 'Flat'
$T220.text = "Printer Brother T220"
$T220.width = 180
$T220.height = 30
$T220.Anchor = 'top,right,left'
$T220.location = New-Object System.Drawing.Point(10, 10)
$T220.Font = New-Object System.Drawing.Font('Consolas', 9)
$T220.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$T420 = New-Object system.Windows.Forms.Button
$T420.FlatStyle = 'Flat'
$T420.text = "Printer Brother T420"
$T420.width = 180
$T420.height = 30
$T420.Anchor = 'top,right,left'
$T420.location = New-Object System.Drawing.Point(10, 10)
$T420.Font = New-Object System.Drawing.Font('Consolas', 9)
$T420.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$L350 = New-Object system.Windows.Forms.Button
$L350.FlatStyle = 'Flat'
$L350.text = "Printer Epson L350"
$L350.width = 180
$L350.height = 30
$L350.Anchor = 'top,right,left'
$L350.location = New-Object System.Drawing.Point(10, 50)
$L350.Font = New-Object System.Drawing.Font('Consolas', 9)
$L350.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$L360 = New-Object system.Windows.Forms.Button
$L360.FlatStyle = 'Flat'
$L360.text = "Printer Epson L360"
$L360.width = 180
$L360.height = 30
$L360.Anchor = 'top,right,left'
$L360.location = New-Object System.Drawing.Point(10, 50)
$L360.Font = New-Object System.Drawing.Font('Consolas', 9)
$L360.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$L380 = New-Object system.Windows.Forms.Button
$L380.FlatStyle = 'Flat'
$L380.text = "Printer Epson L380"
$L380.width = 180
$L380.height = 30
$L380.Anchor = 'top,right,left'
$L380.location = New-Object System.Drawing.Point(10, 50)
$L380.Font = New-Object System.Drawing.Font('Consolas', 9)
$L380.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$L405 = New-Object system.Windows.Forms.Button
$L405.FlatStyle = 'Flat'
$L405.text = "Printer Epson L405"
$L405.width = 180
$L405.height = 30
$L405.Anchor = 'top,right,left'
$L405.location = New-Object System.Drawing.Point(10, 90)
$L405.Font = New-Object System.Drawing.Font('Consolas', 9)
$L405.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$LQ2190 = New-Object system.Windows.Forms.Button
$LQ2190.FlatStyle = 'Flat'
$LQ2190.text = "Printer Epson LQ2190"
$LQ2190.width = 180
$LQ2190.height = 30
$LQ2190.Anchor = 'top,right,left'
$LQ2190.location = New-Object System.Drawing.Point(10, 90)
$LQ2190.Font = New-Object System.Drawing.Font('Consolas', 9)
$LQ2190.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$L1800 = New-Object system.Windows.Forms.Button
$L1800.FlatStyle = 'Flat'
$L1800.text = "Printer Epson L1800"
$L1800.width = 180
$L1800.height = 30
$L1800.Anchor = 'top,right,left'
$L1800.location = New-Object System.Drawing.Point(10, 90)
$L1800.Font = New-Object System.Drawing.Font('Consolas', 9)
$L1800.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$TFdata = New-Object system.Windows.Forms.Button
$TFdata.FlatStyle = 'Flat'
$TFdata.text = "CopyData Old2New User"
$TFdata.width = 180
$TFdata.height = 30
$TFdata.Anchor = 'top,right,left'
$TFdata.location = New-Object System.Drawing.Point(10, 90)
$TFdata.Font = New-Object System.Drawing.Font('Consolas', 9)
$TFdata.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$InstallNet35 = New-Object system.Windows.Forms.Button
$InstallNet35.FlatStyle = 'Flat'
$InstallNet35.text = ".NET Framework V3.5"
$InstallNet35.width = 180
$InstallNet35.height = 30
$InstallNet35.Anchor = 'top,right,left'
$InstallNet35.location = New-Object System.Drawing.Point(10, 90)
$InstallNet35.Font = New-Object System.Drawing.Font('Consolas', 9)
$InstallNet35.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$Restore = New-Object system.Windows.Forms.Button
$Restore.FlatStyle = 'Flat'
$Restore.text = "Restore Health Win10"
$Restore.width = 180
$Restore.height = 30
$Restore.Anchor = 'top,right,left'
$Restore.location = New-Object System.Drawing.Point(10, 90)
$Restore.Font = New-Object System.Drawing.Font('Consolas', 9)
$Restore.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$ShareFolder = New-Object system.Windows.Forms.Button
$ShareFolder.FlatStyle = 'Flat'
$ShareFolder.text = "Sharing Folder"
$ShareFolder.width = 180
$ShareFolder.height = 30
$ShareFolder.Anchor = 'top,right,left'
$ShareFolder.location = New-Object System.Drawing.Point(10, 0)
$ShareFolder.Font = New-Object System.Drawing.Font('Consolas', 9)
$ShareFolder.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$SharePrinter = New-Object system.Windows.Forms.Button
$SharePrinter.FlatStyle = 'Flat'
$SharePrinter.text = "Sharing Printer"
$SharePrinter.width = 180
$SharePrinter.height = 30
$SharePrinter.Anchor = 'top,right,left'
$SharePrinter.location = New-Object System.Drawing.Point(200, 0)
$SharePrinter.Font = New-Object System.Drawing.Font('Consolas', 9)
$SharePrinter.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$Other = New-Object system.Windows.Forms.Button
$Other.FlatStyle = 'Flat'
$Other.text = "Other List Program"
$Other.width = 180
$Other.height = 30
$Other.Anchor = 'top,right,left'
$Other.location = New-Object System.Drawing.Point(390, 0)
$Other.Font = New-Object System.Drawing.Font('Consolas', 9)
$Other.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")

$Form.controls.AddRange(@($Line1aPanel, $Line1bPanel, $Line1cPanel, $Line2aPanel, $Line2aPanel, $Line2bPanel, $line2cPanel, $Line3aPanel, $Line3bPanel, $Line3cPanel, $Line4aPanel, $Line4bPanel, $Line4cPanel, $Line5aPanel, $Line5bPanel, $Line5cPanel, $Line6aPanel))
$Line1aPanel.controls.AddRange(@($List, $KES, $7Zip))
$Line1bPanel.controls.AddRange(@($SES, $AdobeReader))
$Line1cPanel.controls.AddRange(@($TMx64, $Anydesk))
$Line2aPanel.controls.AddRange(@($Attendance, $FileZilla, $Java))
$Line2bPanel.controls.AddRange(@($ChangeIP, $FortiClient, $Mozila))
$Line2cPanel.controls.AddRange(@($Chrome, $Inode, $Teams))
$Line3aPanel.controls.AddRange(@($Office, $RSIM, $UninstallJava))
$Line3bPanel.controls.AddRange(@($PdfOwnGuard, $Rufus, $Win10))
$Line3cPanel.controls.AddRange(@($Query64, $ActiveOffice, $Zoom))
$Line4aPanel.controls.AddRange(@($4000DW, $L350, $L405))
$Line4bPanel.controls.AddRange(@($T420, $L360, $LQ2190))
$Line4cPanel.controls.AddRange(@($T220, $L380, $L1800))
$Line5aPanel.controls.AddRange(@($TFData))
$Line5bPanel.controls.AddRange(@($InstallNet35))
$Line5cPanel.controls.AddRange(@($Restore))
$Line6aPanel.controls.AddRange(@($ShareFolder, $SharePrinter, $Other))

#================================================================================================================================

#Checkpoints
$Date = Get-Date -Format ddMMyyyy
Write-Output "Creating System Restore Point if one does not already exist. If one does, then you will receive a warning. Please wait..."
Checkpoint-Computer -Description $Date

Clear-Host
Show-Header

#================================================================================================================================

#region gui events {
$KES.Add_Click( 
    {            
        clear-host
        Show-Copyright
        $Apps = "Uninstall Kaspersky Endpoint Security"
        Write-Host "Mohon tunggu, sedang mencari GUID Kaspersky Network Agent"
        $Product = Get-WmiObject win32_product | where-Object { $_.name -eq "Kaspersky Security Center Network Agent" }
        Write-Host "Mohon tunggu, sedang Uninstall" $Product.name $Product.Version
        Start-Process "C:\Windows\System32\msiexec.exe" -ArgumentList "/x $($product.IdentifyingNumber) /quiet /noreboot" -Wait
        Write-Host "Process Uninstall" $Product.name "Berhasil"
        Write-Host "Mohon tunggu, sedang mencari GUID Kaspersky Endpoint Security"
        $Product = Get-WmiObject win32_product | where-object { $_.name -eq "Kaspersky Endpoint Security for Windows" }
        Write-Host "Mohon tunggu, sedang Uninstall" $Product.name $Product.Version
        Start-Process "C:\Windows\System32\msiexec.exe" -ArgumentList "/x $($product.IdentifyingNumber) /quiet /noreboot" -Wait
        Start-Process "C:\Windows\System32\msiexec.exe" -ArgumentList "/x $($product.IdentifyingNumber) KLLOGIN=KLAdmin KLPASSWD=P@ssw0rd /quiet /noreboot" -Wait
        clear-host
        Show-Footer
    } )

$SES.Add_Click( 
    {
        clear-host
        Show-Copyright
        $Apps = "Uninstall Symantec Endpoint Security"
        $Driver = "CleanWipe.exe"
        Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 6 MB" 
        invoke-webrequest -uri $LinkCleanWipedb -outfile $FileFolder\CleanWipe.db
        invoke-webrequest -uri $LinkCleanWipeexe -outfile $FileFolder\$Driver
        Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
        start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
        Show-Footer
    } )

$TMx64.Add_Click( 
    {
        clear-host	
        Show-Copyright
        $Apps = 'Install Trend Micro Apex One for 64bit'
        $Driver = 'ApexOne_x64.msi'
        Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 348.86 MB" 
        invoke-webrequest -uri $LinkTM64 -outfile $FileFolder\$Driver
        Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
        start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
        Show-Footer
    } )

	
$7zip.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Install 7zip for 64bit'
        $Driver = '7zip.exe'
        Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 1.31 MB" 
        invoke-webrequest -uri $Link7zip -outfile $FileFolder\$Driver
        Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
        Start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
        Show-Footer
    } )
               
$AdobeReader.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Install Adobe Reader XI for 64bit'
        $Driver = 'AdobeReader.exe'
        $App = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "Adobe*" }
        Write-Host "Mohon tunggu, sedang proses menguninstall file" $App.Name
        $App.uninstall()
        Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 72.34 MB" 
        invoke-webrequest -uri $LinkAdobe -outfile $FileFolder\$Driver
        Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
        start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
        Show-Footer
    } )
                
$Anydesk.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Install Anydesk'
        $Driver = 'Anydesk.exe'
        Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 3.50 MB" 
        invoke-webrequest -uri $LinkAnydesk -outfile $FileFolder\$Driver
        invoke-webrequest -uri $LinkVNC -outfile $FileFolder\Tightvnc.msi
                    
        Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
        $msiPath = "$FileFolder\Tightvnc.msi"
        $msiArgs = @(
            "/quiet",
            "/norestart",
            "ADDLOCAL=Server",
            "VIEWER_ASSOCIATE_VNC_EXTENSION=1",
            "SERVER_REGISTER_AS_SERVICE=1",
            "SERVER_ADD_FIREWALL_EXCEPTION=1",
            "VIEWER_ADD_FIREWALL_EXCEPTION=1",
            "SERVER_ALLOW_SAS=1",
            "SET_USEVNCAUTHENTICATION=1",
            "VALUE_OF_USEVNCAUTHENTICATION=1",
            "SET_PASSWORD=1",
            "VALUE_OF_PASSWORD=matahari",
            "SET_USECONTROLAUTHENTICATION=1",
            "VALUE_OF_USECONTROLAUTHENTICATION=1",
            "SET_CONTROLPASSWORD=1",
            "VALUE_OF_CONTROLPASSWORD=matahari",
            "SET_REMOVEWALLPAPER=1",
            "VALUE_OF_REMOVEWALLPAPER=0"
        )
Start-Process -FilePath "C:\Windows\System32\msiexec" -ArgumentList "/i $msiPath $msiArgs" -Wait
        cmd /c d:\SourceIT\anydesk.exe --install "C:\program files (x86)\AnyDesk" --start-with-win --create-shortcuts --create-desktop-icon --silent
        remove-item "D:\SourceIT\Tightvnc.msi" -force
        remove-item "D:\SourceIT\anydesk.exe" -force
        Show-Footer
    } )
                
$Attendance.Add_Click( 
    {
        clear-host
        Show-Copyright
        $Apps = 'Install Attendance Solution'
        $Driver = 'Attendance.zip'
        $Driver2 = 'Attendance.bat'
        Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 429.55 MB" 
        invoke-webrequest -uri $LinkAttendance -outfile $FileFolder\$Driver
        Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
        expand-archive -path $FileFolder\$Driver $FileFolder\Attendance
        start-process -FilePath $Driver2 -WorkingDirectory $FileFolder\Attendance -wait
        Show-Footer
    } )
               
$ChangeIP.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'ChangeIP'
        $Driver = 'ChangeIP.bat'
        remove-item -fo $FileFolder\$Driver
        if (Get-ChildItem -Path $FileFolder\$Driver -ErrorAction Ignore) {
            Write-Host "Mohon tunggu, sedang proses menjalankan file" $Apps 
            start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
            Show-Footer
        }
        else {
            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 23 kb" 
            invoke-webrequest -uri $LinkChangeIP -outfile $FileFolder\$Driver
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
            Show-Footer
        }
    } )

$Chrome.Add_Click( 
    {
        clear-host
        Show-Copyright
        $Apps = 'Install Google Chrome'
        $Driver = 'Chrome.exe'
            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 1.28 MB" 
            invoke-webrequest -uri $LinkChrome -outfile $FileFolder\$Driver
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
            Show-Footer
    } )
                
$FileZilla.Add_Click( 
    {
        clear-host
        Show-Copyright
        $Apps = 'Install FileZilla Client'
        $Driver = 'FileZilla.exe'
            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 10.80 MB" 
            invoke-webrequest -uri $LinkFilezilla -outfile $FileFolder\$Driver
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
            Show-Footer
    } )
                
$FortiClient.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Install FortiClient'
        $Driver = 'FortiClient.exe'
        $Choice = [Microsoft.VisualBasic.Interaction]::InputBox("Silahkan Input Number : 
    1. Remove FortiClient 
    2. Install FortiClient", "FortiClient", "2")
        if ($Choice -eq "1") {
            $App = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "Forti*" }
            $App.uninstall()
        }
        remove-item -path $FileFolder\$Driver -recurse 
        Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 91.11 MB" 
        invoke-webrequest -uri $LinkForti -outfile $FileFolder\$Driver
        Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
        start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
        Show-Footer
    } )
            
$Inode.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Install Inode'
        $Driver = 'Inode'
            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 18.27 MB" 
            invoke-webrequest -uri $LinkInode -outfile $FileFolder\$Driver
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
            Show-Footer
    } )
            
$Java.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Install Java 6u29'
        $Driver = 'Java_6u29.exe'
            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 16.12 MB" 
            invoke-webrequest -uri $LinkJava6u29 -outfile $FileFolder\$Driver
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
            Show-Footer
    } )
            
$Mozila.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Install Mozila Firefox'
        $Driver = 'Firefox.msi'
            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 57.548 kB" 
            invoke-webrequest -uri $LinkFirefox -outfile $FileFolder\$Driver
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
            Show-Footer
    } )
            
$Teams.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Install Microsoft Teams for Windows Desktop'
        $Driver = 'Teamssetup.exe'
            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 2 MB" 
            invoke-webrequest -uri $LinkTeams -outfile $FileFolder\$Driver
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
            Show-Footer
    } )
            
$Office.Add_Click( 
    {
        clear-host
        Show-Copyright
        $Apps = 'Install Microsoft Office 2013 for 64bit'
        $Driver = 'Office2013.zip'
        $Driver2 = 'setup.exe'
            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 667.54 MB" 
            invoke-webrequest -uri $Linkoffice -outfile $FileFolder\$Driver
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps
            expand-archive -path $FileFolder\$Driver $FileFolder\Office2013
            start-process -FilePath $Driver2 -WorkingDirectory $FileFolder\Office2013 -wait
            Show-Footerd
    } )
            
$PdfOwnGuard.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Install PDFOwnGuard For WIndows 10'
        $Driver = 'Pdfownguard.exe'
            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 5.90 MB" 
            invoke-webrequest -uri $Linkpdfownguard -outfile $FileFolder\$Driver
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
            Show-Footer
    } )
            
$Query64.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Install Microsoft Power Query For Excel x64'
        $Driver = 'PowerQueryx64.msi'

            $uninstall64 = gci "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | foreach { gp $_.PSPath } | ? { $_ -match "Query*" } | select UninstallString
            
            if ($uninstall64) {
                $uninstall64 = $uninstall64.UninstallString -Replace "msiexec.exe", "" -Replace "/I", "" -Replace "/X", ""
                $uninstall64 = $uninstall64.Trim()
                Write-Output "Uninstalling $Apps ..."
                start-process "msiexec.exe" -arg "/X $uninstall64 /qb" -Wait
            }
            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 17.74 MB" 
            invoke-webrequest -uri $Linkquery64 -outfile $FileFolder\$Driver
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
            Show-Footer
    } )
            
$RSIM.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Install RSIM'
        $Driver = 'RSIM.zip'

            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 208.12 MB" 
            invoke-webrequest -uri $Linkrsim -outfile $FileFolder\$Driver
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            expand-archive -path $FileFolder\$Driver $FileFolder\RSIM
            start-process -FilePath jdk_8u20.exe -WorkingDirectory $FileFolder\RSIM -wait
            start-process -FilePath gclientWindows10.4.9.exe -WorkingDirectory $FileFolder\RSIM -wait
            $WshShell = New-Object -comObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut("C:\Users\Public\Desktop\RSIM.lnk")
            $Shortcut.TargetPath = "C:\Program Files (x86)\clientWindows10.4.9\retek\sim\bin\rss.bat"
            $Shortcut.WorkingDirectory = "C:\Program Files (x86)\clientWindows10.4.9\retek\sim\bin"
            $Shortcut.Save()
            Show-Footer
    } )
            
$Rufus.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Running Rufus Create Bootable USB'
        $Driver = 'Rufus.exe'
            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 2 MB" 
            invoke-webrequest -uri $Linkrufus -outfile $FileFolder\$Driver
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
            Show-Footer
    } )
            
                
$ActiveOffice.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Actived Office 2013'
        # Specify cscript path
            $cscript = "C:\Windows\System32\cscript.exe"
            $ospp = "C:\Program Files\Microsoft Office\Office15\OSPP.VBS"

            # Check if OSPP.VBS exists
            if (!(Test-Path $ospp)) {
                Write-Host "OSPP.VBS not found. Please verify Office installation path."
                exit
            }

            # Get the current Office license status and extract the key
            $status = & $cscript $ospp /dstatus | Out-String
            $key = [regex]::Match($status, "Last 5 characters of installed product key: (.{5})").Groups[1].Value

            if ($key) {
                # Remove spaces from the key
                $key = $key -replace '\s',''
                
                Write-Host "Found key: $key"
                
                # Uninstall the current key
                Write-Host "Removing current key..."
                & $cscript $ospp /unpkey:$key
                
                # Install the new key
                Write-Host "Installing new key..."
                & $cscript $ospp /inpkey:3DYND-9DTQV-PRCDC-MCC8P-8W9MG
            } else {
                Write-Host "No product key found"
            }

            Show-Footer
    } )
                
            
$UninstallJava.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Uninstall Java Tools'
        $Driver = 'Javatools.exe'
            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 2 MB" 
            invoke-webrequest -uri $LinkJavaTools -outfile $FileFolder\$Driver
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
            Show-Footer
    } )
            
$Win10.Add_Click( 
    {
        clear-host
        Show-Copyright
        $Apps = 'Windows 10 Update Assistent'
        $Driver = 'Win10Upgrade.exe'
            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 4 MB" 
            invoke-webrequest -uri $LinkWin10 -outfile $FileFolder\$Driver
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
            Show-Footer
    } )
            
$Zoom.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Install Zoom Meeting'
        $Driver = 'Zoom.exe'

            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 67.19 MB" 
            invoke-webrequest -uri $LinkZoom -outfile $FileFolder\$Driver
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
            Show-Footer
    } )
            
$4000DW.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Install Printer Brother HL-T4000DW'
        $Driver = 'Brother_HL-T4000DW.exe'

            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 165.09 MB" 
            invoke-webrequest -uri $Linkbrother -outfile $FileFolder\$Driver
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
            Show-Footer
    } )
            
$T220.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Install Printer Brother T220'
        $Driver = 'Brother_T220.exe'

            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 218.56 MB" 
            invoke-webrequest -uri $Linkt220 -outfile $FileFolder\$Driver
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
            Show-Footer
    } )
            
$T420.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Install Printer Brother T420'
        $Driver = 'Brother_T420.exe'

            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 228.55 MB" 
            invoke-webrequest -uri $Linkt420 -outfile $FileFolder\$Driver
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
            Show-Footer

    } )
            
$LQ2190.Add_Click( 
    {
        clear-host
        Show-Copyright
        $Apps = 'Install Printer Epson LQ2190'
        $Driver = 'LQ2190.zip'
        $Driver = 'Setup.exe'

            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 2.54 MB" 
            invoke-webrequest -uri $Linklq2190 -outfile $FileFolder\$Driver
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            expand-archive -path $FileFolder\$Driver $FileFolder\LQ2190\Setup
            start-process -FilePath $Driver2 -WorkingDirectory $FileFolder\LQ2190\Setup -wait
            Show-Footer

    } )
            
$L350.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Install Printer Epson L350'
        $Driver = 'Printer_L350.exe'
        $Driver2 = 'Scanner_L350.exe'

            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 43.78 MB" 
            invoke-webrequest -uri $Linkl350printer	-outfile $FileFolder\$Driver
            invoke-webrequest -uri $Linkl350scanner	-outfile $FileFolder\$Driver2
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
            start-process -FilePath $Driver2 -WorkingDirectory $FileFolder -wait
            Show-Footer

    } )
            
$L360.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Install Printer Epson L360'
        $Driver = 'Printer_L360.exe'
        $Driver2 = 'Scanner_L360.exe'

            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 53.07 MB" 
            invoke-webrequest -uri $Linkl360printer	-outfile $FileFolder\$Driver
            invoke-webrequest -uri $Linkl360scanner	-outfile $FileFolder\$Driver2
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
            start-process -FilePath $Drive2 -WorkingDirectory $FileFolder -wait
            Show-Footer

    } )
            
$L380.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Install Printer Epson L380'
        $Driver = 'Printer_L380.exe'
        $Driver2 = 'Scanner_L380.exe'

            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 100.18 MB" 
            invoke-webrequest -uri $Linkl380printer	-outfile $FileFolder\$Driver
            invoke-webrequest -uri $Linkl380scanner	-outfile $FileFolder\$Driver2
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
            start-process -FilePath $Drive2 -WorkingDirectory $FileFolder -wait
            Show-Footer
    } )
                
$L405.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Install Printer Epson L405'
        $Driver = 'Printer_L405.exe'
        $Driver2 = 'Scanner_L405.exe'

            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 96.05 MB" 
            invoke-webrequest -uri $Linkl405printer -outfile $FileFolder\$Driver
            invoke-webrequest -uri $Linkl405scanner -outfile $FileFolder\$Driver
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
            start-process -FilePath $Driver2 -WorkingDirectory $FileFolder -wait
            Show-Footer
    } )
            
$L1800.Add_Click( 
    {
        clear-host
        Show-Copyright
        $Apps = 'Install Printer Epson L1800'
        $Driver = 'Printer_L1800.exe'

            Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 28.81 MB" 
            invoke-webrequest -uri $Linkl1800 -outfile $FileFolder\$Driver
            Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
            start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
            Show-Footer
                    
    } )
            
$TFData.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Copy Data Old User'
        $OldName = read-host -prompt 'Input User Login yang Lama ( example = mdsadm ) '
        $NewName = read-host -prompt 'Input User Login yang Baru ( example = Mxxxxxx ) '
        Write-Host "Mohon tunggu, sedang proses menjalankan Script"
        start-sleep 10
        xcopy c:\users\$OldName\Desktop\* /s /e /y c:\users\$NewName\Desktop
        xcopy c:\users\$OldName\Documents\* /s /e /y c:\users\$NewName\Documents
        xcopy c:\users\$OldName\Downloads\* /s /e /y c:\users\$NewName\Downloads
        xcopy c:\users\$OldName\Pictures\* /s /e /y c:\users\$NewName\Pictures
        xcopy c:\users\$OldName\Videos\* /s /e /y c:\users\$NewName\Videos
        Show-Footer
    } )
            
$Restore.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Restore Health for Windows 10'
        invoke-webrequest -uri $debloat -outfile $FileFolder\Windows10Debloater.ps1
        Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan Script" $Apps 
        powershell -executionpolicy ByPass -NoProfile -File "D:\SourceIT\Windows10Debloater.ps1"
        Write-Host "Dism Cleanup-Image StartComponentCleanup"
        Dism /Online /Cleanup-Image /StartComponentCleanup
        Write-Host "Dism Cleanup-Image RestoreHealth"
        Dism /Online /Cleanup-Image /RestoreHealth
        Write-Host "SFC /Scannow"
        SFC /scannow
        Remove-Item "D:\sourceIT\custom-lists.ps1" -force
        clear-host
        Show-Footer
        Write-Host "Silahkan close program and RESTART device, Thanksyou"
        Restart-Computer
    } )
                
$InstallNet35.Add_Click( 
    {
        clear-host 
        Show-Copyright
        $Apps = 'Installation of .Net Framework3.5'
        Write-Host "Initializing the installation of .NET Framework 3.5..."
        DISM /Online /Enable-Feature /FeatureName:NetFx3 /All
        Show-Footer
                    
    } )
            
$ShareFolder.Add_Click(
    {
        clear-host 
        Show-Copyright
        $Apps = "Sharing folder"
        $Choice_Sharing = [Microsoft.VisualBasic.Interaction]::InputBox("Masukkan Angka :  
        1. Untuk membuat share folder 
        2. Untuk men-connect-kan folder yang telah di sharing", "Create Sharing Folder", "2")
        if ($Choice_Sharing -eq "1") {
            Write-Host "Contoh Jika ingin mensharing folder Share di D:\Share"
            $Directory = [Microsoft.VisualBasic.Interaction]::InputBox("Masukkan Lokasi folder yang akan di share ( example = D:\Share ) :", "Sharing Folder")
            New-Item -path $Directory -itemType Directory -ErrorAction Ignore
            $NamaFolder = [Microsoft.VisualBasic.Interaction]::InputBox("Masukkan Nama folder yang akan di share ( example = Share ) :", "Sharing Folder" )
            New-SmbShare -Name $NamaFolder -Path $Directory -FullAccess 'Everyone'
            Show-Footer
            Write-Host "Sharing Folder $NamaFolder berhasil disharing"
        }
        else {
            $Sharing_PC = [Microsoft.VisualBasic.Interaction]::InputBox("Masukkan IP Computer yang telah disharing folder ( example = 172.xx.yy.31 ) :", "Sharing Folder")
            $Sharing_Nama = [Microsoft.VisualBasic.Interaction]::InputBox("Masukkan Nama folder yang telah dishare ( example = Share )  :", "Sharing Folder")
            New-PSDrive -Name "Z" -PSProvider "FileSystem" -Root "\\$Sharing_PC\$Sharing_Nama" -Persist  
            Show-Footer 
            Write-Host "Folder Share $NamaFolder berhasil Connected menjadi Drive Z"
        }

        
    }
)	
            
$SharePrinter.Add_Click(
    {
        clear-host 
        Show-Copyright
        $Apps = "Sharing Printer"
        Write-Host "Untuk Melihat full name printer dengan cara sbb :"
        Write-Host "1. Buka Powershell"
        Write-Host "2. Ketik Get-Printer"
        Write-Host "3. Jika sudah muncul list printer, silahkan input full name printer yang akan di sharing di message box"
                    
        $Printername = [Microsoft.VisualBasic.Interaction]::InputBox("Masukkan Full Nama Printer yang akan di sharing :", "Sharing Printer")
                            
        Set-Printer -Name $Printername -Shared $True -ShareName $Printername
        Clear-Host
        Show-Footer
        Write-Host "Sharing Printer $Printername berhasil disharing"
            
    }
)	
            
# ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
#other List Program
$other.Add_Click(
    {
            
        # This form was created using POSHGUI.com  a free online gui designer for PowerShell
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.Application]::EnableVisualStyles()
            
        $Form = New-Object system.Windows.Forms.Form
        $Form.ClientSize = New-Object System.Drawing.Point(600, 580)
        $Form.StartPosition = 'CenterScreen'
        $Form.FormBorderStyle = 'FixedSingle'
        $Form.MinimizeBox = $false
        $Form.MaximizeBox = $false
        $Form.ShowIcon = $false
        $Form.text = "Install Program Standart Tanpa Admin"
        $Form.TopMost = $false
        $Form.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#252525")
            
        $Line1aPanel = New-Object system.Windows.Forms.Panel
        $Line1aPanel.height = 120
        $Line1aPanel.width = 190
        $Line1aPanel.Anchor = 'top,right,left'
        $Line1aPanel.location = New-Object System.Drawing.Point(10, 10)
            
        $Line1bPanel = New-Object system.Windows.Forms.Panel
        $Line1bPanel.height = 120
        $Line1bPanel.width = 190
        $Line1bPanel.Anchor = 'top,right,left'
        $Line1bPanel.location = New-Object System.Drawing.Point(200, 10)
            
        $Line1cPanel = New-Object system.Windows.Forms.Panel
        $Line1cPanel.height = 120
        $Line1cPanel.width = 190
        $Line1cPanel.Anchor = 'top,right,left'
        $Line1cPanel.location = New-Object System.Drawing.Point(390, 10)
            
        $Line2aPanel = New-Object system.Windows.Forms.Panel
        $Line2aPanel.height = 120
        $Line2aPanel.width = 190
        $Line2aPanel.Anchor = 'top,right,left'
        $Line2aPanel.location = New-Object System.Drawing.Point(10, 120)
            
        $Line2bPanel = New-Object system.Windows.Forms.Panel
        $Line2bPanel.height = 120
        $Line2bPanel.width = 190
        $Line2bPanel.Anchor = 'top,right,left'
        $Line2bPanel.location = New-Object System.Drawing.Point(200, 120)
            
        $Line2cPanel = New-Object system.Windows.Forms.Panel
        $Line2cPanel.height = 120
        $Line2cPanel.width = 190
        $Line2cPanel.Anchor = 'top,right,left'
        $Line2cPanel.location = New-Object System.Drawing.Point(390, 120)
            
        $Line3aPanel = New-Object system.Windows.Forms.Panel
        $Line3aPanel.height = 120
        $Line3aPanel.width = 190
        $Line3aPanel.Anchor = 'top,right,left'
        $Line3aPanel.location = New-Object System.Drawing.Point(10, 240)
            
        $Line3bPanel = New-Object system.Windows.Forms.Panel
        $Line3bPanel.height = 120
        $Line3bPanel.width = 190
        $Line3bPanel.Anchor = 'top,right,left'
        $Line3bPanel.location = New-Object System.Drawing.Point(200, 240)
            
        $Line3cPanel = New-Object system.Windows.Forms.Panel
        $Line3cPanel.height = 120
        $Line3cPanel.width = 190
        $Line3cPanel.Anchor = 'top,right,left'
        $Line3cPanel.location = New-Object System.Drawing.Point(390, 240)
            
        $Line4aPanel = New-Object system.Windows.Forms.Panel
        $Line4aPanel.height = 120
        $Line4aPanel.width = 190
        $Line4aPanel.Anchor = 'top,right,left'
        $Line4aPanel.location = New-Object System.Drawing.Point(10, 360)
            
        $Line4bPanel = New-Object system.Windows.Forms.Panel
        $Line4bPanel.height = 120
        $Line4bPanel.width = 190
        $Line4bPanel.Anchor = 'top,right,left'
        $Line4bPanel.location = New-Object System.Drawing.Point(200, 360)
            
        $Line4cPanel = New-Object system.Windows.Forms.Panel
        $Line4cPanel.height = 120
        $Line4cPanel.width = 190
        $Line4cPanel.Anchor = 'top,right,left'
        $Line4cPanel.location = New-Object System.Drawing.Point(390, 360)
            
        $Line5aPanel = New-Object system.Windows.Forms.Panel
        $Line5aPanel.height = 120
        $Line5aPanel.width = 190
        $Line5aPanel.Anchor = 'top,right,left'
        $Line5aPanel.location = New-Object System.Drawing.Point(10, 400)
            
        $Line5bPanel = New-Object system.Windows.Forms.Panel
        $Line5bPanel.height = 120
        $Line5bPanel.width = 190
        $Line5bPanel.Anchor = 'top,right,left'
        $Line5bPanel.location = New-Object System.Drawing.Point(200, 400)
            
        $Line5cPanel = New-Object system.Windows.Forms.Panel
        $Line5cPanel.height = 120
        $Line5cPanel.width = 190
        $Line5cPanel.Anchor = 'top,right,left'
        $Line5cPanel.location = New-Object System.Drawing.Point(390, 400)
            
        $Line6aPanel = New-Object system.Windows.Forms.Panel
        $Line6aPanel.height = 120
        $Line6aPanel.width = 600
        $Line6aPanel.Anchor = 'top,right,left'
        $Line6aPanel.location = New-Object System.Drawing.Point(10, 530)
            
        $List = New-Object system.Windows.Forms.Label
        $List.text = "LIST PROGRAM NEW"
        $List.AutoSize = $true
        $List.width = 457
        $List.height = 560
        $List.Anchor = 'top,right,left'
        $List.location = New-Object System.Drawing.Point(10, 10)
        $List.Font = New-Object System.Drawing.Font('ALGERIAN', 13, [System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
        $List.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $WPS = New-Object system.Windows.Forms.Button
        $WPS.FlatStyle = 'Flat'
        $WPS.text = "Uninstall WPS Office"
        $WPS.width = 180
        $WPS.height = 30
        $WPS.Anchor = 'top,right,left'
        $WPS.location = New-Object System.Drawing.Point(10, 40)
        $WPS.Font = New-Object System.Drawing.Font('Consolas', 9)
        $WPS.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $WinUpdate = New-Object system.Windows.Forms.Button
        $WinUpdate.FlatStyle = 'Flat'
        $WinUpdate.text = "Windows Update Manually"
        $WinUpdate.width = 180
        $WinUpdate.height = 30
        $WinUpdate.Anchor = 'top,right,left'
        $WinUpdate.location = New-Object System.Drawing.Point(10, 40)
        $WinUpdate.Font = New-Object System.Drawing.Font('Consolas', 9)
        $WinUpdate.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Bookmark = New-Object system.Windows.Forms.Button
        $Bookmark.FlatStyle = 'Flat'
        $Bookmark.text = "Bookmark Google Chrome"
        $Bookmark.width = 180
        $Bookmark.height = 30
        $Bookmark.Anchor = 'top,right,left'
        $Bookmark.location = New-Object System.Drawing.Point(10, 40)
        $Bookmark.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Bookmark.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $ETPtraining = New-Object system.Windows.Forms.Button
        $ETPtraining.FlatStyle = 'Flat'
        $ETPtraining.text = " "
        $ETPtraining.width = 180
        $ETPtraining.height = 30
        $ETPtraining.Anchor = 'top,right,left'
        $ETPtraining.location = New-Object System.Drawing.Point(10, 80)
        $ETPtraining.Font = New-Object System.Drawing.Font('Consolas', 9)
        $ETPtraining.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $CheckDisk = New-Object system.Windows.Forms.Button
        $CheckDisk.FlatStyle = 'Flat'
        $CheckDisk.text = "Check and Repair Disk"
        $CheckDisk.width = 180
        $CheckDisk.height = 30
        $CheckDisk.Anchor = 'top,right,left'
        $CheckDisk.location = New-Object System.Drawing.Point(10, 80)
        $CheckDisk.Font = New-Object System.Drawing.Font('Consolas', 9)
        $CheckDisk.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $CleanUp = New-Object system.Windows.Forms.Button
        $CleanUp.FlatStyle = 'Flat'
        $CleanUp.text = "CleanUp Disk"
        $CleanUp.width = 180
        $CleanUp.height = 30
        $CleanUp.Anchor = 'top,right,left'
        $CleanUp.location = New-Object System.Drawing.Point(10, 80)
        $CleanUp.Font = New-Object System.Drawing.Font('Consolas', 9)
        $CleanUp.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Shortcut = New-Object system.Windows.Forms.Button
        $Shortcut.FlatStyle = 'Flat'
        $Shortcut.text = "Shortcut Applikasi"
        $Shortcut.width = 180
        $Shortcut.height = 30
        $Shortcut.Anchor = 'top,right,left'
        $Shortcut.location = New-Object System.Drawing.Point(10, 10)
        $Shortcut.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Shortcut.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Uninstall16 = New-Object system.Windows.Forms.Button
        $Uninstall16.FlatStyle = 'Flat'
        $Uninstall16.text = "Uninstall Office 2016"
        $Uninstall16.width = 180
        $Uninstall16.height = 30
        $Uninstall16.Anchor = 'top,right,left'
        $Uninstall16.location = New-Object System.Drawing.Point(10, 10)
        $Uninstall16.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Uninstall16.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $SiteFz = New-Object system.Windows.Forms.Button
        $SiteFz.FlatStyle = 'Flat'
        $SiteFz.text = "Setting Site FileZilla"
        $SiteFz.width = 180
        $SiteFz.height = 30
        $SiteFz.Anchor = 'top,right,left'
        $SiteFz.location = New-Object System.Drawing.Point(10, 10)
        $SiteFz.Font = New-Object System.Drawing.Font('Consolas', 9)
        $SiteFz.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $ChangePartition = New-Object system.Windows.Forms.Button
        $ChangePartition.FlatStyle = 'Flat'
        $ChangePartition.text = "Change Drive Partition"
        $ChangePartition.width = 180
        $ChangePartition.height = 30
        $ChangePartition.Anchor = 'top,right,left'
        $ChangePartition.location = New-Object System.Drawing.Point(10, 50)
        $ChangePartition.Font = New-Object System.Drawing.Font('Consolas', 9)
        $ChangePartition.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $L220 = New-Object system.Windows.Forms.Button
        $L220.FlatStyle = 'Flat'
        $L220.text = "Printer Epson L220"
        $L220.width = 180
        $L220.height = 30
        $L220.Anchor = 'top,right,left'
        $L220.location = New-Object System.Drawing.Point(10, 50)
        $L220.Font = New-Object System.Drawing.Font('Consolas', 9)
        $L220.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Button1 = New-Object system.Windows.Forms.Button
        $Button1.FlatStyle = 'Flat'
        $Button1.text = "Uninstall Driver LAN"
        $Button1.width = 180
        $Button1.height = 30
        $Button1.Anchor = 'top,right,left'
        $Button1.location = New-Object System.Drawing.Point(10, 50)
        $Button1.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Button1.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Tzone = New-Object system.Windows.Forms.Button
        $Tzone.FlatStyle = 'Flat'
        $Tzone.text = "Change TimeZone"
        $Tzone.width = 180
        $Tzone.height = 30
        $Tzone.Anchor = 'top,right,left'
        $Tzone.location = New-Object System.Drawing.Point(10, 90)
        $Tzone.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Tzone.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $PowerBI = New-Object system.Windows.Forms.Button
        $PowerBI.FlatStyle = 'Flat'
        $PowerBI.text = "Power BI"
        $PowerBI.width = 180
        $PowerBI.height = 30
        $PowerBI.Anchor = 'top,right,left'
        $PowerBI.location = New-Object System.Drawing.Point(10, 90)
        $PowerBI.Font = New-Object System.Drawing.Font('Consolas', 9)
        $PowerBI.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $VNC = New-Object system.Windows.Forms.Button
        $VNC.FlatStyle = 'Flat'
        $VNC.text = "Tvnviewer"
        $VNC.width = 180
        $VNC.height = 30
        $VNC.Anchor = 'top,right,left'
        $VNC.location = New-Object System.Drawing.Point(10, 90)
        $VNC.Font = New-Object System.Drawing.Font('Consolas', 9)
        $VNC.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $IE_mode = New-Object system.Windows.Forms.Button
        $IE_mode.FlatStyle = 'Flat'
        $IE_mode.text = "Shortcut IE Mode"
        $IE_mode.width = 180
        $IE_mode.height = 30
        $IE_mode.Anchor = 'top,right,left'
        $IE_mode.location = New-Object System.Drawing.Point(10, 10)
        $IE_mode.Font = New-Object System.Drawing.Font('Consolas', 9)
        $IE_mode.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Join_Domain = New-Object system.Windows.Forms.Button
        $Join_Domain.FlatStyle = 'Flat'
        $Join_Domain.text = "Join Domain"
        $Join_Domain.width = 180
        $Join_Domain.height = 30
        $Join_Domain.Anchor = 'top,right,left'
        $Join_Domain.location = New-Object System.Drawing.Point(10, 10)
        $Join_Domain.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Join_Domain.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Change_Hostname = New-Object system.Windows.Forms.Button
        $Change_Hostname.FlatStyle = 'Flat'
        $Change_Hostname.text = "Change Hostname"
        $Change_Hostname.width = 180
        $Change_Hostname.height = 30
        $Change_Hostname.Anchor = 'top,right,left'
        $Change_Hostname.location = New-Object System.Drawing.Point(10, 10)
        $Change_Hostname.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Change_Hostname.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Button8 = New-Object system.Windows.Forms.Button
        $Button8.FlatStyle = 'Flat'
        $Button8.text = ""
        $Button8.width = 180
        $Button8.height = 30
        $Button8.Anchor = 'top,right,left'
        $Button8.location = New-Object System.Drawing.Point(10, 50)
        $Button8.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Button8.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Button9 = New-Object system.Windows.Forms.Button
        $Button9.FlatStyle = 'Flat'
        $Button9.text = ""
        $Button9.width = 180
        $Button9.height = 30
        $Button9.Anchor = 'top,right,left'
        $Button9.location = New-Object System.Drawing.Point(10, 50)
        $Button9.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Button9.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Button10 = New-Object system.Windows.Forms.Button
        $Button10.FlatStyle = 'Flat'
        $Button10.text = ""
        $Button10.width = 180
        $Button10.height = 30
        $Button10.Anchor = 'top,right,left'
        $Button10.location = New-Object System.Drawing.Point(10, 50)
        $Button10.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Button10.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Button11 = New-Object system.Windows.Forms.Button
        $Button11.FlatStyle = 'Flat'
        $Button11.text = ""
        $Button11.width = 180
        $Button11.height = 30
        $Button11.Anchor = 'top,right,left'
        $Button11.location = New-Object System.Drawing.Point(10, 90)
        $Button11.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Button11.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Button12 = New-Object system.Windows.Forms.Button
        $Button12.FlatStyle = 'Flat'
        $Button12.text = ""
        $Button12.width = 180
        $Button12.height = 30
        $Button12.Anchor = 'top,right,left'
        $Button12.location = New-Object System.Drawing.Point(10, 90)
        $Button12.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Button12.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Button13 = New-Object system.Windows.Forms.Button
        $Button13.FlatStyle = 'Flat'
        $Button13.text = ""
        $Button13.width = 180
        $Button13.height = 30
        $Button13.Anchor = 'top,right,left'
        $Button13.location = New-Object System.Drawing.Point(10, 90)
        $Button13.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Button13.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Button14 = New-Object system.Windows.Forms.Button
        $Button14.FlatStyle = 'Flat'
        $Button14.text = ""
        $Button14.width = 180
        $Button14.height = 30
        $Button14.Anchor = 'top,right,left'
        $Button14.location = New-Object System.Drawing.Point(10, 10)
        $Button14.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Button14.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Button15 = New-Object system.Windows.Forms.Button
        $Button15.FlatStyle = 'Flat'
        $Button15.text = ""
        $Button15.width = 180
        $Button15.height = 30
        $Button15.Anchor = 'top,right,left'
        $Button15.location = New-Object System.Drawing.Point(10, 10)
        $Button15.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Button15.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Button16 = New-Object system.Windows.Forms.Button
        $Button16.FlatStyle = 'Flat'
        $Button16.text = ""
        $Button16.width = 180
        $Button16.height = 30
        $Button16.Anchor = 'top,right,left'
        $Button16.location = New-Object System.Drawing.Point(10, 10)
        $Button16.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Button16.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Button17 = New-Object system.Windows.Forms.Button
        $Button17.FlatStyle = 'Flat'
        $Button17.text = ""
        $Button17.width = 180
        $Button17.height = 30
        $Button17.Anchor = 'top,right,left'
        $Button17.location = New-Object System.Drawing.Point(10, 50)
        $Button17.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Button17.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Button18 = New-Object system.Windows.Forms.Button
        $Button18.FlatStyle = 'Flat'
        $Button18.text = ""
        $Button18.width = 180
        $Button18.height = 30
        $Button18.Anchor = 'top,right,left'
        $Button18.location = New-Object System.Drawing.Point(10, 50)
        $Button18.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Button18.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Button19 = New-Object system.Windows.Forms.Button
        $Button19.FlatStyle = 'Flat'
        $Button19.text = ""
        $Button19.width = 180
        $Button19.height = 30
        $Button19.Anchor = 'top,right,left'
        $Button19.location = New-Object System.Drawing.Point(10, 50)
        $Button19.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Button19.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Tzone0 = New-Object system.Windows.Forms.Button
        $Tzone0.FlatStyle = 'Flat'
        $Tzone0.text = ""
        $Tzone0.width = 180
        $Tzone0.height = 30
        $Tzone0.Anchor = 'top,right,left'
        $Tzone0.location = New-Object System.Drawing.Point(10, 90)
        $Tzone0.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Tzone0.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Tzone1 = New-Object system.Windows.Forms.Button
        $Tzone1.FlatStyle = 'Flat'
        $Tzone1.text = ""
        $Tzone1.width = 180
        $Tzone1.height = 30
        $Tzone1.Anchor = 'top,right,left'
        $Tzone1.location = New-Object System.Drawing.Point(10, 90)
        $Tzone1.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Tzone1.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Tzone2 = New-Object system.Windows.Forms.Button
        $Tzone2.FlatStyle = 'Flat'
        $Tzone2.text = ""
        $Tzone2.width = 180
        $Tzone2.height = 30
        $Tzone2.Anchor = 'top,right,left'
        $Tzone2.location = New-Object System.Drawing.Point(10, 90)
        $Tzone2.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Tzone2.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Tzone3 = New-Object system.Windows.Forms.Button
        $Tzone3.FlatStyle = 'Flat'
        $Tzone3.text = ""
        $Tzone3.width = 180
        $Tzone3.height = 30
        $Tzone3.Anchor = 'top,right,left'
        $Tzone3.location = New-Object System.Drawing.Point(10, 90)
        $Tzone3.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Tzone3.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Tzone4 = New-Object system.Windows.Forms.Button
        $Tzone4.FlatStyle = 'Flat'
        $Tzone4.text = ""
        $Tzone4.width = 180
        $Tzone4.height = 30
        $Tzone4.Anchor = 'top,right,left'
        $Tzone4.location = New-Object System.Drawing.Point(10, 90)
        $Tzone4.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Tzone4.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Tzone5 = New-Object system.Windows.Forms.Button
        $Tzone5.FlatStyle = 'Flat'
        $Tzone5.text = ""
        $Tzone5.width = 180
        $Tzone5.height = 30
        $Tzone5.Anchor = 'top,right,left'
        $Tzone5.location = New-Object System.Drawing.Point(10, 90)
        $Tzone5.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Tzone5.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Tzone6 = New-Object system.Windows.Forms.Button
        $Tzone6.FlatStyle = 'Flat'
        $Tzone6.text = ""
        $Tzone6.width = 180
        $Tzone6.height = 30
        $Tzone6.Anchor = 'top,right,left'
        $Tzone6.location = New-Object System.Drawing.Point(10, 0)
        $Tzone6.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Tzone6.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Tzone7 = New-Object system.Windows.Forms.Button
        $Tzone7.FlatStyle = 'Flat'
        $Tzone7.text = ""
        $Tzone7.width = 180
        $Tzone7.height = 30
        $Tzone7.Anchor = 'top,right,left'
        $Tzone7.location = New-Object System.Drawing.Point(200, 0)
        $Tzone7.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Tzone7.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Tzone8 = New-Object system.Windows.Forms.Button
        $Tzone8.FlatStyle = 'Flat'
        $Tzone8.text = ""
        $Tzone8.width = 180
        $Tzone8.height = 30
        $Tzone8.Anchor = 'top,right,left'
        $Tzone8.location = New-Object System.Drawing.Point(390, 0)
        $Tzone8.Font = New-Object System.Drawing.Font('Consolas', 9)
        $Tzone8.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#eeeeee")
            
        $Form.controls.AddRange(@($Line1aPanel, $Line1bPanel, $Line1cPanel, $Line2aPanel, $Line2aPanel, $Line2bPanel, $line2cPanel, $Line3aPanel, $Line3bPanel, $Line3cPanel, $Line4aPanel, $Line4bPanel, $Line4cPanel, $Line5aPanel, $Line5bPanel, $Line5cPanel, $Line6aPanel))
        $Line1aPanel.controls.AddRange(@($List, $WPS, $ETPtraining))
        $Line1bPanel.controls.AddRange(@($WinUpdate, $CheckDisk))
        $Line1cPanel.controls.AddRange(@($Bookmark, $CleanUp))
        $Line2aPanel.controls.AddRange(@($Shortcut, $ChangePartition, $Tzone))
        $Line2bPanel.controls.AddRange(@($Uninstall16, $L220, $PowerBI))
        $Line2cPanel.controls.AddRange(@($SiteFz, $Button1, $VNC))
        $Line3aPanel.controls.AddRange(@($IE_mode, $Button8, $Button11))
        $Line3bPanel.controls.AddRange(@($Join_Domain, $Button9, $Button12))
        $Line3cPanel.controls.AddRange(@($Change_Hostname, $Button10, $Button13))
        $Line4aPanel.controls.AddRange(@($Button14, $Button17, $Tzone0))
        $Line4bPanel.controls.AddRange(@($Button16, $Button18, $Tzone1))
        $Line4cPanel.controls.AddRange(@($Button15, $Button19, $Tzone2))
        $Line5aPanel.controls.AddRange(@($Tzone3))
        $Line5bPanel.controls.AddRange(@($Tzone4))
        $Line5cPanel.controls.AddRange(@($Tzone5))
        $Line6aPanel.controls.AddRange(@($Tzone6, $Tzone7, $Tzone8))
            
        clear-host 
        Show-Header
            
        #region gui events {
        $WPS.Add_Click( 
            {
                clear-host 
                Show-Copyright
                $Apps = "Uninstall WPS"
                Write-Host "Mohon tunggu, sedang menguninstall Program WPS"
                $App = Get-ChildItem -Path HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall | Get-ItemProperty | Where-Object { $_.Publisher -like 'Kingsoft*' } | Select-Object -Property DisplayName, UninstallString

                foreach ($Ver in $App) {
                    if ($Ver.UninstallString) {
                        $DisplayName = $Ver.DisplayName
                        $Uninst = $Ver.UninstallString
                        Write-Output "Uninstalling $DisplayName..."
                        cmd /c $Uninst /s
                    }
                }
                Show-Footer
            } )                
        $WinUpdate.Add_Click( 
            {
                clear-host 
                Show-Copyright
                $Apps = "Install Windows Update Manually"
                Write-Host "Mohon tunggu, Sedang pencarian patch update"
                Set-ExecutionPolicy RemoteSigned -Force
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
                Install-Module PSWindowsUpdate -Force 
                Import-Module -Name PSWindowsUpdate -Force
                cd HKLM:
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update" -Name "AutoUpdateScanPolicy" -Value 2
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "NoAutoUpdate" -Value 0
                cd C:
                Restart-Service -Name "wuauserv"
                Get-WindowsUpdate
                Write-Host "Process Download & install Windows Update"
                Install-WindowsUpdate -AcceptAll -AutoReboot
                Show-Footer
                write-Host "Windows Update berhasil, PC akan restart Otomatis"
            } )
        $Bookmark.Add_Click(
            {
                clear-host 
                Show-Copyright
                $Apps = "Create Bookmarks on Google Chrome"
                Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 1.31 MB" 
                invoke-webrequest -uri $BookmarkChrome -outfile C:\tmp\Bookmarks
                Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses " $Apps 
                Copy-Item -Path "C:\tmp\Bookmarks*" -Destination "$Loc_Appdata\Local\Google\Chrome\user data\default"  -Recurse
                Remove-Item -Path "C:\tmp\Bookmarks*" -Recurse -Force
                Show-Footer
            }
        )
            
        $ETPtraining.Add_Click( 
            {
                clear-host
                Show-Copyright
             } )
               
        $CheckDisk.Add_Click({ 
                Clear-Host 
                Show-Copyright
                $Apps = "Check Disk"
                
                Write-Host "Checking disk health status..." -ForegroundColor Yellow
                
                # Get all physical disks
                $disks = Get-PhysicalDisk
                
                if ($disks) {
                    foreach ($disk in $disks) {
                        Write-Host "`nChecking Disk $($disk.DeviceID)" -ForegroundColor Cyan
                        Write-Host "Health Status: $($disk.HealthStatus)" -ForegroundColor Yellow
                        
                        if ($disk.HealthStatus -ne "Healthy") {
                            Write-Host "Issues detected on Disk $($disk.DeviceID). Running CHKDSK..." -ForegroundColor Red
                            try {
                                Start-Process -FilePath "chkdsk.exe" -ArgumentList "/f /r /x $($disk.DeviceID)" -Wait -NoNewWindow
                                Write-Host "CHKDSK completed for Disk $($disk.DeviceID)" -ForegroundColor Green
                            } catch {
                                Write-Host "Error running CHKDSK on Disk $($disk.DeviceID): $_" -ForegroundColor Red
                            }
                        } else {
                            Write-Host "Disk $($disk.DeviceID) is healthy. No CHKDSK needed." -ForegroundColor Green
                        }
                        
                        Write-Host "Model: $($disk.Model)" -ForegroundColor Gray
                        Write-Host "Size: $([math]::Round($disk.Size/1GB, 2)) GB" -ForegroundColor Gray
                    }
                } else {
                    Write-Host "No physical disks found!" -ForegroundColor Red
                }
                
                Write-Host "`nDisk check completed." -ForegroundColor Green
                Show-Footer
            })
                
        $CleanUp.Add_Click( 
            {
                clear-host 
                Show-Copyright
                $Apps = "Deep Clean Up"
                # Requires -RunAsAdministrator

                function Show-Banner {
                    Write-Host "=================================" -ForegroundColor Cyan
                    Write-Host "       Windows Deep Cleanup       " -ForegroundColor Cyan
                    Write-Host "=================================" -ForegroundColor Cyan
                    Write-Host ""
                }

                function Show-Footer {
                    Write-Host ""
                    Write-Host "=================================" -ForegroundColor Cyan
                    Write-Host "         Cleanup Complete         " -ForegroundColor Cyan
                    Write-Host "=================================" -ForegroundColor Cyan
                }

                function Start-DeepCleanup {
                    # Get current user
                    try {
                        $currentUser = (Get-WmiObject -Class Win32_ComputerSystem).Username
                        Write-Host "Running cleanup for user: $currentUser" -ForegroundColor Green
                    }
                    catch {
                        Write-Warning "Could not determine current user. Continuing anyway..."
                    }

                    # Define cleanup paths with environment variables
                    $Loc_Appdata = $env:LOCALAPPDATA
                    $cleanupPaths = @(
                        "$env:windir\Temp\*"
                        "$Loc_Appdata\Temp\*"
                        "$env:windir\Prefetch\*"
                        "$env:windir\SoftwareDistribution\Download\*"
                        "$env:windir\Logs\CBS\*"
                    )

                    # Basic file cleanup
                    Write-Host "`nPerforming basic file cleanup..." -ForegroundColor Yellow
                    foreach ($path in $cleanupPaths) {
                        try {
                            if (Test-Path $path) {
                                Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
                                Write-Host "Successfully cleaned: $path" -ForegroundColor Green
                            }
                        }
                        catch {
                            Write-Warning "Failed to clean: $path`nError: $($_.Exception.Message)"
                        }
                    }

                    # Advanced Disk Cleanup
                    Write-Host "`nInitiating Advanced Disk Cleanup..." -ForegroundColor Yellow
                    $SageSet = "StateFlags0099"
                    $Base = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\"
                    
                    $Locations = @(
                        "Active Setup Temp Folders"
                        "BranchCache"
                        "Downloaded Program Files"
                        "GameNewsFiles"
                        "GameStatisticsFiles"
                        "GameUpdateFiles"
                        "Internet Cache Files"
                        "Memory Dump Files"
                        "Offline Pages Files"
                        "Old ChkDsk Files"
                        "D3D Shader Cache"
                        "Delivery Optimization Files"
                        "Diagnostic Data Viewer database files"
                        "Service Pack Cleanup"
                        "Setup Log Files"
                        "System error memory dump files"
                        "System error minidump files"
                        "Temporary Files"
                        "Temporary Setup Files"
                        "Temporary Sync Files"
                        "Thumbnail Cache"
                        "Update Cleanup"
                        "Upgrade Discarded Files"
                        "User file versions"
                        "Windows Defender"
                        "Windows Error Reporting Archive Files"
                        "Windows Error Reporting Queue Files"
                        "Windows Error Reporting System Archive Files"
                        "Windows Error Reporting System Queue Files"
                        "Windows ESD installation files"
                        "Windows Upgrade Log Files"
                    )

                    # Set cleanup flags
                    Write-Host "Setting cleanup flags..." -ForegroundColor Yellow
                    foreach ($Location in $Locations) {
                        try {
                            Set-ItemProperty -Path $($Base + $Location) -Name $SageSet -Type DWORD -Value 2 -ErrorAction SilentlyContinue
                        }
                        catch {
                            Write-Warning "Failed to set flag for: $Location"
                        }
                    }

                    # Run Disk Cleanup
                    try {
                        Write-Host "Running Disk Cleanup utility..." -ForegroundColor Yellow
                        $CleanupArgs = "/sagerun:$([string]([int]$SageSet.Substring($SageSet.Length-4)))"
                        Start-Process -Wait "$env:SystemRoot\System32\cleanmgr.exe" -ArgumentList $CleanupArgs
                        Write-Host "Disk Cleanup completed successfully" -ForegroundColor Green
                    }
                    catch {
                        Write-Warning "Disk Cleanup encountered an error: $($_.Exception.Message)"
                    }

                    # Cleanup Registry Flags
                    Write-Host "`nCleaning up registry flags..." -ForegroundColor Yellow
                    foreach ($Location in $Locations) {
                        try {
                            Remove-ItemProperty -Path $($Base + $Location) -Name $SageSet -Force -ErrorAction SilentlyContinue
                        }
                        catch {
                            Write-Warning "Failed to remove registry flag for: $Location"
                        }
                    }
                }

                # Main execution
                try {
                    # Check for admin privileges
                    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
                    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
                        throw "This script requires administrator privileges. Please run as administrator."
                    }

                    Show-Banner
                    Start-DeepCleanup
                    Show-Footer
                }
                catch {
                    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
                    exit 1
                }
            } )
                
        $Shortcut.Add_Click( 
            {
                clear-host 
                Show-Copyright
                if ($Octet_4 -eq 31) {
                    #Shortcut Monthly Insentive
                    $Url_Apk = "http://192.168.6.92/forms/frmservlet?config=spapps"
                    $Nama_Apk = "Monthly Incentive"
                   
                    $NameApk = [System.String]::Concat("C:\Program Files (x86)\Internet Explorer\", $Nama_Apk, ".vbs")
                    invoke-webrequest -uri $IEMode -outfile $nameapk
                    $file = $nameapk
    
                    $searchText = @{
                        "about:blank" = $Url_APk
                    }
                
                    # Membaca isi file ke dalam variabel
                    $content = Get-Content $file
                
                    # Mengganti teks tertentu dalam isi file
                    $searchText.GetEnumerator() | ForEach-Object {
                        $content = $content -replace $_.Key, $_.Value
                    }
                
                    # Menulis isi file yang telah diubah ke dalam file
                    Set-Content $file $content
                
    
                    $WshShell = New-Object -comObject WScript.Shell
                    $targetPath = $nameapk
                    $iconLocation = "C:\Program Files (x86)\Internet Explorer\iexplore.exe"
                    $iconFile = "IconFile=" + $iconLocation
                    $path = "$loc_users\$Nama_Apk.url"
                    $Shortcut = $WshShell.CreateShortcut($path)
                    $Shortcut.TargetPath = $targetPath
                    $Shortcut.Save()
                
                    Add-Content $path "HotKey=0"
                    Add-Content $path "$iconfile"
                    Add-Content $path "IconIndex=0"
    
                    $fileAttributes = (Get-Item $nameapk).Attributes
                    $fileAttributes += [System.IO.FileAttributes]::Hidden
                    Set-ItemProperty -Path $nameapk -Name Attributes -Value $fileAttributes
                    clear-host
                    write-Host "Shortcut Monthly Insentive Berhasil dibuat, Silahkan cek di desktop"
                    Show-Footer
                }

                elseif ($Octet_4 -eq 36) {
                    $Apps = "Shortcut for ASM"
                    #Shortcut Monthly Insentive
                    $Url_Apk = "http://192.168.6.92/forms/frmservlet?config=spapps"
                    $Nama_Apk = "Monthly Incentive"

                    $NameApk = [System.String]::Concat("C:\Program Files (x86)\Internet Explorer\", $Nama_Apk, ".vbs")
                    invoke-webrequest -uri $IEMode -outfile $nameapk
                    $file = $nameapk
    
                    $searchText = @{
                        "about:blank" = $Url_APk
                    }
                
                    # Membaca isi file ke dalam variabel
                    $content = Get-Content $file
                
                    # Mengganti teks tertentu dalam isi file
                    $searchText.GetEnumerator() | ForEach-Object {
                        $content = $content -replace $_.Key, $_.Value
                    }
                
                    # Menulis isi file yang telah diubah ke dalam file
                    Set-Content $file $content
                
    
                    $WshShell = New-Object -comObject WScript.Shell
                    $targetPath = $nameapk
                    $iconLocation = "C:\Program Files (x86)\Internet Explorer\iexplore.exe"
                    $iconFile = "IconFile=" + $iconLocation
                    $path = "$loc_users\$Nama_Apk.url"
                    $Shortcut = $WshShell.CreateShortcut($path)
                    $Shortcut.TargetPath = $targetPath
                    $Shortcut.Save()
                
                    Add-Content $path "HotKey=0"
                    Add-Content $path "$iconfile"
                    Add-Content $path "IconIndex=0"
    
                    $fileAttributes = (Get-Item $nameapk).Attributes
                    $fileAttributes += [System.IO.FileAttributes]::Hidden
                    Set-ItemProperty -Path $nameapk -Name Attributes -Value $fileAttributes
                    clear-host
                    write-Host "Shortcut Monthly Insentive Berhasil dibuat, Silahkan cek di desktop"
            
                    #Shortcut PPK
                    $Url_Apk = "http://mdsapps:7070/sysHRD/en/neoclassic/login/login"
                    $Nama_Apk = "PPK"

                    $NameApk = [System.String]::Concat("C:\Program Files (x86)\Internet Explorer\", $Nama_Apk, ".vbs")
                    invoke-webrequest -uri $IEMode -outfile $nameapk
                    $file = $nameapk
    
                    $searchText = @{
                        "about:blank" = $Url_APk
                    }
                
                    # Membaca isi file ke dalam variabel
                    $content = Get-Content $file
                
                    # Mengganti teks tertentu dalam isi file
                    $searchText.GetEnumerator() | ForEach-Object {
                        $content = $content -replace $_.Key, $_.Value
                    }
                
                    # Menulis isi file yang telah diubah ke dalam file
                    Set-Content $file $content
    
                    $WshShell = New-Object -comObject WScript.Shell
                    $targetPath = $nameapk
                    $iconLocation = "C:\Program Files (x86)\Internet Explorer\iexplore.exe"
                    $iconFile = "IconFile=" + $iconLocation
                    $path = "$loc_users\$Nama_Apk.url"
                    $Shortcut = $WshShell.CreateShortcut($path)
                    $Shortcut.TargetPath = $targetPath
                    $Shortcut.Save()
                
                    Add-Content $path "HotKey=0"
                    Add-Content $path "$iconfile"
                    Add-Content $path "IconIndex=0"
    
                    $fileAttributes = (Get-Item $nameapk).Attributes
                    $fileAttributes += [System.IO.FileAttributes]::Hidden
                    Set-ItemProperty -Path $nameapk -Name Attributes -Value $fileAttributes
                    clear-host
                    write-Host "Shortcut PPK Berhasil dibuat, Silahkan cek di desktop"
                    Show-Footer
                }

                elseif ($Octet_4 -eq 32 -or $Octet_4 -eq 34) {
                    $Apps = "Shortcut for EDP And HR"
                    #Shortcut Monthly Insentive
                    $Url_Apk = "http://192.168.6.92/forms/frmservlet?config=spapps"
                    $Nama_Apk = "Monthly Incentive"

                    $NameApk = [System.String]::Concat("C:\Program Files (x86)\Internet Explorer\", $Nama_Apk, ".vbs")
                    invoke-webrequest -uri $IEMode -outfile $nameapk
                    $file = $nameapk
    
                    $searchText = @{
                        "about:blank" = $Url_APk
                    }
                
                    # Membaca isi file ke dalam variabel
                    $content = Get-Content $file
                
                    # Mengganti teks tertentu dalam isi file
                    $searchText.GetEnumerator() | ForEach-Object {
                        $content = $content -replace $_.Key, $_.Value
                    }
                
                    # Menulis isi file yang telah diubah ke dalam file
                    Set-Content $file $content
                
    
                    $WshShell = New-Object -comObject WScript.Shell
                    $targetPath = $nameapk
                    $iconLocation = "C:\Program Files (x86)\Internet Explorer\iexplore.exe"
                    $iconFile = "IconFile=" + $iconLocation
                    $path = "$loc_users\$Nama_Apk.url"
                    $Shortcut = $WshShell.CreateShortcut($path)
                    $Shortcut.TargetPath = $targetPath
                    $Shortcut.Save()
                
                    Add-Content $path "HotKey=0"
                    Add-Content $path "$iconfile"
                    Add-Content $path "IconIndex=0"
    
                    $fileAttributes = (Get-Item $nameapk).Attributes
                    $fileAttributes += [System.IO.FileAttributes]::Hidden
                    Set-ItemProperty -Path $nameapk -Name Attributes -Value $fileAttributes
                    clear-host
                    write-Host "Shortcut Monthly Insentive Berhasil dibuat, Silahkan cek di desktop"
            
            
                    #Shortcut PPK
                    $Url_Apk = "http://mdsapps:7070/sysHRD/en/neoclassic/login/login"
                    $Nama_Apk = "PPK"

                    $NameApk = [System.String]::Concat("C:\Program Files (x86)\Internet Explorer\", $Nama_Apk, ".vbs")
                    invoke-webrequest -uri $IEMode -outfile $nameapk
                    $file = $nameapk
    
                    $searchText = @{
                        "about:blank" = $Url_APk
                    }
                
                    # Membaca isi file ke dalam variabel
                    $content = Get-Content $file
                
                    # Mengganti teks tertentu dalam isi file
                    $searchText.GetEnumerator() | ForEach-Object {
                        $content = $content -replace $_.Key, $_.Value
                    }
                
                    # Menulis isi file yang telah diubah ke dalam file
                    Set-Content $file $content
    
                    $WshShell = New-Object -comObject WScript.Shell
                    $targetPath = $nameapk
                    $iconLocation = "C:\Program Files (x86)\Internet Explorer\iexplore.exe"
                    $iconFile = "IconFile=" + $iconLocation
                    $path = "$loc_users\$Nama_Apk.url"
                    $Shortcut = $WshShell.CreateShortcut($path)
                    $Shortcut.TargetPath = $targetPath
                    $Shortcut.Save()
                
                    Add-Content $path "HotKey=0"
                    Add-Content $path "$iconfile"
                    Add-Content $path "IconIndex=0"
    
                    $fileAttributes = (Get-Item $nameapk).Attributes
                    $fileAttributes += [System.IO.FileAttributes]::Hidden
                    Set-ItemProperty -Path $nameapk -Name Attributes -Value $fileAttributes
                    clear-host
                    write-Host "Shortcut PPK Berhasil dibuat, Silahkan cek di desktop"
            
                    #Shortcut HRCV
                    $Url_Apk = "http://192.168.6.105/hrcv/Login.aspx"
                    $Nama_Apk = "HRCV"

                    $NameApk = [System.String]::Concat("C:\Program Files (x86)\Internet Explorer\", $Nama_Apk, ".vbs")
                    invoke-webrequest -uri $IEMode -outfile $nameapk
                    $file = $nameapk
    
                    $searchText = @{
                        "about:blank" = $Url_APk
                    }
                
                    # Membaca isi file ke dalam variabel
                    $content = Get-Content $file
                
                    # Mengganti teks tertentu dalam isi file
                    $searchText.GetEnumerator() | ForEach-Object {
                        $content = $content -replace $_.Key, $_.Value
                    }
                
                    # Menulis isi file yang telah diubah ke dalam file
                    Set-Content $file $content
    
                    $WshShell = New-Object -comObject WScript.Shell
                    $targetPath = $nameapk
                    $iconLocation = "C:\Program Files (x86)\Internet Explorer\iexplore.exe"
                    $iconFile = "IconFile=" + $iconLocation
                    $path = "$loc_users\$Nama_Apk.url"
                    $Shortcut = $WshShell.CreateShortcut($path)
                    $Shortcut.TargetPath = $targetPath
                    $Shortcut.Save()
                
                    Add-Content $path "HotKey=0"
                    Add-Content $path "$iconfile"
                    Add-Content $path "IconIndex=0"
    
                    $fileAttributes = (Get-Item $nameapk).Attributes
                    $fileAttributes += [System.IO.FileAttributes]::Hidden
                    Set-ItemProperty -Path $nameapk -Name Attributes -Value $fileAttributes
                    clear-host
                    write-Host "Shortcut HRCV Berhasil dibuat, Silahkan cek di desktop"
                
                    #Shortcut Attendance
                    invoke-webrequest -uri $LinkAtt -outfile "C:\Users\Public\Desktop\Attendance.exe"
                    write-Host "Shortcut Attendance Berhasil dibuat, Silahkan cek di desktop"
                    Show-Footer
                }

                elseif ($Octet_4 -eq 37 -or $Octet_4 -eq 38) {
                    $Apps = "Shortcut for XPDC"
                    #Shortcut RSIM
                    $WshShell = New-Object -comObject WScript.Shell
                    $Shortcut = $WshShell.CreateShortcut("C:\Users\Public\Desktop\RSIM.lnk")
                    $Shortcut.TargetPath = "C:\Program Files (x86)\clientWindows10.4.9\retek\sim\bin\rss.bat"
                    $Shortcut.WorkingDirectory = "C:\Program Files (x86)\clientWindows10.4.9\retek\sim\bin"
                    $Shortcut.Save()
                    write-Host "Shortcut RSIM Berhasil dibuat, Silahkan cek di desktop"
            
                    Show-Footer
                }
                else { Write-Host "Shortcut Website tidak terbentuk karena IP anda $IP4 tidak sesuai dengan penggunaan Web Shortcut, terimakasih" }
            } )
               
        $Uninstall16.Add_Click( 
            {
                clear-host 
                Show-Copyright
                $Apps = " Uninstall Office 2016"
                function Uninstall-Office2016 {
                    Write-Host "Starting Office 2016 uninstallation process..." -ForegroundColor Yellow
                
                    # Create an array to store all target applications
                    $officeApps = @(
                        @{Name = "Microsoft Office Professional Plus 2016*"; Display = "Microsoft Office 2016"},
                        @{Name = "Skype for Business 2016*"; Display = "Skype for Business 2016"},
                        @{Name = "Skype for Business Basic 2016*"; Display = "Skype for Business Basic 2016"},
                        @{Name = "Skype*"; Display = "Skype"}
                    )
                
                    $uninstallCount = 0
                
                    foreach ($app in $officeApps) {
                        Write-Host "Searching for $($app.Display)..." -ForegroundColor Cyan
                        $installation = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like $app.Name }
                        
                        if ($installation) {
                            try {
                                Write-Host "Uninstalling $($app.Display)..." -ForegroundColor Yellow
                                $installation.Uninstall()
                                Write-Host "Successfully uninstalled $($app.Display)" -ForegroundColor Green
                                $uninstallCount++
                            }
                            catch {
                                Write-Host "Error uninstalling $($app.Display): $_" -ForegroundColor Red
                            }
                        }
                        else {
                            Write-Host "No installation found for $($app.Display)" -ForegroundColor Gray
                        }
                    }
                
                    if ($uninstallCount -eq 0) {
                        Write-Host "`nNo Office 2016 components were found to uninstall." -ForegroundColor Yellow
                    }
                    else {
                        Write-Host "`nUninstallation process completed. Successfully removed $uninstallCount component(s)." -ForegroundColor Green
                    }
                
                    # Optional: Clean up registry keys
                    Write-Host "`nCleaning up remaining registry entries..." -ForegroundColor Cyan
                    $regPaths = @(
                        "HKLM:\SOFTWARE\Microsoft\Office\16.0",
                        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
                        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\propertyBag",
                        "HKCU:\Software\Microsoft\Office\16.0"
                    )
                
                    foreach ($path in $regPaths) {
                        if (Test-Path $path) {
                            try {
                                Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
                                Write-Host "Removed registry key: $path" -ForegroundColor Green
                            }
                            catch {
                                Write-Host "Could not remove registry key $path : $_" -ForegroundColor Red
                            }
                        }
                    }
                
                    Write-Host "`nCleanup process completed." -ForegroundColor Green
                    Write-Host "Please restart your computer to complete the uninstallation process." -ForegroundColor Yellow
                }

                Uninstall-Office2016
            } )
                
        $SiteFz.Add_Click( 
            {
                clear-host 
                Show-Copyright
                $Apps = "Setting Filezilla Application"
                $file = "$loc_Roaming\FileZilla\sitemanager.xml"
                invoke-webrequest -uri $LinkFz -outfile $file
                #Rename Site FileZilla
                $Kodetoko = [Microsoft.VisualBasic.Interaction]::InputBox("Masukkan kodetoko:", "Create Site in Filezilla", "445")
                $digit = $Kodetoko % 10

                switch ($digit) {
                    0 { $pass = "bm9s" ; break }
                    1 { $pass = "c2F0dQ==" ; break }
                    2 { $pass = "ZHVh" ; break }
                    3 { $pass = "dGlnYQ==" ; break }
                    4 { $pass = "ZW1wYXQ=" ; break }
                    5 { $pass = "bGltYQ==" ; break }
                    6 { $pass = "ZW5hbQ==" ; break }
                    7 { $pass = "dHVqdWg=" ; break }
                    8 { $pass = "ZGVsYXBhbg==" ; break }
                    9 { $pass = "c2VtYmlsYW4=" ; break }
                }

                $searchText = @{
                    "Dummy" = "mds$Kodetoko"
                    "run"   = $Kodetoko
                    "Pwd"   = $pass
            
                }
            
                # Membaca isi file ke dalam variabel
                $content = Get-Content $file
            
                # Mengganti teks tertentu dalam isi file
                $searchText.GetEnumerator() | ForEach-Object {
                    $content = $content -replace $_.Key, $_.Value
                }
            
                # Menulis isi file yang telah diubah ke dalam file
                Set-Content $file $content
            
                Show-Footer
            } )
                
        $ChangePartition.Add_Click( 
            {
                clear-host 
                Show-Copyright
                $OldDrive = [Microsoft.VisualBasic.Interaction]::InputBox("Masukkan Old Drive:" , "Change Drive Partition", "E")
                $NewDrive = [Microsoft.VisualBasic.Interaction]::InputBox("Masukkan New Drive:", "Change Drive Partition", "D")
                $partition = Get-Partition -DriveLetter $OldDrive
                $partition | Set-Partition -NewDriveLetter $NewDrive
                Write-Host "Process Change Drive Partition dari $OldDrive menjadi $NewDrive Success"
            } )
                
        $L220.Add_Click( 
            {
                clear-host 
                Show-Copyright
                $Apps = 'Printer Epson L220'
                $Driver = 'Printer_L220.exe'
                $Driver2 = 'Scanner_L220.exe'
                if (Get-ChildItem -Path $FileFolder\$Driver -ErrorAction Ignore) {
                    Write-Host "Mohon tunggu, sedang proses menjalankan file" $Apps 
                    start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
                    start-process -FilePath $Driver2 -WorkingDirectory $FileFolder -wait
                    Show-Footer
                }
                else {
                    Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 78.30 MB" 
                    invoke-webrequest -uri $Linkl220printer	-outfile $FileFolder\$Driver
                    invoke-webrequest -uri $Linkl220scanner	-outfile $FileFolder\$Driver2
                    Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
                    start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
                    start-process -FilePath $Driver2 -WorkingDirectory $FileFolder -wait
                    Show-Footer
                }
            } )
            
        $Button1.Add_Click( 
            {
                clear-host 
                Show-Copyright
                $desc = get-netadapter | where { $_.name -like "ethernet" }
                $infdriver = Get-CimInstance win32_PnPSignedDriver | where { $_.friendlyname -like $desc.interfacedescription }
                Write-Host "Mohon tunggu sedang uninstall driver Ethernet"
                pnputil /delete-driver $infdriver.Infname /uninstall /force /reboot 
            
            } )
            
        $Tzone.Add_Click( 
            {
                clear-host 
                Show-Copyright
                $Apps = "Change Timezone"
                $Tz = [Microsoft.VisualBasic.Interaction]::InputBox("Masukkan angka :
                1 = Change Timezone untuk Waktu Indonesia Barat (UTC+07:00)
                2 = Change Timezone untuk Waktu Indonesia Tengah (UTC+08:00)
                3 = Change Timezone untuk Waktu Indonesia Timur (UTC+09:00)" , "Change TimeZone", "1")
                if ($Tz -eq 1) {
                    Clear-Host
                    cmd.exe /c %windir%\system32\tzutil /s "SE Asia Standard Time"
                    Write-Host " TimeZone sudah dirubah menjadi (UTC+07:00)"
                    Show-Footer
                        
                }
                elseif ($Tz -eq 2) {
                    Clear-Host
                    cmd.exe /c %windir%\system32\tzutil /s "Singapore Standard Time"
                    Write-Host " TimeZone sudah dirubah menjadi (UTC+08:00)"
                    Show-Footer
                    
                }
                elseif ($Tz -eq 3) {
                    Clear-Host
                    cmd.exe /c %windir%\system32\tzutil /s "Tokyo Standard Time"
                    Write-Host " TimeZone sudah dirubah menjadi (UTC+09:00)"
                    Show-Footer                    
                }
                else { 
                    Clear-Host
                    Write-Host "Please Input Pilihan 1 , 2 atau 3"
                }
                        
                
            } )
            
        $PowerBI.Add_Click( 
            {
                clear-host 
                Show-Copyright
                $Apps = 'Power BI'
                $Driver = 'PowerBI.exe'
                if (Get-ChildItem -Path $FileFolder\$Driver -ErrorAction Ignore) {
                    Write-Host "Mohon tunggu, sedang proses menjalankan file" $Apps 
                    start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
                    Show-Footer
                }
                else {
                    Write-Host "Mohon tunggu, file sedang di download " $Apps "size file : 400.30 MB" 
                    invoke-webrequest -uri $LinkPowerBI	-outfile $FileFolder\$Driver
                    Write-Host "Download" $Apps "Success. Mohon tunggu, sedang proses menjalankan file" $Apps 
                    start-process -FilePath $Driver -WorkingDirectory $FileFolder -wait
                    Show-Footer
                } 
            } )
            
            
        $VNC.Add_Click( 
            {
                clear-host 
                Show-Copyright
                $ErrorActionPreference = 'SilentlyContinue'
                Remove-Item -Path "$env:TEMP\*" -Recurse -Force
                $folderPath_VNC = "D:\VNC"
                if (-Not (Test-Path -Path $folderPath_VNC -PathType Container)) {
                    New-Item D:\VNC -itemType Directory -ErrorAction Ignore -Force
                    $folder = Get-Item -Path $folderPath_VNC
                    $folder.Attributes = $folder.Attributes -bor [System.IO.FileAttributes]::Hidden
                    invoke-webrequest -uri "https://raw.githubusercontent.com/Rangga4869/script/refs/heads/main/tightvnc.msi" -outfile "$folderPath_VNC\VNC.zip"
                    expand-archive -path "$folderPath_VNC\VNC.zip" "$folderPath_VNC\"
                    remove-item "$folderPath_VNC"\VNC.zip -force
                }
                $IPVNC = [Microsoft.VisualBasic.Interaction]::InputBox("Masukkan IP Address yang akan diremote (172.xx.xx.xx):", "Remote IP Address")
                $Apps = "Running VNC on $IPVNC"
                cmd /c D:\VNC\tvnviewer.exe -host="$IPVNC" -password=matahari
                Show-Footer
            } )
            
        $IE_mode.Add_Click( 
            {
                clear-host 
                Show-Header
                $Apps = 'Create IE Mode'

                $IEMode = [Microsoft.VisualBasic.Interaction]::InputBox("Silahkan Input Nomor Shortcut yang diinginkan ( Contoh : untuk shortcut alphapos ketik 1) :
    
     1. Shortcut Alphapos Web   
     2. Shortcut HRCV
     3. Shortcut Monthly Incentive ( SSI ) 
     4. Shortcut Report RSIM
     5. Shortcut PPK
    ", "Create Shortcut", "1")
        switch ($IEMode) {
            '1' { $Url_Apk = "http://$Octet.11:8500/alphaposwebv2/main.htm" ; $Nama_Apk = "Alphapos" ; Break }
            '2' { $Url_Apk = "http://192.168.6.105/hrcv/Login.aspx" ; $Nama_Apk = "HRCV" ; Break }
            '3' { $Url_Apk = "http://192.168.6.92/forms/frmservlet?config=spapps" ; $Nama_Apk = "Monthly Incentive" ; Break }
            '4' { $Url_Apk = "http://192.168.7.15:8882/report_prdds/login.jsp" ; $Nama_Apk = "Report RSIM" ; Break }
            '5' { $Url_Apk = "http://mdsapps:7070/sysHRD/en/neoclassic/login/login" ; $Nama_Apk = "PPK" ; Break }
            default { $PC = "1" }    
        }
                $NameApk = [System.String]::Concat("C:\Program Files (x86)\Internet Explorer\", $Nama_Apk, ".vbs")
                invoke-webrequest -uri $IEMode -outfile $nameapk
                $file = $nameapk

                $searchText = @{
                    "about:blank" = $Url_APk
                }
            
                # Membaca isi file ke dalam variabel
                $content = Get-Content $file
            
                # Mengganti teks tertentu dalam isi file
                $searchText.GetEnumerator() | ForEach-Object {
                    $content = $content -replace $_.Key, $_.Value
                }
            
                # Menulis isi file yang telah diubah ke dalam file
                Set-Content $file $content
            

                $WshShell = New-Object -comObject WScript.Shell
                $targetPath = $nameapk
                $iconLocation = "C:\Program Files (x86)\Internet Explorer\iexplore.exe"
                $iconFile = "IconFile=" + $iconLocation
                $path = "$loc_users\$Nama_Apk.url"
                $Shortcut = $WshShell.CreateShortcut($path)
                $Shortcut.TargetPath = $targetPath
                $Shortcut.Save()
            
                Add-Content $path "HotKey=0"
                Add-Content $path "$iconfile"
                Add-Content $path "IconIndex=0"

                $fileAttributes = (Get-Item $nameapk).Attributes
                $fileAttributes += [System.IO.FileAttributes]::Hidden
                Set-ItemProperty -Path $nameapk -Name Attributes -Value $fileAttributes
                clear-host
                write-Host "Shortcut $apk Berhasil dibuat, Silahkan cek di desktop"
                Show-Footer
            } )
            
        $Join_Domain.Add_Click( 
            {
                clear-host 
                Show-Header
                $confirm = [Microsoft.VisualBasic.Interaction]::InputBox("Apakah sudah konfirmasi akan join domain ? :
1. Sudah konfirmasi  
2. Belum konfirmasi", "Confirm Join Domain", "1")
If ($confirm -eq 1) {
    
                $hostname = $env:COMPUTERNAME
                if ($hostname -match "^desktop\d{6}$") {
                    write-Host "Hostname di komputer anda, tidak sesuai dengan standart IT, silahkan restart terlebih dahulu lalu jalankan kembali"
                    Restart-Computer -Force
                    } else {
                    $domain = "matahari.id"
                    $password = "Vermouth09" | ConvertTo-SecureString -asPlainText -Force
                    $username = "$domain\sam104358"
                    $credential = New-Object System.Management.Automation.PSCredential($username,$password)
                    Add-Computer -DomainName $domain -Credential $credential
                    Restart-Computer
                    Show-Footer    
                    }
} else {
    Write-host "Silahkan confirmasi terlebih dahulu lalu jalankan kembali, terimakasih"
}
            } )
            
        $Change_Hostname.Add_Click( 
            {
                clear-host 
                Show-Header
                #Change Hostname
                if ($Octet_1 -eq "172") {
                    if ((gwmi win32_computersystem).partofdomain -eq $false) {

                        switch ($Octet_4) {
                            29 { $div = '01'; $Cat = 'L' ; break }
                            30 { $div = '01'; $Cat = 'LS' ; break }
                            31 { $div = 'EDP'; $Cat = 'DS' ; break }
                            32 { $div = 'HRS'; $Cat = 'DS' ; break }
                            33 { $div = 'STAFF'; $Cat = 'DS' ; break }
                            34 { $div = 'ADMHR'; $Cat = 'DS' ; break }
                            35 { $div = 'VM'; $Cat = 'DS' ; break }
                            36 { $div = 'ASM'; $Cat = 'DS' ; break }
                            37 { $div = 'XPDC'; $Cat = 'DS' ; break }
                            38 { $div = 'XPDC2'; $Cat = 'DS' ; break }
                            39 { $div = 'OSS'; $Cat = 'DS' ; break }
                            default { Write-Host "IP Address tidak sesuai dengan standar IT" }
                        }        
                        $NameComputer = [System.String]::Concat($Code, $Cat, $Str, $Div)
                        $NewnameComputer = [Microsoft.VisualBasic.Interaction]::InputBox("Silahkan input $NameComputer lalu dibelakangnya diberi huruf lain" , "Input other computername")
                        Rename-Computer -NewName $NewnameComputer -Force
                        Write-Host "Process ganti Computername menjadi" $NewnameComputer "Suksess"
                        Write-Host "Silahkan restart komputer agar komputer dapat berubah"
                    } 
                }      
            } )
            
        $Button8.Add_Click( 
            {
                clear-host 
                Show-Header


      
            } )
            
        $Button9.Add_Click( 
            {
                clear-host 
                Show-Header
      
            } )
            
        $Button11.Add_Click( 
            {
                clear-host 
                Show-Header


          
            } )
            
        $WinUpdate.Add_Click( 
            {
                clear-host 
                Show-Header
      
            } )
            
        $Button12.Add_Click( 
            {
                clear-host 
                Show-Header


         
            } )
            
        $Button13.Add_Click( 
            {
                clear-host 
                Show-Header


         
            } )
            
        $Button14.Add_Click( 
            {
                clear-host 
                Show-Header
     
            } )
            
        $Button15.Add_Click( 
            {
                clear-host 
                Show-Header



            } )
            
        $Button16.Add_Click( 
            {
                clear-host 
                Show-Header
        
            } )
            
        $Tzone1.Add_Click( 
            {
                clear-host 
                Show-Header
        
            } )
            
        $Button17.Add_Click( 
            {
                clear-host 
                Show-Header
          
            } )
            
        $Button18.Add_Click( 
            {
                clear-host 
                Show-Header
          
            } )
            
        $Button19.Add_Click( 
            {
                clear-host 
                Show-Header

      
            } )
                
        $Tzone0.Add_Click( 
            {
                clear-host 
                Show-Header
         
            } )
            
        $Tzone2.Add_Click( 
            {
                clear-host 
                Show-Header

        
            } )
            
        $Tzone3.Add_Click( 
            {
                clear-host 
                Show-Header

           
            } )
            
        $Tzone5.Add_Click( 
            {
                clear-host 
                Show-Header
          
            } )
                
        $Tzone4.Add_Click( 
            {
                clear-host 
                Show-Header

         
            } )
            
        $Tzone6.Add_Click(
            {
                clear-host 
                Show-Header


        
            } )	
            
        $Tzone7.Add_Click(
            {
                clear-host 
                Show-Header
         
            } )	
            
        $Tzone8.Add_Click(
            {
                clear-host 
                Show-Header

        
            } )
            
        $Button10.Add_Click( 
            {
                clear-host 
                Show-Header

              
            } )
                
            
            
        [void]$Form.ShowDialog()
            
            
    }
)
#exit program other

[void]$Form.ShowDialog()
