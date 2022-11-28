$kdbxPath=""
$keepassKeyFilePath=""
$pwWordlist=""

﻿foreach($line in Get-Content $pwWordlist) {
    Write-Output $line
    & "C:\Program Files (x86)\KeePass Password Safe 2\KeePass.exe" $kdbxPath -pw:"$line" -keyfile:$keepassKeyFilePath
    Write-Output $LASTEXITCODE
    Read-Host -Prompt 'Input the user name'
}
