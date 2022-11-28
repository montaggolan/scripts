# Based on https://social.technet.microsoft.com/Forums/windowsserver/en-US/010e3f4f-755b-4cb4-b574-69e5db6e3278/powershell-outlook-2010-converting-pst-to-xml
# Version 1.1
# Date 05.04.2019
# Function to extract specific attributes from mails based on keywords. Requires Outlook.


param (
    [string] $keywords=$null,
    [string] $outputFile=$null, #Note that if full path is not given, outpufile will be placed in home directory
    [string] $inputFile=$null,
    [string] $inputFolder=$null,
    [boolean] $enableBodySearch=$false
    #[boolean] $storeBody=$false
)

# Globals
$newci = [System.Globalization.CultureInfo]"en-US"
[system.threading.Thread]::CurrentThread.CurrentCulture = $newci
[System.Threading.Thread]::CurrentThread.CurrentUICulture = $newci
$keywordlist='bird','is','the','word'
$curwd=$(Convert-Path ".")
$interestingFolders = "Deleted Items", "Drafts", "Inbox", "Sent Items", "Recoverable Items", "Conversation History"
#"Recoverable Items\Deletions", "Recoverable Items\Purges"
#$exclusionTo = ""

if ($keywords){
   foreach($kw in $keywords){
   $keywordlist+=$kw
   }
}

if ($inputFile){
    $inputFile=(Resolve-Path $inputFile | Select -ExpandProperty path)
    if([System.IO.File]::Exists($inputFile))
    {
        $pstFile=(Split-Path $inputFile -Leaf)
        $pstPath=(Split-Path $inputFile)
    }
    else{
        Write-Host "Can't Find the given file. Exiting. To scan current directory with all the pst files, run without giving input file and folder."
        exit
    }
}
Elseif($inputFolder){
    if(Test-Path $inputFolder){
        $pstPath=$inputFolder
    }
    else{
        Write-Host "The given folder path can't be found. Exiting. To scan current directory with all the pst files, run without giving input file and folder."
        exit
    }
}
else{
    $pstPath=$curwd
}

# To get all the mail items in a PST by recursing through the folder structure.
function Get-AllMAPIFolders( [Object]$RootFolder )
{
  $folderList = New-Object 'System.Collections.Generic.List[Object]'
  $folderList.Add($RootFolder)
  foreach( $subFolder in $RootFolder.Folders ) {
     if (-Not ($subFolder.FolderPath -Match "Calendar Logging")){
     foreach( $intFolder in $interestingFolders ){
      if ( $subFolder.FolderPath -Match $intFolder ) {
         if ( $subFolder.Folders.Count -gt 0 ) {
            $folderList.AddRange((Get-AllMAPIFolders $subFolder))
         } else {
        $folderList.Add($subFolder)
        }
      }
    }
    }
  }
  return $folderList
}

function Extract-Fields([Object] $allFolders)
{
    foreach( $folder in $allFolders )
    {
        foreach ( $item in $folder.Items )
        {
          #if ( $exclusionTo -contains $item.To) {continue}
          foreach ($kw in $keywordlist){
              if($item.SenderName -match $kw -or $item.Subject -match $kw -or $item.To -match $kw -or $item.CC -match $kw -or ($enableBodySearch -and $item.Body -match $kw))
              {
                  if ( $outputFile)
                  {
                      $fPath = "Folder: " + $folder.FolderPath
                      Add-Content -Path $outputFile -Value $fPath
                      $sender = "Sender: " + $item.SenderName
                      Add-Content -Path $outputFile -Value $sender
                      $subject = "Subject: " + $item.Subject
                      Add-Content -Path $outputFile -Value $subject
                      $sentTo = "Sent To: " + $item.To
                      Add-Content -Path $outputFile -Value $sentTo
                      $cc = "CC: " + $item.CC
                      Add-Content -Path $outputFile -Value $cc
                      $bcc = "BCC: " + $item.BCC
                      Add-Content -Path $outputFile -Value $bcc
                      $recuPar = "Received By: " + $item.ReceivedByName
                      Add-Content -Path $outputFile -Value $recuPar
                      $sendDate = "Send Date: " + $item.SentOn
                      Add-Content -Path $outputFile -Value $sendDate
                      if ($item.Attachments.Count -gt 0) {
                          Add-Content -Path $outputFile -Value "Attachments:"
                          foreach($attachment in $item.Attachments)
                          {
                              $attDispName = "    DisplayName: " + $attachment.DisplayName
                              Add-Content -Path $outputFile -Value $attDispName
                              $attFileName = "    FileName: " + $attachment.FileName
                              Add-Content -Path $outputFile -Value $attFileName
                          }
                      }
                      Add-Content -Path $outputFile -Value "----------------"
                  }
                  else
                  {
                      Write-Host "Folder: ", $folder.FolderPath
                      Write-Host "Sender: ", $item.SenderName
                      Write-Host "Subject: ", $item.Subject
                      Write-Host "Sent To: ", $item.To
                      Write-Host "CC: ", $item.CC
                      Write-Host "BCC: ", $item.BCC
                      Write-Host "Received By: ", $item.ReceivedByName
                      Write-Host "Send Date: ", $item.SentOn
                      if ($item.Attachments.Count -gt 0) {
                          Write-Host "Attachments:"
                          foreach($attachment in $item.Attachments)
                          {
                              Write-Host "    DisplayName:", $attachment.DisplayName
                              Write-Host "    FileName:", $attachment.FileName
                          }
                      }
                      Write-Host "----------------"
                  }
              }
          }
        }
    }
}

function main ()
{
    # Outlook needs to be running already to attach to it
    $oProc = ( Get-Process | where { $_.Name -eq "OUTLOOK" } )
    if ( $oProc -eq $null ) { Start-Process outlook -WindowStyle "Minimized"; Start-Sleep -Seconds 5 }
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    if($pstFile){
        $namespace.AddStoreEx($pstPath+'\'+$pstFile, "olStoreDefault")
        $pstStore = ( $nameSpace.Stores | where { $_.FilePath -eq $pstPath+'\'+$pstFile } )
        $pstRootFolder = $pstStore.GetRootFolder()
        $allFolders = Get-AllMAPIFolders $pstRootFolder
        Extract-Fields $allFolders
    }
    else{
        foreach( $file in (Get-ChildItem $pstPath -Filter *.pst)){
            $tempPath=$file.FullName
            $namespace.AddStoreEx($tempPath, "olStoreDefault")
            $pstStore = ( $nameSpace.Stores | where { $_.FilePath -eq $tempPath } )
            $pstRootFolder = $pstStore.GetRootFolder()
            $allFolders = Get-AllMAPIFolders $pstRootFolder
            Extract-Fields $allFolders
        }
    }
}

main
