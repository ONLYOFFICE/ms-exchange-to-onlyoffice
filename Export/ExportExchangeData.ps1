# ---- Begin script Export-PSTs.ps1 ----

param (
    [string]$dir = "C:\Temp"
 )

echo "Working directory: $dir"

New-Item -ItemType Directory -Force -Path $dir

$dirName = Split-Path $dir -Leaf

New-SMBShare -Name "$dirName" -Path "$dir" -ContinuouslyAvailable 0 -FullAccess "everyone"

$usersListPath = "$dir\users.csv"

if (Test-Path $usersListPath) {
  Remove-Item $usersListPath
}

Get-Mailbox | Select DisplayName, PrimarySmtpAddress, Alias | Export-Csv $usersListPath

$pstDirName = "PST"

$pstpath = "$dir\$pstDirName"

Remove-Item -Recurse -ErrorAction Ignore -Path $pstpath

New-Item -ItemType Directory -Force -Path $pstpath

$hostName = $env:computername

Get-MailboxExportRequest | Remove-MailboxExportRequest -Confirm:$false -Force

$AllMailboxes = Get-Mailbox

$i = 0

foreach ($Mailbox in $AllMailboxes) { 
    try {
        $i += 1

        $path = "\\$hostName\$dirName\$pstDirName\$($Mailbox.Alias).pst"

        echo "$i. Exporting $path"

        New-MailboxExportRequest -Mailbox $Mailbox -FilePath "$path"

        while(!(Get-MailboxExportRequest -Mailbox $Mailbox -Status Completed) -and !(Get-MailboxExportRequest -Mailbox $Mailbox -Status Failed))
        {
            #Sleep for 1 second
            Start-Sleep -s 1
        }
        
        Get-MailboxExportRequest -Mailbox $Mailbox | Remove-MailboxExportRequest -Confirm:$false -Force

    }
    catch
    {
        
    }
}

Remove-SmbShare -Name "$dirName" -Force

# ----  ./Export-PSTs.ps1  ----