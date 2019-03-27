# Inspired by, some parts taken from https://winscp.net/eng/docs/library_powershell#example
# SFTP-Backup script using the WinSCP .NET module by Thomas Erbesdobler <t.erbesdobler@team103.com>

# parameters
[CmdletBinding()]
Param(
    [switch]$test_emails
)

# Configuration
$hostname = "host"
$username = "user"
$password = "password"
$hostkey = "ssh-rsa 2048 fingerprint"

[string] $remote_source = "/backup"
[string[]] $backup_directories = "test1", "test2"
[string] $local_target = "D:\backup"

[float] $min_age_copy_hours = 6
[int] $min_count_keep = 2
[int] $max_count_transfer_workday = 1
[int] $max_count_transfer_weekend = 10

# EMail notification
$email_enabled=$True
$custom_subject_prefix='subject'
$mail_from='sender address'

$mail_to = @()
$rcpt1 = 'receiver', 'smtp server', 25
$mail_to = $rcpt1,$null

# Variable initialization
$email_message = ""

# Functions
function log($msg, $err=$False) {
    if ($err) {
        write-host -ForegroundColor Red $msg
    } else {
        write-host $msg
    }

    if ($email_enabled) {
        $script:email_message += "`n"

        if ($err) {
            $script:email_message += "ERROR: $($msg)`n"
        } else {
            $script:email_message += $msg
        }
    }
}

function log_exit($exit_code) {
    if ($email_enabled)
    {        
        $script:email_message += "`n`nEnding at $(Get-Date) "

        if ($exit_code -ne 0) {
            $pr=[System.Net.Mail.MailPriority]::High
            $subject = "$($custom_subject_prefix)failed"

            $script:email_message += "with failures."
        } else {
            $pr=[System.Net.Mail.MailPriority]::Normal
            $subject = "$($custom_subject_prefix)succeeded"

            $script:email_message += "successfully."
        }

        foreach ($receiver in $mail_to) {
            if($receiver)
            {
                $params = @{
                    "From" = $mail_from
                    "To" = $receiver[0]
                    "Priority" = $pr
                    "Subject" = $subject
                    "Body" = $email_message
                    "Encoding" = [System.Text.Encoding]::UTF8
                    "SmtpServer" = $receiver[1]
                    "Port" = $receiver[2]
                    "DeliveryNotificationOption" = [System.Net.Mail.DeliveryNotificationOptions]::None
                }

                send-mailmessage @params
            }
        }
    }

    exit $exit_code
}

function ensure_dir([string]$dir) {
    if (!(Test-Path -Path $dir)) {
        New-Item -ItemType directory -Path $dir
    }
}

function is_weekend([System.Datetime]$date) {
    if ($date.DayOfWeek -eq [System.DayOfWeek]::Friday -and
        $date.Hour -gt 18) {
        return $True
    }

    if ($date.DayOfWeek -eq [System.DayOfWeek]::Saturday) {
        return $True
    }

    if ($date.DayOfWeek -eq [System.DayOfWeek]::Sunday -and
        $date.Hour -lt 6) {
        return $True
    }

    return $False
}

function backup($session, [string]$dir)
{
    $ret_code = 0

    [System.DateTime] $now = Get-Date

    $remote = "$remote_source/$dir"
    $local = (Join-Path $local_target $dir)

    # Retrieve files in remote directory
    $remote_dir = $session.ListDirectory($remote)

    $remote_files = @()

    foreach ($fileinfo in $remote_dir.Files)
    {
        if ($fileinfo.Name -ne '.' -and $fileinfo.Name -ne '..') {
            $remote_files += $fileinfo
        }
    }

    # Find all files that could be deleted and all that could be copied
    $remote_files = ($remote_files | Sort-Object {-$_.LastWriteTime},{$_.Name})

    $to_delete = @()
    $to_copy = @()
    $element_counter = 0

    foreach($file in $remote_files)
    {
        $element_counter += 1

        if (($now - $file.LastWriteTime).TotalHours -ge $min_age_copy_hours) {
            $to_copy += $file
        }

        if ($element_counter -gt $min_count_keep) {
            $to_delete += $file
        }
    }

    # Create local directory if it does not exist
    $ld = "$local\"
    ensure_dir $ld

    # Remove files from copy list that were already copied
    $local_files = Get-ChildItem -File -Path $local

    $tmp = $to_copy
    $to_copy = @()

    # Only add as many files as we are allowed to copy. Prefer newer files
    # (this is ensured by the arrays' sorting).
    $counter = 0

    if (is_weekend $now) {
        $max_count = $max_count_transfer_weekend
    } else {
        $max_count = $max_count_transfer_workday
    }

    foreach ($rf in $tmp)
    {
        $found = $False

        foreach ($lf in $local_files) {
            if ($lf.Name -eq $rf.Name -and $lf.LastWriteTime -ge $rf.LastWriteTime) {
                $found = $True
                break
            }
        }

        if (!$found) {
            if ($counter -lt $max_count) {
                $to_copy += $rf
                $counter += 1
            } else {
                break
            }
        }
    }

    # Copy files
    $transferOptions = New-Object WinSCP.TransferOptions
    $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
    $transferOptions.ResumeSupport.State = [WinSCP.TransferResumeSupportState]::On

    foreach ($file in $to_copy) {
        $rf = "$remote/$($file.Name)"

        log "  Get ""$rf"" -> ""$ld"""

        $transferResult = $session.GetFiles(
            $rf,
            $ld,
            $False,
            $transferOptions)

        if (-not $transferResult.isSuccess) {
            $ret_code = 1

            foreach($err in $transferResult.Failures) {
                log $err $True
            }
        }
    }

    # Remove files from delete list that must still be retained
    # Those are all files that we don't have locally yet.
    $local_files = Get-ChildItem -File -Path $local

    $tmp = $to_delete
    $to_delete = @()

    foreach ($df in $tmp)
    {
        $found = $False

        foreach ($lf in $local_files) {
            if ($lf.Name -eq $df -and $lf.LastWriteTime -ge $df.LastWriteTime) {
                $found = $True
                break
            }
        }

        if ($found) {
            $to_delete += $df
        }
    }

    # Delete files
    foreach ($file in $to_delete) {
        $rf = "$remote/$($file.Name)"

        log "    rm ""$rf"""

        $removalResult = $session.RemoveFiles($rf)

        if (-not $removalResult.isSuccess) {
            $ret_code = 1

            foreach($err in $removalResult.Failures) {
                log $err $True
            }
        } 
    }

    return $ret_code
}

# Program flow starts here
if ($test_emails) {
    $email_enabled = $True
    $email_message += "Test EMail from backup script at $(Get-Date)."

    log_exit 0
}

if ($email_enabled) {
    $email_message += "I began to backup files at $(Get-Date):`n"
    $email_message += "From: $($hostname):$remote_source`n"
    $email_message += "To:   $local_target`n"
}

try
{
    # Load WinSCP .NET assembly
    Add-Type -Path (Join-Path $PSScriptRoot "WinSCPnet.dll")

    # Setup session options
    $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
        Protocol = [WinSCP.Protocol]::Sftp
        HostName = $hostname
        UserName = $username
        Password = $password
        SshHostKeyFingerprint = $hostkey
    }

    $session = New-Object WinSCP.Session
    $exit_code = 0

    try
    {
        # Connect
        $session.Open($sessionOptions)

        foreach ($d in $backup_directories)
        {
            if ((backup $session $d) -ne 0) {
                $exit_code = 1
            }
        }
    }
    finally
    {
        # Disconnect, clean up
        $session.Dispose()
    }

    log_exit $exit_code
}
catch
{
    log "Error: $($_.Exception.Message)"
    log_exit 1
}