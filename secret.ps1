<#
.SYNOPSIS
This is a simple powershell script for creating and sending secret santa pairings over a sendgrid email service

.DESCRIPTION
This script will take a .csv file containing the columns 'name' and 'email', an output logfile, and an email address to send from. Then it will randomly pair all attendees. The algorithm used guarantees that no one will receive their own name and since all pairings happen at once guarntees random pairings.

.EXAMPLE
./SecretSanta.ps1 -Signups signups.csv -LogFile out.log -SendAddr SecretSanta@gmail.com
#>
param($Signups, $LogFile, $SendAddr)

if (!(Test-Path $signups))
{
    write-host "You must supply a signups of signups."
    return
}

function makePairings ([System.Collections.ArrayList]$attendees) 
{
    $first = $attendees[0]
    $last = $attendees[-1]

    for($i = 0; $i -lt $attendees.count - 1; $i++)
    {
        $attendees[$i].match = $attendees[$i + 1].name
    }
    $last.match = $first.name
}

function sendOutEmails([System.Collections.ArrayList]$matches, $credentials) 
{
    #$username = cat .\user.txt
    #$password = cat .\password.txt | ConvertTo-SecureString
    #$cred = new-object -typename System.Management.Automation.PSCredential `
    #        -ArgumentList $username, $password
    foreach($entry in $matches) 
    {
        Send-MailMessage -From $SendAddr `
        -To $entry.email `
        -Subject "Secret santa pairings!" `
        -body "Hi $($entry.name). You will be gifting for $($entry.match)! Happy Holidays!" `
        -SmtpServer "smtp.sendgrid.net" `
        -port 587 `
        -credential $credentials `
        -usessl
    }
}

function GetRandomRows()
{
    $rows = new-object System.Collections.ArrayList
    Import-CSV $signups | foreach-object {
        $hash = (@{Name=$_.name;email=$_.email;match=""})
        $row = new-object psobject -Property $hash
        $rows.add($row) > $null
    }

    $rows = $rows | Sort-Object {get-random}
    return $rows
}

function CleanupWorkspace()
{
    clear-content output.txt
}

function Main()
{
    $rows = GetRandomRows
    makePairings $rows
    $rows | Export-Csv $LogFile
    $cred = Get-Credential
    sendOutEmails $rows $cred
}

Main




