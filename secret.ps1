<#
.SYNOPSIS
This is a simple powershell script for creating and sending secret santa pairings over a sendgrid email service

.DESCRIPTION
This script will take a .csv file containing the columns 'name' and 'email', an output logfile, and an email address to send from. Then it will randomly pair all attendees. The algorithm used guarantees that no one will receive their own name and since all pairings happen at once guarntees random pairings.

.EXAMPLE
./SecretSanta.ps1 -Signups signups.csv -LogFile out.log -SendAddr SecretSanta@gmail.com
#>
Param(
    [Parameter(Mandatory=$True, Position=0)]
    $Signups, 
    [Parameter(Mandatory=$True, Position=1)]
    $SendAddr, 
    [Parameter(Mandatory=$True, Position=2)]
    $LogFile, 
    [Switch]$RunLocally
)
try {
    if (!(Test-Path $Signups))
    {
        Write-Error "You must supply a valid path to a csv containing signups."
        exit
    }
} catch {
    Write-Error "You must supply a valid path to a csv containing signups."
    exit
}

function MakePairings ([System.Collections.ArrayList]$attendees) 
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
    $rows = New-Object System.Collections.ArrayList
    Import-CSV $signups | ForEach-Object {
        $hash = (@{Name=$_.name;email=$_.email;match=""})
        $row = New-Object PSObject -Property $hash
        $rows.add($row) > $null
    }

    $rows = $rows | Sort-Object { Get-Random }
    return $rows
}

function Main()
{
    $rows = GetRandomRows
    MakePairings $rows
    $rows | Export-Csv $LogFile
    Add-Content $LogFile $(Get-Date)
    $cred = Get-Credential
    if(!$RunLocally)
    {
        sendOutEmails $rows $cred
    }
}

Main




