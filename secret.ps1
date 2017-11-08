param($File)

if ($File -eq "" -or $File -eq $null)
{
    write-host "You must supply a file you scrub"
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

function sendOutEmails([System.Collections.ArrayList]$matches) 
{
    $username = "azure_2e52969a583f648d72302eab0139df2b@azure.com"
    $password = cat .\password.txt | ConvertTo-SecureString
    $cred = new-object -typename System.Management.Automation.PSCredential `
            -ArgumentList $username, $password
    foreach($entry in $matches) 
    {
        Send-MailMessage -From "SecretSanta@compeclub.com" `
        -To $entry.email `
        -Subject "Secret santa pairings!" `
        -body "Hi $($entry.name). You will be buying for $($entry.match)!" `
        -SmtpServer "smtp.sendgrid.net" `
        -port 587 `
        -credential $cred `
        -usessl
    }
}

$rows = new-object System.Collections.ArrayList
Import-CSV $File | foreach-object {
    $hash = (@{Name=$_.name;email=$_.email;match=""})
    $row = new-object psobject -Property $hash
    $rows.add($row) > $null
}

$rows = $rows | Sort-Object {get-random}
makePairings $rows
clear-content output.txt
$rows | Export-Csv output.txt
#sendOutEmails $rows


