#region notes


#endregion

#region TODO
<#

    1) Error handling
    2) Start where the script left off

#>
#endregion


Param (

    
    [Parameter(Mandatory = $true,Position = 1 )]
    [string]$FilePath = ( Join-Path "$($env:USERPROFILE)\Desktop" -ChildPath ( ( Get-Date -Format yyyy-MM-dd ) + '_' + 'V11-6-16' + '_' + 'MailStats' + '.csv' ) ),
    
    [Parameter(Mandatory = $false)]
    [string]$StartDate = '2001-01-01',
    
    [Parameter(Mandatory = $false)]
    [datetime]$EndDate = ( Get-Date ).AddDays(1)


)

BEGIN {

#region Variables

    $restUri = 'https://outlook.office365.com/api/beta/users'
    $UserCredential = Get-Credential

    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    Import-PSSession $Session -AllowClobber
    $users = Get-Mailbox -Filter { RecipientTypeDetails -ne 'DiscoveryMailbox'  } -SortBy DisplayName

    $Start = Get-Date $StartDate -Format yyyy-MM-dd
    $End = Get-Date $EndDate -Format yyyy-MM-dd

    $select = 'SentDateTime,ReceivedDateTime,Sender,ToRecipients,CCRecipients,BCCRecipients,Subject'
    $filter = "ReceivedDateTimege $Start and ReceivedDateTime le $End"
        
#endregion

}

PROCESS {

#region Processing

    $results = @(
    
    for ( $i = 0 ; $i -lt $users.Count ; $i++ ) {


        $ProgressProperties = [psobject]@{
                                        
            Activity = 'Getting mail information'
            CurrentOperation = "$Percent% complete"
            PercentComplete = $Percent = [math]::Round( ( $i / $users.Count * 100 ) , 2 )
            Status = $users[$i].UserPrincipalName
                                        
            }

        Write-Progress @ProgressProperties
    
        $top = 25
        $skip = 0

        do {          
             $batch = Invoke-RestMethod -Method Get -Uri "$restUri/$($users[$i].UserPrincipalName)/messages?&`$filter=$filter&`$select=$select&`$top=$top&`$skip=$skip" -Credential $UserCredential
             $results += $batch.value
             $skip += 25
       
           }

          until ( $batch.'@odata.nextLink' -eq $null )

    }

#endregion


}

END {

$results | select SentDateTime, ReceivedDateTime, @{ n = 'Sender' ; e = { $_.Sender.EmailAddress.Address } }, @{ n = 'ToRecipients' ; e = { $_.ToRecipients.EmailAddress | %{ $_.Address } } }, @{ n = 'CCRecipients' ; e = { $_.CCRecipients.EmailAddress | %{ $_.Address } } },  @{ n = 'BccRecipients' ; e = { $_.BccRecipients.EmailAddress | %{ $_.Address } } }, Subject |Export-csv -NoTypeInformation $FilePath


#region Cleanup

Get-PSSession | Remove-PSSession
Remove-Variable Session

#endregion


}