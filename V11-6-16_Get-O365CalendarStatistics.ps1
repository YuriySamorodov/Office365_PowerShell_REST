#region NOTES


#endregion



#region TODO

<#

    1) Error handling
    2) Start where the script left off

#>
#endregion


Param (
   
    [Parameter(Mandatory = $true,Position = 1 )]
    [string]$FilePath = ( Join-Path "$($env:USERPROFILE)\Desktop" -ChildPath ( ( Get-Date -Format "yyyy-MM-dd_HH-mm-ss" ) + '_' + 'V11-6-16' + '_' + 'CalendarStats' + '.csv' ) ),
    
    [Parameter(Mandatory = $false)]
    [string]$StartDate = '2001-01-01',
    
    [Parameter(Mandatory = $false)]
    [datetime]$EndDate = ( Get-Date ).AddDays(1),

    [Parameter(Mandatory = $false,Position = 1 )]
    [string]$Log = ( Join-Path "$($env:USERPROFILE)\Desktop" -ChildPath ( ( Get-Date -Format yyyy-MM-dd ) + '_' + 'V11-6-16' + '_' + 'Log' + '_' + 'CalendarStats' + '.log' ) )
    


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

        $select = 'Organizer,Attendees,ResponseStatus,Start,End,Subject'
        $filter = "CreatedDateTime ge $Start and CreatedDateTime le $End"      

#endregion

}

PROCESS {

#region Processing

     get-date -Format "yyyy-MM-dd HH:mm:ss" | Out-File -FilePath  $Log -Append

    $results = @()
    
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

             $batch = Invoke-RestMethod -Method Get -Uri "$restUri/$($users[$i].UserPrincipalName)/calendarview?StartDateTime=$($Start)T01:00:00&EndDateTime=$($End)T23:00:00&`$skip=$skip" -Credential $UserCredential
             $results += $batch.value | select *, @{ Name = 'User' ; Expression = { $user  }  } 
             $skip += 25
       
           }

          until ( $batch.'@odata.nextLink' -eq $null )

        $users[$i].UserPrincipalName | Out-File $Log -Append

    }

#endregion


}

END {

$results | select id, @{ Name = 'Meeting Title' ; Expression = { $_.Subject } }, user, @{ n = 'Organizer' ; e = { $_.Organizer.EmailAddress.Address } }, @{ n = 'Attendee:Type:ResponseTime:ResponseStatus' ; e = { $_.Attendees | %{ $_.EmailAddress.Address + ':' + $_.Type + ':' + $_.Status.Time + ':' + $_.Status.Response  } } }, @{ n = 'Start' ; e = { $_.Start.DateTime } }, @{ n = 'End' ; e = { $_.End.DateTime } } , @{ n = 'ResponseSatus' ; e = { $_.ResponseStatus.Response } }, @{ n = 'ResponseTime' ; e = { $_.ResponseStatus.Time } } | Export-Csv -NoTypeInformation $filePath


#region Cleanup

Get-PSSession | Remove-PSSession
Remove-Variable Session

#endregion


}