#Calendar


Param (

    
    [Parameter(Mandatory = $true,Position = 1 )]
    [string]$FilePath,
    
    [Parameter(Mandatory = $true)]
    [datetime]$StartDate,
    
    [Parameter(Mandatory = $true)]
    [datetime]$EndDate


)

Get-PSSession | Remove-PSSession
 
#Varibles
$restUri = 'https://outlook.office365.com/api/beta/users'
$UserCredential = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber
$users = Get-Mailbox -Filter { RecipientTypeDetails -ne 'DiscoveryMailbox'  } -SortBy DisplayName

$Start = Get-Date $StartDate -Format yyyy-MM-dd
$End = Get-Date $EndDate -Format yyyy-MM-dd

$select = 'Organizer,Attendees,ResponseStatus,Start,End,Subject'
$filter = "CreatedDateTime ge $Start and CreatedDateTime le $End"


#Checking time


#Checking time - Powershell Switch
<#
if ( $startDate = $null ) {

    $filter = "`$filter=ReceivedDateTime le $endDate"

}

elseif ( $endDate = $null ) {

    $filter = '`$filter=ReceivedDateTime ge $startDate'

}



if ( $startDate -eq $null -and $endDate -eq $null ) {

    $filter = $null

}




#Start Date should not be newer than end date

if ( $endDate > $startDate ) {

    Write-Error -Message 'Start date should be newer than end date' -WarningAction Stop -Category InvalidData

}


#>



$results = @()

foreach ( $user in $users | ForEach-Object UserPrincipalName ) {
    
    $top = 25
    $skip = 0
     
    do {          
         $batch = Invoke-RestMethod -Method Get -Uri "$restUri/$user/calendarview?StartDateTime=$($Start)T01:00:00&EndDateTime=$($End)T23:00:00&`$skip=$skip" -Credential $UserCredential
         $results += $batch.value | select *, @{ Name = 'User' ; Expression ={ $user  }  } #, @{ Name = 'CalendarName' ; Expression = { ( Invoke-RestMethod -Method GET -Uri $_.'Calendar@odata.navigationLink' -Credential $UserCredential ).Name }  }
         $skip += 25
       }  until ( $batch.'@odata.nextLink' -eq $null )
}

$results | select id, !@{ Name = 'Meeting Title' ; Expression = { $_.Subject } }, user, @{ n = 'Organizer' ; e = { $_.Organizer.EmailAddress.Address } }, @{ n = 'Attendee:Type:ResponseTime:ResponseStatus' ; e = { $_.Attendees | %{ $_.EmailAddress.Address + ':' + $_.Type + ':' + $_.Status.Time + ':' + $_.Status.Response  } } }, @{ n = 'Start' ; e = { $_.Start.DateTime } }, @{ n = 'End' ; e = { $_.End.DateTime } } , @{ n = 'ResponseSatus' ; e = { $_.ResponseStatus.Response } }, @{ n = 'ResponseTime' ; e = { $_.ResponseStatus.Time } } | Export-Csv -NoTypeInformation $filePath
 
