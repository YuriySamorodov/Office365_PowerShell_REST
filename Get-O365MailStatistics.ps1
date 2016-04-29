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
    [string]$FilePath = '',
    
    [Parameter(Mandatory = $false)]
    [datetime]$StartDate = 0,
    
    [Parameter(Mandatory = $false)]
    [datetime]$EndDate = ( Get-Date ).AddDays(1)


)


#Varibles
    $restUri = 'https://outlook.office365.com/api/beta/users'
    $UserCredential = Get-Credential

    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    Import-PSSession $Session -AllowClobber
$users = Get-Mailbox -Filter { RecipientTypeDetails -ne 'DiscoveryMailbox'  } -SortBy DisplayName

$Start = Get-Date $StartDate -Format yyyy-MM-dd
$End = Get-Date $EndDate -Format yyyy-MM-dd

$select = 'SentDateTime,ReceivedDateTime,Sender,ToRecipients,BCCRecipients,Subject'
$filter = "ReceivedDateTime ge $Start and ReceivedDateTime le $End"


#Checking time

$startDate = Get-Date '01/01/2015' -Format yyyy-MM-dd
$endDate = Get-Date '03/10/2016' -Format yyyy-MM-dd


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

    #$mailbatch = Invoke-RestMethod -Uri "$restUri/$user/messages?`$filter=$filter&`$select=$select&`$top=$top" -Credential $UserCredential -Method Get ; $results = $mailbatch.value
     
    do {          
         $batch = Invoke-RestMethod -Method Get -Uri "$restUri/$user/messages?`$filter=$filter&`$select=$select&`$top=$top&`$skip=$skip" -Credential $UserCredential
         $results += $batch.value
         $skip += 25
       
       }

      until ( $batch.'@odata.nextLink' -eq $null )

}


$results | select SentDateTime, ReceivedDateTime, @{ n = 'Sender' ; e = { $_.Sender.EmailAddress.Address } }, @{ n = 'ToRecipients' ; e = { $_.ToRecipients.EmailAddress | %{ $_.Address } } }, Subject | Export-csv -NoTypeInformation $FilePath

#region Cleanup

Get-PSSession | Remove-PSSession
Remove-Variable Session

#endregion