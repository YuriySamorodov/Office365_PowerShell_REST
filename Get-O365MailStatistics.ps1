﻿#Mail

#Varibles
$restUri = 'https://outlook.office365.com/api/beta/users'
$Login = 'viastak@bie-executive.com'
$Password = 'C1sP4l6*1' | ConvertTo-SecureString -AsPlainText -Force
$UserCredential = New-Object System.Management.Automation.PSCredential( $Login , $Password )

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session
$users = Get-Mailbox -Filter { RecipientTypeDetails -ne 'DiscoverySearch'  } -SortBy UserPrincipalName 

$startDate = Get-Date '01/01/2015' -Format yyyy-MM-dd
$endDate = Get-Date '01/10/2017' -Format yyyy-MM-dd

$select = 'SentDateTime,ReceivedDateTime,Sender,ToRecipients,BCCRecipients,Subject'
$filter = "ReceivedDateTime ge $startDate and ReceivedDateTime le $endDate"


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
     
    do { Write-Host $user
         
         $batch = Invoke-RestMethod  -Uri "$restUri/$user/messages?``$filter=$filter&`$select=$select&`$top=$top&`$skip=$skip" -Credential $UserCredential -Method Get
         $results += $batch.value
         $skip += 25

      }

      until ( $batch.'@odata.nextLink' -eq $null )

}


$results | select SentDateTime, ReceivedDateTime, @{ n = 'Sender' ; e = { $_.Sender.EmailAddress.Address } }, @{ n = 'ToRecipients' ; e = { $_.ToRecipients.EmailAddress | %{ $_.Address } } }, Subject
