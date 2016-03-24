#Mail

#Varibles

<<<<<<< HEAD
$Login = 'viastak@bie-executive.com'
=======
$Login = 'viastak@bie-executive.com-'
>>>>>>> 4922930347455599da00e27c992dea2ac0380ec1
$Password = 'C1sP4l6*1' | ConvertTo-SecureString -AsPlainText -Force
$UserCredential = New-Object System.Management.Automation.PSCredential( $Login , $Password )
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session
$restUri = 'https://outlook.office365.com/api/beta/users'


<<<<<<< HEAD
#$users = Get-Mailbox

$startDate = Get-Date '01/01/2015' -Format yyyy-MM-dd
$endDate = Get-Date '01/10/2016' -Format yyyy-MM-dd

$search = '`$select=SendTime,ReceivedDateTime,Sender,ToRecipients,BCCRecipients,Subject'


#Checking time
=======
$users = Get-Mailbox

$startDate = Get-Date '01/01/2015' -Format yyyy-MM-dd
$endDate = Get-Date '03/10/2016' -Format yyyy-MM-dd

$select = 'SentDateTime,ReceivedDateTime,Sender,ToRecipients,BCCRecipients,Subject'


#Checking time - Powershell Switch
>>>>>>> 4922930347455599da00e27c992dea2ac0380ec1

if ( $startDate = $null ) {

    $filter = "`$filter=ReceivedDateTime le $endDate"

}

elseif ( $endDate = $null ) {

<<<<<<< HEAD
    $filter = "`$filter=ReceivedDateTime ge $startDate"

}

$filter = "`$filter=ReceivedDateTime ge $startDate and ReceivedDateTime le $endDate"
=======
    $filter = '`$filter=ReceivedDateTime ge $startDate'

}

$filter = "ReceivedDateTime ge $startDate and ReceivedDateTime le $endDate"
>>>>>>> 4922930347455599da00e27c992dea2ac0380ec1


if ( $startDate -eq $null -and $endDate -eq $null ) {

    $filter = $null

}




#Start Date should not be newer than end date

if ( $endDate > $startDate ) {

    Write-Error -Message 'Start date should be newer than end date' -WarningAction Stop -Category InvalidData

}


<<<<<<< HEAD
foreach ( $user in 'ben.hawkins@bie-executive.com' ) {
    
    $results = @()
        
    $mailbatch = Invoke-RestMethod -Uri "$restUri/$user/messages?$filter" -Credential $UserCredential -Method Get ; $results = $mailbatch.value
     
    do { $mailbatch = Invoke-RestMethod -Uri $mailbatch.'@odata.nextLink' -Credential $UserCredential -Method Get ; $results += $mailbatch.value
=======
$results = @()
 #test
foreach ( $user in Get-Mailbox | ForEach-Object UserPrincipalName ) {
    
    #$results = @()


    $top = 25
    $skip = 0

    #$mailbatch = Invoke-RestMethod -Uri "$restUri/$user/messages?`$filter=$filter&`$select=$select&`$top=$top" -Credential $UserCredential -Method Get ; $results = $mailbatch.value
     
    do { Write-Host $user
         
         $mailbatch = Invoke-RestMethod  -Uri "$restUri/$user/messages?`$top=$top&`$skip=$skip" -Credential $UserCredential -Method Get
         $results += $mailbatch.value
         $skip += 25
>>>>>>> 4922930347455599da00e27c992dea2ac0380ec1
      
      }

      until ( $mailbatch.'@odata.nextLink' -eq $null )

<<<<<<< HEAD
}

$results| select SentDateTime, ReceivedDateTime, @{ n = 'Sender' ; e = { $_.Sender.EmailAddress.Address } }, @{ n = 'ToRecipients' ; e = { $_.ToRecipients.EmailAddress | %{ $_.Address } } }, Subject
=======

$results | select @{Name = 'Mailbox' ; Expression = { $user } }, SentDateTime, ReceivedDateTime, @{ n = 'Sender' ; e = { $_.Sender.EmailAddress.Address } }, @{ n = 'ToRecipients' ; e = { $_.ToRecipients.EmailAddress | %{ $_.Address } } }, Subject

}

>>>>>>> 4922930347455599da00e27c992dea2ac0380ec1
