############################################################################################
#                                                                                          #
# The sample scripts are not supported under any Microsoft standard support                #
# program or service. The sample scripts are provided AS IS without warranty               #
# of any kind. Microsoft further disclaims all implied warranties including, without       #
# limitation, any implied warranties of merchantability or of fitness for a particular     #
# purpose. The entire risk arising out of the use or performance of the sample scripts     #
# and documentation remains with you. In no event shall Microsoft, its authors, or         #
# anyone else involved in the creation, production, or delivery of the scripts be liable   #
# for any damages whatsoever (including, without limitation, damages for loss of business  #
# profits, business interruption, loss of business information, or other pecuniary loss)   #
# arising out of the use of or inability to use the sample scripts or documentation,       #
# even if Microsoft has been advised of the possibility of such damages                    #
#                                                                                          #
# Author: doqi@microsoft.com                                                               #
############################################################################################
<#
1. Create the mail-enabled security group
2. Grant this AAD application to only access the members of SG for application permission mail.readwrite 
https://learn.microsoft.com/en-us/exchange/permissions-exo/application-rbac#test-service-principal-access

3. Run Connect-Exchangeonline

4. Run the script to move the specific messages to Junk Folder
$mailboxes = Get-DistributionGroupMember -Identity 'TestMailenabledSG@domain.onmicrosoft.com' | Select-Object Identity, Alias,PrimarySmtpAddress;
$startDate = '2025-01-01T22:00:00Z';
$endDate = '2025-02-01T05:00:00Z' ;
$tomailAddress = 'targetDistributionGroup@domain.onmicrosoft.com';
$mailboxes | ForEach-Object{
    $userid = $psitem.PrimarySmtpAddress
    .\CleanSpecificMessagesfromInboxFolder.ps1 -userid $userid -toEmailAddress $tomailAddress -startDate $startDate -endDate $endDate;
}
#>
[CmdletBinding()]
Param (
    [Parameter(Position = 0, Mandatory = $True)]
    [String] $userid,
    [Parameter(Position = 1, Mandatory = $True)]
    [String]$toEmailAddress,
    [Parameter(Position = 2, Mandatory = $True)]
    [String]$startDate,
    [Parameter(Position = 3, Mandatory = $True)]
    [String]$endDate 
)
Class MailData {
    [String]$Id
    [String]$Subject
    [String]$Sender
    [String]$ReceivedDateTime
}
$tenantId = "";
$clientId = "";
$tb = ''
$cert = get-item "cert:\localmachine\my\$tb"
Connect-MgGraph -TenantId "$tenantId" -ClientId $clientId -Certificate $cert;
$headers = @{Prefer = "IdType=`"ImmutableId`"" };
$pagesize = 200;

$logFile = "C:\Temp\$userid.csv";

$inboxFolder = Get-MgUserMailFolder -UserId $userId -Filter "displayName eq 'Inbox'" -Headers $headers;
$inboxFolderId = $inboxFolder.Id;
$junkFolder = Get-MgUserMailFolder -UserId $userId -Filter "displayName eq 'Junk Email'" -Headers $headers;
$junkFolderId = $junkFolder.Id;

$msgFilter = "parentFolderId eq '$inboxFolderId' and receivedDateTime ge $startDate and receivedDateTime lt $endDate";
$messages = @(Get-MgUserMessage -UserId $userid -all -Filter $msgFilter -Property "id,from,parentFolderId,ToRecipients,ccRecipients,subject" -Headers $headers -PageSize $pagesize);
$records = New-Object Collections.Generic.List[MailData];
$messages | ForEach-Object {
    $message = $PSItem;
    # Check the to recipiet
    $isMove = $false;
    for ($i = 0; $i -lt $message.ToRecipients.Count; $i++) {
        if ($message.ToRecipients[$i].EmailAddress.Address -eq $toEmailAddress) {
            $isMove = $True;
            break;  
        }
    }
    if (-not $isMove) {
        for ($j = 0; $j -lt $message.ccRecipients.Count; $j++) {
            if ($message.ccRecipients[$j].EmailAddress.Address -eq $toEmailAddress) {
                $isMove = $True;
                break;  
            }
        }
    }
    if ($isMove) {   
        $movedMsg = Move-MgUserMessage -UserId $userid -MessageId $message.Id -Confirm:$false -DestinationId $junkFolderId;
        $data = New-Object MailData;
        $data.Id = $movedMsg.Id;
        $data.Subject = $movedMsg.Subject;
        $data.Sender = $movedMsg.Sender.EmailAddress.Address;
        $data.ReceivedDateTime = $movedMsg.ReceivedDateTime;
        $records.Add($data); 
    }
}
$records | Export-Csv -Path $logFile -Append;
Disconnect-Graph;
