<#
 .Synopsis
  IT Admin can use this PowerShell module to clear the specific messages or events from the mailbox based on the filter or search operator.

 .Description
  IT Admin can use this PowerShell module to clear the specific messages or events from the mailbox based on the filter or search operator.
https://learn.microsoft.com/en-us/graph/search-query-parameter?tabs=http
1. For the messages, Search can use the following properties:bccRecipients,body,ccRecipients,toRecipients, Attachments
2. Can't use the search operator on the events
https://learn.microsoft.com/en-us/graph/filter-query-parameter?tabs=http
1. For the messages, Filter can use the following properties: createdDateTime,from,hasAttachments,importance,inferenceClassification,internetMessageId,parentFolderId,receivedDateTime,sender,subject
2. for the events, can't use the Filter operator on the organizer property, but can use the filter operator on the start and end properties.

 .Example
   # Installl and import this PowerShell Module   

 .Example
    using module .\ZIZHUGraphEXO.psm1;
    $userId = 'leeg@vjqg8.onmicrosoft.com';
    $tenantId = "cff343b2-f0ff-416a-802b-28595997daa2";
    $clientId = "51d1eabe-88f4-4082-aaaa-3f94470baec8";
    $tb = '4A4F1E95449F05319B156B9DF619B6C0E5355236';
    # can also use the client secret
    $cert = get-item "cert:\localmachine\my\$tb";
    Connect-MgGraph -TenantId "$tenantId" -ClientId $clientId -Certificate $cert;
    $records = New-Object 'Collections.Generic.List[MailData]';
    Clear-SpecificMessagesfromMailbox -userid $userid -queryString "subject eq 'Microsoft Entra ID Protection Weekly Digest'" -processingMode "deleteditems" -isFilter $true -records $records -maxMails 120;
    $records | Format-List;
    Clear-SpecificMessagesfromMailbox -userid $userid -queryString '"body:70f12e5b-e7a4-4ef3-950e-fcb58d5f0534"' -processingMode "softdelete" -isFilter $false -records $records -maxMails 20;
    $records | Format-List;
    Clear-SpecificMessagesfromMailbox -userid $userid -queryString '"attachment:GraphLog.xlsx"' -processingMode "harddelete" -isFilter $false -maxMails 20 -records $records;
    $records | Format-List;
    $userId= 'leeg@vjqg8.onmicrosoft.com';
    $startDate = "2025-03-01T00:00:00Z"
    $endDate = "2025-03-31T23:59:59Z"
    $senderEmail = 'testuser@onms.com'
    $filter = "from/emailAddress/address eq '$senderEmail' and receivedDateTime ge $startDate and receivedDateTime le $endDate"
    Clear-SpecificMessagesfromMailbox -userid $userid -queryString $filter -processingMode "junk" -isFilter $true -maxMails 5 -records $records;
    $records | Format-List;
    $records = New-Object 'Collections.Generic.List[EventData]';
    Clear-SpecificEventsfromMailbox -userid $userId -filterString "subject eq 'event2025040603'" -processingMode "softdelete" -maxEvents 5 -records $records;
    $records = New-Object 'Collections.Generic.List[EventData]';
    Clear-SpecificEventsfromMailbox -userid $userId -filterString "isCancelled eq true" -processingMode "harddelete" -maxEvents 5000 -records $records;
    $records = New-Object 'Collections.Generic.List[EventData]';
    $records | Format-List;

    $startDate = "2025-04-01T00:00:00Z"
    $endDate = "2025-04-10T23:59:59Z"
    $organizerEmail = 'testuser@onms.com'
    $filter = "start/dateTime ge '$startDate' and start/dateTime le '$endDate'"
    Clear-SpecificEventsfromMailbox -userid $userid -filterString $filter -processingMode "softdelete" -maxEvents 50 -records $records -organizerEmail $organizerEmail;
    $records | Format-List;
    Disconnect-Graph;
#>
Class MailData {
    [String]$Subject
    [String]$Id
    [String]$Sender
    [String]$ReceivedDateTime
    [String]$ParentFolderId
}
Class EventData {
    [String]$Subject
    [String]$Id
    [String]$Organizer
    [String]$CreatedDateTime
    [String]$Type
}
<#
.SYNOPSIS
Remove the specific messages from the mailbox based on the filter or search operator

.DESCRIPTION
Remove the specific messages from the mailbox based on the filter or search operator

.PARAMETER userid
user id of the mailbox to be processed

.PARAMETER queryString
the query string to be used for filtering or searching the messages

.PARAMETER processingMode
the processing mode to be used for the messages. It can be one of the following: 'junk', 'deleteditems', 'softdelete', or 'harddelete'

.PARAMETER isFilter
decide to use the filter or search operator. If true, use the filter operator; otherwise, use the search operator

.PARAMETER records
the processed messages will be stored in this list. If not provided, a new list will be created

.PARAMETER maxMails
the maximum number of messages to be processed. If the number of found messages is greater than this value, the processing will be stopped

.PARAMETER pageSize
the page size to be used for the Graph list messages. The default value is 50

.EXAMPLE
$records = New-Object 'Collections.Generic.List[MailData]';
Clear-SpecificMessagesfromMailbox -userid $userid -queryString "subject eq 'Microsoft Entra ID Protection Weekly Digest'" -processingMode "deleteditems" -isFilter $true -records $records -maxMails 120;
$records | Format-List;
Clear-SpecificMessagesfromMailbox -userid $userid -queryString '"body:70f12e5b-e7a4-4ef3-950e-fcb58d5f0534"' -processingMode "softdelete" -isFilter $false -records $records -maxMails 20;
$records | Format-List;
Clear-SpecificMessagesfromMailbox -userid $userid -queryString '"attachment:GraphLog.xlsx"' -processingMode "harddelete" -isFilter $false -maxMails 20 -records $records;
$records | Format-List;
$userId= 'leeg@vjqg8.onmicrosoft.com';
$startDate = "2025-03-01T00:00:00Z"
$endDate = "2025-03-31T23:59:59Z"
$senderEmail = 'doqi@microsoft.com'
$filter = "from/emailAddress/address eq '$senderEmail' and receivedDateTime ge $startDate and receivedDateTime le $endDate"
Clear-SpecificMessagesfromMailbox -userid $userid -queryString $filter -processingMode "junk" -isFilter $true -maxMails 5 -records $records;
$records | Format-List;

.NOTES
General notes
#>
function Clear-SpecificMessagesfromMailbox {
    [cmdletbinding()]
    param (
        [Parameter(Position = 0, Mandatory = $True)]
        [String] $userid,
        [Parameter(Position = 1, Mandatory = $True)]
        [String]$queryString,
        [Parameter(Position = 2, Mandatory = $True)]
        [ValidateSet("junk", "deleteditems", "softdelete", "harddelete")]
        [String]$processingMode,
        [Parameter(Position = 3, Mandatory = $True)]
        [bool]$isFilter,
        [Parameter(Position = 4, Mandatory = $false)]
        [System.Collections.Generic.List[MailData]]$records,
        [Parameter(Position = 5, Mandatory = $false)]
        [int]$maxMails = 20,
        [Parameter(Position = 6, Mandatory = $false)]
        [int]$pageSize = 50
    )
    $headers = @{Prefer = "IdType=`"ImmutableId`"" };
    $deletedItemsFolder = Get-MgUserMailFolder -UserId $userId -Filter "displayName eq 'Deleted Items'" -Headers $headers;
    $deletedItemsFolderId = $deletedItemsFolder.Id;
    $junkFolder = Get-MgUserMailFolder -UserId $userId -Filter "displayName eq 'Junk Email'" -Headers $headers;
    $junkFolderId = $junkFolder.Id;
    $sentItemsFolder = Get-MgUserMailFolder -UserId $userId -Filter "displayName eq 'Sent Items'" -Headers $headers;
    $sentItemsFolderId = $sentItemsFolder.Id;
    $specificFolderList = @($deletedItemsFolderId, $junkFolderId, $sentItemsFolderId);

    $messages = $null;
    if ($isFilter) {
        $messages = @(Get-MgUserMessage -UserId $userid -all -Filter $queryString -Property "id,from,parentFolderId,ToRecipients,ccRecipients,ReceivedDateTime,Sender,subject" -Headers $headers -PageSize $pagesize);
    }
    else {
        $messages = @(Get-MgUserMessage -UserId $userid -all -Property "id,from,parentFolderId,ToRecipients,ccRecipients,ReceivedDateTime,Sender,subject" -Headers $headers -PageSize $pagesize -Search $queryString);
    }
    $messages = @($messages | Where-Object { $_.ParentFolderId -notin $specificFolderList });
    if ($messages.Count -gt $maxMails) {
        $foundCount = $messages.Count;
        Write-Host "The number $foundCount of found messages to be processed is greater than the maximum limit of $maxMails. Please check the query string or adjust the parameter maxMails to proceed."
        return;
    }
    if ($null -eq $messages) {
        Write-Host "No messages found to be processed. Please check the query string to proceed."
        return;
    }
    if ($null -eq $records) {
        $records = New-Object Collections.Generic.List[MailData];
    }
    $messages | ForEach-Object {
        $message = $PSItem;
        # Check the to recipiet
        if ($processingMode -eq "junk") {
            Move-MgUserMessage -UserId $userid -MessageId $message.Id -Confirm:$false -DestinationId $junkFolderId;
        }
        elseif ($processingMode -eq "deleteditems") {
            Move-MgUserMessage -UserId $userid -MessageId $message.Id -Confirm:$false -DestinationId $deletedItemsFolderId;
        }
        elseif ($processingMode -eq "softdelete") {
            Remove-MgUserMessage -UserId $userid -MessageId $message.Id -Confirm:$false;
        } 
        elseif ($processingMode -eq "harddelete") {
            $url = "https://graph.microsoft.com/v1.0/users/$($userid)/messages/$($message.Id)/permanentDelete";
            Invoke-MgGraphRequest -Method POST $url;
        }
        else {
            Write-Host "Invalid processing mode. Please use 'junk', 'deleteditems', 'softdelete', or 'harddelete'."
            return;
        }
        $movedMsg = $message;
        $data = New-Object MailData;
        $data.Id = $movedMsg.Id;
        $data.Subject = $movedMsg.Subject;
        $data.Sender = $movedMsg.Sender.EmailAddress.Address;
        $data.ReceivedDateTime = $movedMsg.ReceivedDateTime;
        $data.ParentFolderId = $movedMsg.ParentFolderId;
        $records.Add($data);
    }
    return;
}
<#
.SYNOPSIS
Remove the specific events from the mailbox based on the filter operator

.DESCRIPTION
Remove the specific events from the mailbox based on the filter operator

.PARAMETER userid
user id of the mailbox to be processed

.PARAMETER filterString
the filter string to be used for filtering the events

.PARAMETER processingMode
the processing mode to be used for the events. It can be one of the following: 'softdelete', or 'harddelete'

.PARAMETER records
the processed events will be stored in this list. If not provided, a new list will be created

.PARAMETER maxEvents
the maximum number of events to be processed. If the number of found events is greater than this value, the processing will be stopped

.PARAMETER pageSize
the page size to be used for the Graph list events. The default value is 50

.PARAMETER organizerEmail
the email address of the organizer to be used for filtering the events. If not provided, all events will be processed

.EXAMPLE
$records = New-Object 'Collections.Generic.List[EventData]';
Clear-SpecificEventsfromMailbox -userid $userId -filterString "subject eq 'event2025040603'" -processingMode "softdelete" -maxEvents 5 -records $records;
$records = New-Object 'Collections.Generic.List[EventData]';
Clear-SpecificEventsfromMailbox -userid $userId -filterString "isCancelled eq true" -processingMode "harddelete" -maxEvents 5000 -records $records;
$records = New-Object 'Collections.Generic.List[EventData]';
$records | Format-List;

$startDate = "2025-04-01T00:00:00Z"
$endDate = "2025-04-10T23:59:59Z"
$organizerEmail = 'doqi@microsoft.com'
$filter = "start/dateTime ge '$startDate' and start/dateTime le '$endDate'"
Clear-SpecificEventsfromMailbox -userid $userid -filterString $filter -processingMode "softdelete" -maxEvents 50 -records $records -organizerEmail $organizerEmail;
$records | Format-List;

.NOTES
General notes
#>
function Clear-SpecificEventsfromMailbox {
    [cmdletbinding()]
    param (
        [Parameter(Position = 0, Mandatory = $True)]
        [String] $userid,
        [Parameter(Position = 1, Mandatory = $True)]
        [String]$filterString,
        [Parameter(Position = 2, Mandatory = $True)]
        [ValidateSet("softdelete", "harddelete")]
        [String]$processingMode,
        [Parameter(Position = 3, Mandatory = $false)]
        [System.Collections.Generic.List[EventData]]$records,
        [Parameter(Position = 4, Mandatory = $false)]
        [int]$maxEvents = 10,
        [Parameter(Position = 5, Mandatory = $false)]
        [int]$pageSize = 50,
        [Parameter(Position = 6, Mandatory = $false)]
        [String]$organizerEmail = $null 
    )
    $headers = @{Prefer = "IdType=`"ImmutableId`"" };  
    $events = $null;
    $events = @(Get-MgUserEvent -UserId $userid -all -Filter $filterString -Property "id,organizer,subject,location,importance,createdDateTime,type" -Headers $headers -PageSize $pagesize);

    if ($events.Count -gt $maxEvents) {
        $foundCount = $events.Count;
        Write-Host "The number $foundCount of found events to be processed is greater than the maximum limit of $maxEvents. Please check the query string or adjust the parameter maxEvents to proceed."
        return;
    }
    if ($PSBoundParameters.ContainsKey('organizerEmail') -and (-not [string]::isnullorwhitespace($organizerEmail))) {
        $events = @($events | Where-Object { $_.Organizer.EmailAddress.Address -eq $organizerEmail });
    }
    if ($null -eq $events -or $events.Count -eq 0) {
        Write-Host "No events found to be processed. Please check the query string to proceed."
        return;
    }
    if ($null -eq $records) {
        $records = New-Object Collections.Generic.List[EventData];
    }
    $events | ForEach-Object {
        $userEvent = $PSItem;
        if ($processingMode -eq "softdelete") {
            Remove-MgUserEvent -UserId $userid -EventId $userEvent.Id -Confirm:$false;
        } 
        elseif ($processingMode -eq "harddelete") {
            $eventId = $userEvent.Id;
            $url = "https://graph.microsoft.com/v1.0/users/$($userid)/events/$($eventId)/permanentDelete";
            Invoke-MgGraphRequest -Method POST $url;            
        }
        else {
            Write-Host "Invalid processing mode. Please use 'softdelete', or 'harddelete'."
            return;
        }
        $removedEvent = $userEvent;
        $data = New-Object EventData;
        $data.Id = $removedEvent.Id;
        $data.Subject = $removedEvent.Subject;
        $data.Organizer = $removedEvent.Organizer.EmailAddress.Address;
        $data.CreatedDateTime = $removedEvent.CreatedDateTime;
        $data.Type = $removedEvent.Type;
        $records.Add($data);
    }
    return;
}
Export-ModuleMember -Function Clear-SpecificMessagesfromMailbox, Clear-SpecificEventsfromMailbox;