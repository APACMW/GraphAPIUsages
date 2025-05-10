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
.SYNOPSIS
    This script creates an external connection and schema for Cotoso Test Check data in Microsoft Graph.
    It is easier to follow up the steps in the script than to read the documentation https://learn.microsoft.com/en-us/training/modules/copilot-graph-connectors/
#>
Connect-MgGraph -Scopes @("ExternalConnection.ReadWrite.All", "ExternalItem.ReadWrite.All")
# Get the external connection for Cotoso Test Check. If not exists, create it and set the schema
$externalConnectionId = "CotosoTestCheck";
$externalConnection = Get-MgExternalConnection -ExternalConnectionId $externalConnectionId -ErrorAction SilentlyContinue;
if ($null -eq $externalConnection) {
    $params = @{
        id          = "$($externalConnectionId)"
        displayName = "Cotoso Test Check"
        name        = "Cotoso Test Check"
        description = "Connection to index Cotoso Test Check data"
    }    
    $externalConnection = New-MgExternalConnection -BodyParameter $params;
    if ($null -eq $externalConnection) {
        Write-Host "Failed to create external connection $($externalConnectionId)"
        return;
    }
    $params = @{
        baseType   = "microsoft.graph.externalItem"
        properties = @(
            @{
                name          = "CheckTitle"
                type          = "String"
                isSearchable  = "true"
                isRetrievable = "true"
                isQueryable   = "true"
                labels        = @(
                    "title"
                )
            }
            @{
                name        = "id"
                type        = "String"
                ExactMatch  = "true"
                isQueryable = "true"
            }
            @{
                name          = "Year"
                type          = "Int64"
                isQueryable   = "true"                
                isRetrievable = "true"
            }
            @{
                name          = "AssistanceStatus"
                type          = "String"
                isQueryable   = "true"                
                isRetrievable = "true"
            }
            @{
                name          = "Scope"
                type          = "String"
                isQueryable   = "true"                
                isRetrievable = "true"
            }
            @{
                name          = "SubScope"
                type          = "String"
                isQueryable   = "true"                
                isRetrievable = "true"
            }
            @{
                name          = "KeyResultArea"
                type          = "String"
                isQueryable   = "true"                
                isRetrievable = "true"
            }
            @{
                name          = "InsertedAt"
                type          = "DateTime"
                isQueryable   = "true"                
                isRetrievable = "true"
                labels        = @(
                    "lastModifiedDateTime"
                )
            }
            @{
                name          = "KPI"
                type          = "String"
                isQueryable   = "true"                
                isRetrievable = "true"
            }
            @{
                name          = "taskDescription"
                type          = "String"
                isQueryable   = "true"                
                isRetrievable = "true"
            }
            @{
                name          = "taskOwner"
                type          = "String"
                isQueryable   = "true"                
                isRetrievable = "true"
            }
            @{
                name          = "ControlTypeName"
                type          = "String"
                isQueryable   = "true"                
                isRetrievable = "true"
            }
            @{
                name          = "Conclusion"
                type          = "String"
                isQueryable   = "true"                
                isRetrievable = "true"
            }
        )
    }
    Update-MgExternalConnectionSchema -ExternalConnectionId $externalConnection.id -BodyParameter $params -ResponseHeadersVariable responseHeaders;
    $location = ([string]$responseHeaders.Location);
    $connectionOperationId = $location.Substring($location.LastIndexOf('/') + 1);
    do {
        $response = Get-MgExternalConnectionOperation -ExternalConnectionId $externalConnection.Id -ConnectionOperationId $connectionOperationId;
        Start-Sleep -Seconds 30;
    }while ($response.Status -notin @("completed", "failed"));

    if ($response.Status -eq "failed") {
        Write-Host "Failed to create external connection schema $($externalConnection.Id)"
        return;
    }
}

$groupId = "0f6a79e6-9aba-40ea-8c17-c814f9af5e0c";
$records = Import-Csv -Path 'D:\Temp\Test\MYASSRTEST.csv' -Delimiter ',';
foreach ($record in $records) {
    $CheckTitle = $record.CheckTitle;
    $externalItemId = $record.id;
    $year = [int]$record.Year;
    $AssistanceStatus = $record.AssistanceStatus;
    $scope = $record.Scope;
    $subScope = $record.SubScope;
    $keyResultArea = $record.KeyResultArea;
    $insertedAt = ([datetime]$record.InsertedAt).ToString("yyyy-MM-ddTHH:mm:ss");
    $KPI = $record.KPI;
    $taskDescription = $record.taskDescription;
    $taskOwner = $record.taskOwner;
    $controlTypeName = $record.ControlTypeName;
    $conclusion = $record.Conclusion;
    $data = @($record.CheckTitle, $record.Year, $record.AssistanceStatus, $record.Scope, $record.SubScope, $record.KeyResultArea, $record.InsertedAt, $record.KPI, `
            $record.taskDescription, $record.taskOwner, $record.ControlTypeName, $record.Conclusion, $record.CosoPrincipleName, $record.CosoPrincipleDescription, $record.CosoPrincipleCategory, `
            $record.CosoPrincipleCategoryDescription, $record.CosoPrincipleCategoryType, $record.CosoPrincipleCategoryTypeDescription, $record.CosoPrincipleCategoryTypeSubDescription, $record.CosoPrincipleCategoryTypeSubDescription);
    $content = [string]::Join(';', $data);
    $params = @{
        acl        = @(
            @{
                type       = "group"
                value      = "$($groupId)"
                accessType = "grant"
            }
        )
        properties = @{
            CheckTitle       = "$($CheckTitle)"
            id               = "$($externalItemId)"
            year             = $($year)
            AssistanceStatus = "$($AssistanceStatus)"
            scope            = "$($scope)"
            subScope         = "$($subScope)"
            keyResultArea    = "$($keyResultArea)"
            insertedAt       = "$($insertedAt)"
            KPI              = "$($KPI)"
            taskDescription  = "$($taskDescription)"
            taskOwner        = "$($taskOwner)"
            controlTypeName  = "$($controlTypeName)"
            conclusion       = "$($conclusion)"
        }
        content    = @{
            value = "$($content)"
            type  = "text"
        }
    }

    $body = $params | ConvertTo-Json -Depth 10;
    $itemUrl = "https://graph.microsoft.com/v1.0/external/connections/$($externalConnection.Id)/items/$($externalItemId)";
    $itemUrl = $itemUrl.Replace(" ", "%20");
    $itemUrl = $itemUrl.Replace("#", "%23");
    $itemUrl = $itemUrl.Replace("&", "%26");        
    Invoke-MgGraphRequest -Method PUT -Body $body $itemUrl -ContentType "application/json";    
}
Disconnect-MgGraph;