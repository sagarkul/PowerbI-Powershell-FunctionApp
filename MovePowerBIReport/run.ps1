using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "Move PowerBI Report function processed a request."
try {
    # Interact with query parameters or the body of the request.
    $source_group_id = $Request.Query.sourceGroupId
    $source_report_id = $Request.Query.sourceReportId
    $target_group_name = $Request.Query.targetGroupName

    if (-not $source_report_id ) {
        $source_group_id = $Request.Body.sourceGroupId
    }
    if (-not $source_report_id ) {
        $source_report_id = $Request.Body.sourceReportId
    }
    if (-not $target_group_name  ) {
        $target_group_name = $Request.Body.targetGroupName
    }
    if (-not $source_report_id -or 
        -not $source_report_id -or 
        -not $target_group_name ) {
        throw "missing parameter. make sure you provide sourceGroupId,sourceReportId and targetGroupName."
    }


    "source_group_id: $source_group_id"

    # $body = "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response."

    # if ($name) {
    #     $body = "Hello, $name. This HTTP triggered function executed successfully."
    # }
    
    if (-not (Get-Module PowerBIPS)) {
        "Import-Module PowerBIPS -UseWindowsPowerShell"
        Import-Module PowerBIPS -UseWindowsPowerShell
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        #[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls, [Net.SecurityProtocolType]::Tls11, [Net.SecurityProtocolType]::Tls12, [Net.SecurityProtocolType]::Ssl3
    }

    # Authenticate with service principal
    $clientId = "871342d7-cb0e-44ef-8aa6-a73520a5394a" 
    $clientSecret = "dLB7Q~w67nLe9v.v3m0gBIn4jSl~2BZ2zrD7r"
    $tenant = "8f102725-4bbc-4a0f-ba03-1a0e913cd18e"

    #PowerShell – The underlying connection was closed: An unexpected error occurred on a send.
    #https://blog.darrenjrobinson.com/powershell-the-underlying-connection-was-closed-an-unexpected-error-occurred-on-a-send/
    $authToken = Get-PBIAuthToken -clientId $clientId -clientSecret $clientSecret -tenantId $tenant -Verbose
    
    "authToken: $authToken"

    if (-not ($authToken)) {
        throw "Unable to create authToken for client: $clientId"
    }

    # get source workspace
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor
    [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12
    $source_group = Get-PBIGroup -authToken $authToken -id $source_group_id
    if (-not ($source_group)) {
        throw "No source group with given id: $source_group_id found."
    }

    # Get report object
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor
    [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12
    $source_report = Get-PBIReport -authToken $authToken -id $source_report_id -groupId $source_group_id
    if (-not ($source_report)) {
        throw "No source report with given id: $source_report_id found"
    }
    
    # Get target workspace
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor
    [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12
    $target_group = Get-PBIGroup -authToken $authToken -name $target_group_name
    if (-not ($target_group)) {
        throw "No target group with given name : $target_group_name found."
    }

    $target_group_id = $($target_group.id)
    
    $target_report = Get-PBIReport -authToken $authToken -name "$($source_report.name)" -groupId $target_group_id 
    
    "target_report: $target_report"
    if ($target_report) {
        # update existing report
        "update existing report"
        $target_report_id = $($target_report.id)
        "Set-PBIReportContent -authToken $authToken -groupId $source_group_id -report $source_report_id -targetGroupId $target_group_id -targetReportId $target_report_id"
        $target_report = Set-PBIReportContent -authToken "$authToken" -report $source_report_id -groupId $source_group_id -targetGroupId $target_group_id -targetReportId $target_report_id 
    }
    else {
        # create copy of report
        "create copy of report"
        $target_report = Copy-PBIReports -authToken "$authToken" -groupId $source_group_id -report $source_report_id -targetWorkspaceId "$target_group_id" -targetModelId "$($source_report.datasetId)"  
    } 
        
    if (-not ($target_report)) {
        throw "Unable to copy report to workspace: $target_group_name"
    }
    
    $body = $target_report
}
catch { 
    Write-Host "An error occurred:"
    Write-Host $_.Exception
    Write-Host $_.ScriptStackTrace
}
finally {
    if ($Error) {
        # Associate values to output bindings by calling 'Push-OutputBinding'.
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
                StatusCode = [HttpStatusCode]::BadRequest
                Body       = $Error
            })    
    }
    else {
        # Associate values to output bindings by calling 'Push-OutputBinding'.
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
                StatusCode = [HttpStatusCode]::OK
                Body       = $body
            })    
    }
}