Function Get-CherwellApiToken {
    [CmdletBinding()]
    $uri = "$($env:APPSETTING__CherwellBaseUri)token?auth_mode=Internal&api_key$($env:APPSETTING__CherwellApiUser)"
    $header = @{
        'Content-Type' = 'application/x-www-form-urlencoded'
        'Accept' = 'application/json'
    }
    $body = "grant_type=password&client_id=$($env:APPSETTING__CherwellApiClientId)&username=$($env:APPSETTING__CherwellApiUser)&password=$($env:APPSETTING__CherwellApiSecret)"

    Write-Verbose "Retrieving Auth Token from Cherwell"
    try {
        $result = Invoke-WebRequest -Headers $header -Body $body -Uri $uri -Method Post
    }
    catch {
        throw $_
    }

    Return ($result.Content|ConvertFrom-Json).access_token
}

Function Get-CherwellBusObj {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$BusObjName,
        [switch]$Schema,
        [switch]$Template,
        [switch]$RequiredOnly
    ) #end param

    # Check acceptable conditions
    if ($Schema -and $Template) {
        Write-Error "ERROR: Selecting both Schema and Template is not supported"
        Return
    }
    elseif ($RequiredOnly -and -not $Template) {
        Write-Error "ERROR: Can only select RequiredOnly with the Template switch"
        Return
    }

    # Setup Vars
    $baseuri = $env:APPSETTING__CherwellBaseUri
    $header = @{
        'Content-Type' = 'application/json'
        'Accept' = 'application/json'
        'Authorization' = "Bearer $(Get-CherwellApiToken)"
    }

    # Get the business object summary
    try {
        Write-Verbose "Retrieving Business Object Summary"
        $summaryUri = $baseUri + "api/V1/getbusinessobjectsummary/busobname/$($BusObjName)"
        $summaryResults = Invoke-WebRequest  -Uri $summaryUri -Header $header -Method GET
    }
    catch {
        throw $_
    }

    # Get the business object
    $busObReturnedName = ($summaryResults.content|convertfrom-json).displayName
    $busObId = ($summaryResults.content|convertfrom-json).busobId
    Write-Verbose "Set business object to $($busObReturnedName) ($($busObId))"
    if ($Schema) {
        try {
            Write-Verbose "Retrieving Business Object Schema"
            $schemaUri = $baseUri + "api/V1/getbusinessobjectschema/busobid/" + $busObId
            $schemaResults = Invoke-WebRequest -Uri $schemaUri -Header $header -Method GET
        }
        catch {
            throw $_
        }

        Return $schemaResults.content|convertfrom-json
    }
    # Get the business object template
    elseif ($Template) {
        try {
            # Create request for the business object template POST method
            $getTemplateUri = $baseUri + "api/V1/GetBusinessObjectTemplate"
            $templateRequest =
            @{
                busObId = $busObId;
                includeRequired = $true;
                includeAll = if($RequiredOnly) {$false} else {$true}
            } | ConvertTo-Json
            $templateResponse = Invoke-WebRequest -Uri $getTemplateUri -Header $header -Body $templateRequest -Method Post

            Return $templateResponse.content|ConvertFrom-Json
        }
        catch {
            throw $_
        }
    }
    else {
        Return $summaryResults.Content|ConvertFrom-Json
    }
}


Function Get-CherwellConfigItem {
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$CompanyName,
        [string]$SerialNumber,
        [string]$CiRecId,
        [array]$AssetStatus
    )

    #-------------------------------------------------------------------------------
    # Get Customer Data
    #-------------------------------------------------------------------------------

    # Validate search criteria, we must have a Company Name
    if (-not $CompanyName) {
        Write-Error 'ERROR: Search must contain a company name'
        Return
    }

    $baseUri = $env:APPSETTING__CherwellBaseUri
    $requestHeader = @{
        'Content-Type' = 'application/json'
        'Authorization' = "Bearer $(Get-CherwellApiToken)"
    }

    # Get the business object summary
    Write-Host "Getting business object schema"
    $schema = Get-CherwellBusObj -BusObjName 'ConfigurationItem' -Schema
    Write-Host "Set search business object to $(($schema.name)) ($($schema.busObId))"
    $busObId = $schema.busobId

    # Get the fieldIds so we can use it for a search.
    $companyNameField = $schema.fieldDefinitions | Where-Object {$_.displayName -eq "Company Name"}
    $recIdField = $schema.fieldDefinitions | Where-Object {$_.displayName -eq "RecId"}
    $serialNumberField = $schema.fieldDefinitions | Where-Object {$_.displayName -eq "SerialNumber"}
    $assetStatusField = $schema.fieldDefinitions | Where-Object {$_.displayName -eq "Asset Status"}

    # Create company name filter
    if ($CompanyName) {
        Write-Host "Adding Customer filter value of $($CompanyName)"
        $companyNameFilter =
        @{
            fieldId = $($companyNameField.fieldId);
            operator = "eq";
            value = "$CompanyName";
        }
    }
    # Create serial number filter
    if ($SerialNumber) {
        Write-Host "Adding product name filter value of $($SerialNumber)"
        $serialNumberFilter =
        @{
            fieldId = $($serialNumberField.fieldId);
            operator = "eq";
            value = "$SerialNumber"
        }
    }
    # Create RecId filter
    if ($CiRecId) {
        Write-Host "Adding product name filter value of $($SerialNumber)"
        $recIdFilter =
        @{
            fieldId = $($recIdField.fieldId);
            operator = "eq";
            value = "$CiRecId"
        }
    }
    # Create Asset Status filter
    if ($AssetStatus) {
        Write-Host "Adding product name filter value of $($SerialNumber)"
        [System.Collections.ArrayList]$assetStatusFilterList = @()
        foreach ($status in $AssetStatus) {
            [void]$assetStatusFilterList.Add( "$($status.Replace(' ',''))Filter" )
            New-Variable -Name "$($status.Replace(' ',''))Filter" -Force -Value @{
                fieldId = $($assetStatusField.fieldId);
                operator = "eq";
                value = "$status"
            }
        }
    }

    $searchResultsRequest =
    @{
        busObID = $busobId;
        pageSize = 50000;
        fields = @(
            '93442459576dd09c', #Asset Tag
            'BO:9343f8800bd7c,FI:9343f88b85daf1bc3', #Serial Number
            'BO:9343f8800bd7c,FI:937905400191ae67d', #Name
            '9343f93b9a15ddf46f', #Manufacturer
            '940cd0f30a4d90babb', #Warehouse Building
            '9343f93bb67101661f', #Model
            'BO:9343f8800bd7c,FI:93587a248bc7c6b7b7' #Product

        )
        filters = @(
            # add filters if they exist
            if ($companyNameFilter)  {$companyNameFilter}
            if ($serialNumberFilter) {$serialNumberFilter}
            if ($recIdFilter) {$recIdFilter}
            if ($AssetStatus) {
                foreach ($filter in $assetStatusFilterList) {
                    Get-Variable -Name $filter -ValueOnly
                }
            }
            )
    } | ConvertTo-Json

    # Debug output of search request json
    #$searchResultsRequest

    # Run the search
    $searchUri = $baseUri + "api/V1/getsearchresults"
    try {
        Write-Host "Retrieving customer records"
        $searchResponse = Invoke-WebRequest -Uri $searchUri -Header $requestHeader -Body $searchResultsRequest -Method POST
    }
    catch {
        throw $_
    }
    $searchResults = $searchResponse.Content|ConvertFrom-Json

    Return ($searchResults.businessObjects)

}
