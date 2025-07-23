<#
Written by Don Morgan
This module exposes the Business Central API via native PowerShell
#>

########## Begin Internal functions ##########
function GetOauthToken{
    <#
    .SYNOPSIS
        Gets an Oauth token using an app registration (client Id, tenant Id, client secret).
    .DESCRIPTION
        Business Central requires the use of Oauth and deprecated basic auth (i.e. API keys/tokens).
    .NOTES
        The access token has a 1h lifetime by default per Entra Id settings.
    .LINK
        Docs on getting an auth token using an app secret: https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-client-creds-grant-flow#first-case-access-token-request-with-a-shared-secret
        Business Central API: https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/administration/automation-apis-using-s2s-authentication
        More API docs: https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/    
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientSecret,
        [Parameter(Mandatory = $true)]
        [string]$ClientId,
        [Parameter(Mandatory = $true)]
        [string]$TenantId
    )

    $bodyParts = @(
        "grant_type=client_credentials",
        "client_id=$ClientId",
        "client_secret=$ClientSecret",
        "scope=https://api.businesscentral.dynamics.com/.default"
    )
    $body =  $bodyParts | Join-String -Separator "&"

    $headers = @{}
    $headers.Add("Content-Type", "application/x-www-form-urlencoded")

    $authUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $req = Invoke-WebRequest -Uri $authUri -Method Post -Body $body -Headers $headers

    $OauthToken = ($req.Content|ConvertFrom-Json).access_token

    Return $OauthToken
}

function InvokeBusinessCentralApi{
    <#
    .SYNOPSIS
        This is the main internal function for this module. It handles making the API calls to a given endpoint, authentication (once connected), etc.
    .DESCRIPTION
        Handles common functionality for calling API endpoints and adding auth headers.

        This function has multiple values for "mode" to facilitate different APIs, including third-party.
        Modes:
        - default. Used for most calls to the standard REST API.
        - NoCompanyContext. Used when enumerating companies or doing environment-level operations.
        - ODataEndpoint. Used for accessing OData APIs (web services).
        - ThirdPartyApi. Used for calling third-party APIs - refer to the "BusinessCentralApiExtras" module for more info.

    .NOTES
        This function is deliberately designed with optional parameters (e.g. ApiPublisher) not being required or validated
    #>
    param(
        [ValidatePattern('/.*')] #require the endpoint start with '/'
        [string]$Endpoint,
        [ValidateSet("Get","Post","Delete","Patch","Upload")]
        [string]$Method = "get",
        [string]$Filepath,
        $Body,
        [ValidateSet("NoCompanyContext","ODataEndpoint","ThirdPartyApi","default")]
        [string]$Mode = "default",
        [ValidatePattern('?.*')] #require the query start with '?'
        [string]$UriQuery,
        #Below parameters are for when using ThirdPartyApi mode
        [string]$ApiPublisher,
        [string]$ApiGroup,
        [ValidateSet("v1.0","v2.0")]
        [string]$ApiVersion
    )

    #Validate that an API token and environment/company contexts are set
    if([string]::IsNullOrEmpty($env:BusinessCentralApiToken)){
        throw 'please run the "Connect-BusinessCentralApi" cmdlet first'
    }
    if([string]::IsNullOrEmpty($env:BusinessCentralApiEnvironmentContext) -and $Mode -ne "NoCompanyContext"){
        throw 'please set an environment context using the "Set-BusinessCentralEnvironmentContext" cmdlet first'
    }
    if(([string]::IsNullOrEmpty($env:BusinessCentralApiCompanyContext) -or [string]::IsNullOrEmpty($env:BusinessCentralApiCompanyNameContext)) -and $Mode -ne "NoCompanyContext"){
        throw 'please set a company context using the "Set-BusinessCentralCompanyContext" cmdlet first'
    }

    $Environment = $env:BusinessCentralApiEnvironmentContext
    $Company = $env:BusinessCentralApiCompanyContext
    $CompanyName = $env:BusinessCentralApiCompanyNameContext
    $TenantId = $env:BusinessCentralApiTenantId

    #Build URL based on the format needed
    switch($Mode){
        "NoCompanyContext"{
            $ApiBaseUrl = "https://api.businesscentral.dynamics.com/v2.0/$Environment/api/v2.0"    
        }
        "ODataEndpoint"{
            $ApiBaseUrl = "https://api.businesscentral.dynamics.com/v2.0/$TenantId/$Environment/ODataV4/Company(`'$CompanyName`')"    
        }
        "ThirdPartyApi"{
            #Some simple validation since ParameterSets aren't possible and dynamicparameters would be overcomplicated
            if(
                [string]::IsNullOrEmpty($ApiPublisher) -or
                [string]::IsNullOrEmpty($ApiGroup) -or
                [string]::IsNullOrEmpty($ApiVersion)
            ){
                throw "Third party API parameters are required when using ThirdPartyApi mode"
            }
            $ApiBaseUrl = "https://api.businesscentral.dynamics.com/v2.0/$TenantId/$Environment/api/$ApiPublisher/$ApiGroup/$ApiVersion/Companies($Company)"
        }
        "default"{
            $ApiBaseUrl = "https://api.businesscentral.dynamics.com/v2.0/$Environment/api/v2.0/companies($Company)"
        }
    }

    $ApiUrl = $ApiBaseUrl + $Endpoint
    if($UriQuery){
        $ApiUrl += $UriQuery
    }

    if($env:BusinessCentralApiVerbosity -eq "debug"){
        Write-Host -ForegroundColor Yellow "Business Central environment: $env:BusinessCentralApiEnvironmentContext"
        Write-Host -ForegroundColor Yellow "API endpoint being called: $Method $ApiUrl"
        Write-Host -ForegroundColor Yellow "API call body: $Body"
    }

    $Token = $env:BusinessCentralApiToken
    $headers = @{
        Authorization = "Bearer $Token"
        Accept        = "application/json"
        "Content-Type" = "application/json"
    }

    switch($Method){
        "Get"{
            $Request = (Invoke-WebRequest -Uri $ApiUrl -Method Get -Headers $headers).content | Convertfrom-Json
        }
        "Post"{
            $Request = Invoke-WebRequest -Uri $ApiUrl -Method Post -Headers $headers -Body $Body
        }
        "Delete"{
            $Request = Invoke-WebRequest -Uri $ApiUrl -Method Delete -Headers $headers
        }
        "Patch"{
            #Need to add the if-match header as it's required for patch calls (updating objects)
            #Seems this is due to potential caching in webservers:
            #https://developer.mozilla.org/en-US/docs/Web/HTTP/Headers/If-Match
            $headers.Add("If-Match", '*')
            $Request = Invoke-WebRequest -Uri $ApiUrl -Method Patch -Headers $headers -Body $Body
        }
        "Upload"{
            #Need to add the if-match header as it's required for patch calls (updating objects)
            #Seems this is due to potential caching in webservers:
            #https://developer.mozilla.org/en-US/docs/Web/HTTP/Headers/If-Match
            $headers.Add("If-Match", '*')
            $Request = Invoke-WebRequest -Uri $ApiUrl -Method Patch -Headers $headers -InFile $Filepath
        }
    }
    
    switch($Mode){
        "ODataEndpoint"{
            if($Method -eq "Post"){
                Return $Request.Content | ConvertFrom-Json
            }
            else{
                Return $Request
            }
        }
        default{
                Return $Request
        }
    }
}

#Debugging cmdlet
function Set-BusinessCentralApiVerbosity{
    param(
        [bool]$Debug
    )

    if($Debug){
        $env:BusinessCentralApiVerbosity = "debug"
        Write-Host -ForegroundColor Green "Business Central debug mode enabled"
    }
    else{
        $env:BusinessCentralApiVerbosity = $null
        Write-Host -ForegroundColor Yellow "Business Central debug mode disabled"
    }
}

########## End Internal Functions ##########

#Tenant and company level cmdlets
function Connect-BusinessCentralApi{
    <#
    .SYNOPSIS
        Connects to Business Central via app registration.
    .NOTES
        Once connected, you will need to set an environment and company context via "Set-BusinessCentralEnvironmentContext" and "Set-BusinessCentralCompanyContext" before using other cmdlets.
    .EXAMPLE
        $clientId = "e32b5db8-a84e-4af2-8bb8-e434382a962d"
        $secret = "blahblahblahsomesecretblahblah~"
        $tenantId = "dd80a757-da1b-442a-a25f-199d5fee6a9e"
        Connect-BusinessCentralApi -ClientSecret $secret -ClientId $clientId -TenantId $tenantId
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientSecret,
        [Parameter(Mandatory = $true)]
        [string]$ClientId,
        [Parameter(Mandatory = $true)]
        [string]$TenantId
    )

    try{
        $OauthToken = GetOauthToken -ClientSecret $ClientSecret -ClientId $ClientId -TenantId $TenantId
    }
    catch{
        Write-Error "Failed to get Oauth token"
    }

    $env:BusinessCentralApiTenantId = $TenantId
    $env:BusinessCentralApiToken = $OauthToken
    #Oauth token default lifetime is 1h
    Write-Host -ForegroundColor Green "Connected - API token good for one hour."
}
function Get-BusinessCentralEnvironment{
    <#
    .SYNOPSIS
        Gets Business Central environments.
    .DESCRIPTION
        Gets environments, e.g. production and sandbox environments.
        Also can be used to get the currently set environment context.
    .NOTES
        This cmdlet doesn't use InvokeBusinessCentralApi since it uses the admin API instead of the normal/application API.
    .EXAMPLE
        #List environments
        Get-BusinessCentralEnvironment
    .EXAMPLE
        #Show current environment context
        Get-BusinessCentralEnvironment -Current
    #>
    param(
        [switch]$Current
    )

    #Get currently set environment
    if($Current){
        return $env:BusinessCentralApiEnvironmentContext
    }
    #List all environments
    else{

        #Validate that an API token is set
        if([string]::IsNullOrEmpty($env:BusinessCentralApiToken)){
            throw 'please run the "Connect-BusinessCentralApi" cmdlet first'
        }

        $Token = $env:BusinessCentralApiToken
        $headers = @{
            Authorization = "Bearer $Token"
            Accept        = "application/json"
            "Content-Type" = "application/json"
        }
    
        $environmentsApiUrl = "https://api.businesscentral.dynamics.com/admin/v2.21/applications/environments"
        $environments = (Invoke-WebRequest -Uri $environmentsApiUrl -Method GET -Headers $headers).content | Convertfrom-Json
    
        Return $environments.value
    }

}
function Set-BusinessCentralEnvironmentContext{
    <#
    .SYNOPSIS
        Sets the environment that further cmdlets should be executed in.
    .DESCRIPTION
        The URI for a given API endpoint includes the environment (e.g. production or sandbox), this cmdlet sets an environment variable that is used by InvokeBusinessCentralApi in subsequent cmdlets so you don't need to specify the environment with each API call.
    .EXAMPLE
        Set-BusinessCentralEnvironmentContext -EnvironmentName "Contoso-Production"
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/administration/tenant-admin-center-environments
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/administration/tenant-environment-topology
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$EnvironmentName
    )

    $env:BusinessCentralApiEnvironmentContext = $EnvironmentName
    Write-Host -ForegroundColor Green "Set Business Central API environment context to $EnvironmentName"
}
function Set-BusinessCentralCompanyContext{
    <#
    .SYNOPSIS
        Sets the Business Central company that further cmdlets should be executed in.
    .DESCRIPTION
        The URI for a given API endpoint includes the company, this cmdlet sets an environment variable that is used by InvokeBusinessCentralApi in subsequent cmdlets so you don't need to specify the company with each API call.
    .EXAMPLE
        $Company = Get-BusinessCentralCompany | Where-Object{$_.name -eq "My Company"}
        Set-BusinessCentralCompanyContext -CompanyId $Company.id
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/about-new-company
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/administration/tenant-environment-topology
    #>
    param(
        [Parameter(Mandatory = $true, ParameterSetName = "ById")]
        [string]$CompanyId,
        [Parameter(Mandatory = $true, ParameterSetName = "ByName")]
        [string]$CompanyName
    )

    switch($PSCmdlet.ParameterSetName){
        "ById"{
            $BcCompany = Get-BusinessCentralCompany | Where-Object{$_.id -eq $CompanyId}
        }
        "ByName"{
            $BcCompany = Get-BusinessCentralCompany | Where-Object{$_.name -eq $CompanyName}
        }
    }

    $env:BusinessCentralApiCompanyNameContext = $BcCompany.name
    $env:BusinessCentralApiCompanyContext = $BcCompany.id
    $CompanyName = $env:BusinessCentralApiCompanyNameContext
    $CompanyId = $env:BusinessCentralApiCompanyContext

    Write-Host -ForegroundColor Green "Set Business Central API company context to $CompanyName ($CompanyId)"
}

#Object level cmdlets
function Get-BusinessCentralCompany{
    <#
    .SYNOPSIS
        Gets companies in a Business Central environment
    .EXAMPLE
        Get-BusinessCentralCompany
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/resources/dynamics_company
    #>
    $Endpoint = "/companies"

    $Companies = InvokeBusinessCentralApi -Endpoint $Endpoint -Mode NoCompanyContext

    Return $Companies.value
}
function Get-BusinessCentralCustomer{
    <#
    .SYNOPSIS
        Gets customer records, or a specific customer by Id.
    .EXAMPLE
        #Get specific customer
        Get-BusinessCentralCustomer -Id 12345678
    .EXAMPLE
        #Get all customers
        Get-BusinessCentralCustomer
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_customer_get
    #>
    param(
        [Parameter(Mandatory = $false)]
        [string]$Id
    )

    $Endpoint = "/customers"
    if($Id){
        $Endpoint += "($Id)"
    }

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint

    if($Id){
        Return $Request
    }
    else{
        Return $Request.value
    }   
}
function New-BusinessCentralCustomer{
    <#
    .SYNOPSIS
        Creates a new Business Central customer record with the given properties.
    .NOTES
        Returns the customer object that was created.
    .EXAMPLE
        $NewCustomerSplat = @{
            Display = "Fabrikam LTd"
            Number = "12345678""
            Type = "Company"
            AddressLine1 = "4321 Somewhere Lane"
            AddressLine2 = "Suite 1"
            City = "Schenectady"
            State = "New York"
            Country = "US"
            PostalCode = "12345"
        }
        $NewCustomer = New-BusinessCentralCustomer @NewCustomerSplat
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_customer_create
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$DisplayName,
        [Parameter(Mandatory = $true)]
        [ValidateSet("Company","Person")]
        [string]$Type,
        #Optional fields below here
        [string]$Number,
        [string]$AddressLine1,
        [string]$AddressLine2,
        [string]$City,
        [ValidatePattern('[A-Z][A-Z]')]
        [string]$State,
        [ValidatePattern('[A-Z][A-Z]')]
        [string]$Country,
        [string]$PostalCode,
        [string]$PhoneNumber,
        [string]$Email,
        [string]$Website,
        [string]$salespersonCode
    )

    #Dynamically create a hashtable with whatever attributes were specified. Have to do this since you can't have a null key value in hashtables and you may not use all params when creating a new object
    $Attributes = @{}
    $Params = $PSBoundParameters
    $Keys = $PsBoundParameters.Keys
    foreach ($Key in $Keys){
        $Attributes.Add($Key,$Params.$Key)
    }

    $Body = $Attributes | ConvertTo-Json

    $Endpoint = "/customers"

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint -Method Post -Body $Body

    Return $Request.content | ConvertFrom-Json
}
function Set-BusinessCentralCustomer{
    <#
    .SYNOPSIS
        Updates a Business Central customer record.
    .DESCRIPTION
        Updates a Business Central customer record. Supports updating single or multiple properties at once.
    .EXAMPLE
        Set-BusinessCentralCustomer -Id 12345678 -DisplayName "New Name"
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_customer_update
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id,
        #Optional fields below here
        [string]$Number,
        [string]$DisplayName,
        [ValidateSet("Company","Person")]
        [string]$Type,
        [string]$AddressLine1,
        [string]$AddressLine2,
        [string]$City,
        [ValidatePattern('[A-Z][A-Z]')]
        [string]$State,
        [ValidatePattern('[A-Z][A-Z]')]
        [string]$Country,
        [string]$PostalCode,
        [string]$PhoneNumber,
        [string]$Email,
        [string]$Website,
        [string]$salespersonCode
    )

    #Dynamically create a hashtable with whatever attributes were specified. Have to do this since you can't have a null key value in hashtables and you may not use all params when creating a new object
    $Attributes = @{}
    $Params = $PSBoundParameters
    $Keys = $PsBoundParameters.Keys
    foreach ($Key in $Keys){
        $Attributes.Add($Key,$Params.$Key)
    }

    $Body = $Attributes | ConvertTo-Json

    $Endpoint = "/customers($Id)"

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint -Method Patch -Body $Body

    if($Request.StatusCode -ne '200'){
        Write-Error "Failed to update customer $Request"
    }
}
function Remove-BusinessCentralCustomer{
    <#
    .SYNOPSIS
        Deletes a Business Central customer record.
    .NOTES
        Does not delete the associated company-type contact record.
    .EXAMPLE
        Remove-BusinessCentralCustomer -Id 12345678
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_customer_delete
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id
    )

    $Endpoint = "/customers($Id)"

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint -Method Delete

    if($Request.StatusCode -ne '204'){
        Write-Error "Failed to delete customer $Request"
    }
}
function Get-BusinessCentralContact{
    <#
    .SYNOPSIS
        Gets Business Central contacts.
    .EXAMPLE
        #Get all contacts    
        Get-BusinessCentralContact

        #Get specific contact by Id
        Get-BusinessCentralContact -Id 12345678
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_contact_get
    #>
    param(
        [Parameter(Mandatory = $false)]
        [string]$Id
    )

    If($Id){
        $Endpoint = "/contacts($Id)"
        
    }
    else{
        $Endpoint = "/contacts"
    }

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint

    if($Id){
        Return $Request    
    }
    else{
        Return $Request.value
    }
}
function New-BusinessCentralContact{
    <#
    .SYNOPSIS
        Creates a new Business Central contact with the given properties.
    .EXAMPLE
        $NewContactSplat = @{
        DisplayName = "Jane Doe"
        Number = "12345678"
        AddressLine1 = "4321 Somewhere Lane"
        AddressLine2 = "Suite 1"
        City = "NYC"
        State = "New York"
        Country = "US"
        PostalCode = "12345"
        MobilePhoneNumber = "800-555-1212"
        Type = "Person"
        CompanyNumber = "87654321"
    }
    $NewContact = New-BusinessCentralContact @NewContactSplat
    .NOTES
        To create a contact associated with a customer (company) using the CompanyNumber property, you must resolve the customer contact number using Get-BusinessCentralContactRelation
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_contact_create
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$DisplayName,
        #Optional fields below here
        [string]$Number,
        [string]$JobTitle,
        [string]$CompanyNumber,
        [ValidateSet("Company","Person")]
        [string]$Type,
        [string]$AddressLine1,
        [string]$AddressLine2,
        [string]$City,
        [ValidatePattern('[A-Z][A-Z]')]
        [string]$State,
        [ValidatePattern('[A-Z][A-Z]')]
        [string]$Country,
        [string]$PostalCode,
        [string]$PhoneNumber,
        [string]$MobilePhoneNumber,
        [string]$Email
    )

    #Dynamically create a hashtable with whatever attributes were specified. Have to do this since you can't have a null key value in hashtables and you may not use all params when creating a new object
    $Attributes = @{}
    $Params = $PSBoundParameters
    $Keys = $PsBoundParameters.Keys
    foreach ($Key in $Keys){
        $Attributes.Add($Key,$Params.$Key)
    }

    $Body = $Attributes | ConvertTo-Json

    $Endpoint = "/contacts"

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint -Method Post -Body $Body

    Return $Request.content | ConvertFrom-Json
}
function Set-BusinessCentralContact{
    <#
    .SYNOPSIS
        Updates a contact in Business Central.
    .EXAMPLE
        Set-BusinessCentralContact -Id 12345678 -CompanyNumber 1029384756
        Set-BusinessCentralContact -Id 12345678 -CompanyNumber $ContactRelations.contactNumber
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_contact_update
        https://learn.microsoft.com/en-us/dynamics365/business-central/application/base-application/enum/microsoft.crm.businessrelation.contact-business-relation-link-to-table#values
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id,
        #Optional fields below here
        [string]$DisplayName,
        [string]$JobTitle,
        [ValidateSet("Company","Person")]
        [string]$Type,
        [string]$CompanyNumber,
        [string]$Number,
        [string]$AddressLine1,
        [string]$AddressLine2,
        [string]$City,
        [ValidatePattern('[A-Z][A-Z]')]
        [string]$State,
        [ValidatePattern('[A-Z][A-Z]')]
        [string]$Country,
        [string]$PostalCode,
        [string]$PhoneNumber,
        [string]$MobilePhoneNumber,
        [string]$Email
    )

    #Dynamically create a hashtable with whatever attributes were specified. Have to do this since you can't have a null key value in hashtables and you may not use all params when creating a new object
    $Attributes = @{}
    $Params = $PSBoundParameters
    $Keys = $PsBoundParameters.Keys
    foreach ($Key in $Keys){
        $Attributes.Add($Key,$Params.$Key)
    }

    $Body = $Attributes | ConvertTo-Json

    $Endpoint = "/contacts($Id)"

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint -Method Patch -Body $Body

    if($Request.StatusCode -ne '200'){
        Write-Error "Failed to update contact $Request"
    }
}
function Remove-BusinessCentralContact{
    <#
    .SYNOPSIS
        Deletes a contact from Business Central.
    .NOTES
        If you delete the company contact (with type: company) for a customer, it will also delete all related person contacts, but not the customer record.
    .EXAMPLE
        Remove-BusinessCentralContact -Id 12345678
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_contact_delete
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id
    )

    $Endpoint = "/contacts($Id)"

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint -Method Delete

    if($Request.StatusCode -ne '204'){
        Write-Error "Failed to delete contact $Request"
    }
}
function Get-BusinessCentralContactRelation{
    <#
    .SYNOPSIS
        Gets contacts related to a given vendor/customer.
    .EXAMPLE
        Get-BusinessCentralContactRelation -CustomerId 12345678
    .NOTES
        When relating person contacts with a company contact, you can use "Get-BusinessCentralContactRelation -CustomerId 12345678 | where-object{$_.contacttype -eq "company"}"
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/resources/dynamics_contactinformation
    #>
    param(
        [parameter(Mandatory = $true,ParameterSetName = "VendorRelation")]
        [string]$VendorId,
        [parameter(Mandatory = $true,ParameterSetName = "CustomerRelation")]
        [string]$CustomerId
    )

    switch($PsCmdlet.ParameterSetName){
        "VendorRelation" {
            $Endpoint = "/vendors($VendorId)/contactsInformation"
        }
        "CustomerRelation" {
            $Endpoint = "/customers($CustomerId)/contactsInformation"
        }
    }

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint

    Return $Request.value
}
function Get-BusinessCentralSalesQuote{
    <#
    .SYNOPSIS
        Gets Business Central sales quotes.
    .EXAMPLE
        #Get all sales quotes.
        Get-BusinessCentralSalesQuote

        #Get specific sales quote by Id
        Get-BusinessCentralSalesQuote -Id 12345678
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_salesquote_get    
    #>
    param(
        [Parameter(Mandatory = $false)]
        [string]$Id
    )

    If($Id){
        $Endpoint = "/salesQuotes($Id)"
    }
    else{
        $Endpoint = "/salesQuotes"
    }

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint

    if($Id){
        Return $Request    
    }
    else{
        Return $Request.value
    }   
}
function New-BusinessCentralSalesQuote{
    <#
    .SYNOPSIS
        Creates a new Business Central sales quotes with the given properties.
    .NOTES
        Returns the sales quote object that was created.
    .EXAMPLE
        $NewSalesQuoteSplat = @{
            CustomerId = 12345678 #note: this is the GUID, not the number
        }
        $SalesQuote = New-BusinessCentralSalesQuote @NewSalesQuoteSplat
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_salesquote_create
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$CustomerId,
        #Optional fields below here
        [string]$Number,
        [string]$BillToName,
        [string]$RequestedDeliveryDate,
        [string]$OrderDate,
        [string]$ExternalDocumentNumber
    )

    #Dynamically create a hashtable with whatever attributes were specified. Have to do this since you can't have a null key value in hashtables and you may not use all params when creating a new object
    $Attributes = @{}
    $Params = $PSBoundParameters
    $Keys = $PsBoundParameters.Keys
    foreach ($Key in $Keys){
        $Attributes.Add($Key,$Params.$Key)
    }

    $Body = $Attributes | ConvertTo-Json

    $Endpoint = "/salesQuotes"

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint -Method Post -Body $Body

    Return $Request.content | ConvertFrom-Json
}
function Set-BusinessCentralSalesQuote{
    <#
    .SYNOPSIS
        Creates a new Business Central sales quote with the given properties.
    .NOTES
        Returns the sales quote object that was created.
    .EXAMPLE
        $NewSalesQuoteSplat = @{
            CustomerId = 12345678 #note: this is the GUID, not the number
        }
        $SalesOrder = Set-BusinessCentralSalesQuote @NewSalesQuoteSplat
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_salesquote_update
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id,
        #Optional fields below here
        [string]$Number,
        [string]$BillToName,
        [string]$RequestedDeliveryDate,
        [string]$OrderDate,
        [string]$ExternalDocumentNumber
    )

    #Dynamically create a hashtable with whatever attributes were specified. Have to do this since you can't have a null key value in hashtables and you may not use all params when creating a new object
    $Attributes = @{}
    $Params = $PSBoundParameters
    $Keys = $PsBoundParameters.Keys
    foreach ($Key in $Keys){
        $Attributes.Add($Key,$Params.$Key)
    }

    $Body = $Attributes | ConvertTo-Json

    $Endpoint = "/salesQuotes/($Id)"

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint -Method Patch -Body $Body

    Return $Request.content | ConvertFrom-Json
}
function Remove-BusinessCentralSalesQuote{
    <#
    .SYNOPSIS
        Deletes a sales quote from Business Central.
    .EXAMPLE
        Remove-BusinessCentralSalesQuote -Id 12345678
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_salesquote_delete
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id
    )

    $Endpoint = "/salesQuotes($Id)"

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint -Method Delete

    if($Request.StatusCode -ne '204'){
        Write-Error "Failed to delete sales quote $Request"
    }
}
function Get-BusinessCentralSalesOrder{
    <#
    .SYNOPSIS
        Gets Business Central sales orders.
    .EXAMPLE
        #Get all sales orders    
        Get-BusinessCentralSalesOrder

        #Get specific sales order by Id
        Get-BusinessCentralSalesOrder -Id 12345678
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_salesorder_get
    #>
    param(
        [Parameter(Mandatory = $false)]
        [string]$Id
    )

    If($Id){
        $Endpoint = "/salesOrders($Id)"
    }
    else{
        $Endpoint = "/salesOrders"
    }

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint

    if($Id){
        Return $Request    
    }
    else{
        Return $Request.value
    }   
}
function New-BusinessCentralSalesOrder{
    <#
    .SYNOPSIS
        Creates a new Business Central sales order with the given properties.
    .NOTES
        Returns the sales order object that was created.
    .EXAMPLE
        $NewSalesorderSplat = @{
            CustomerId = 12345678 #note: this is the GUID, not the number
        }
        $SalesOrder = New-BusinessCentralSalesOrder @NewSalesorderSplat
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_salesorder_create
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$CustomerId,
        #Optional fields below here
        [string]$Number,
        [string]$BillToName,
        [string]$RequestedDeliveryDate,
        [string]$OrderDate,
        [string]$ExternalDocumentNumber
    )

    #Dynamically create a hashtable with whatever attributes were specified. Have to do this since you can't have a null key value in hashtables and you may not use all params when creating a new object
    $Attributes = @{}
    $Params = $PSBoundParameters
    $Keys = $PsBoundParameters.Keys
    foreach ($Key in $Keys){
        $Attributes.Add($Key,$Params.$Key)
    }

    $Body = $Attributes | ConvertTo-Json

    $Endpoint = "/salesOrders"

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint -Method Post -Body $Body

    Return $Request.content | ConvertFrom-Json
}
function Set-BusinessCentralSalesOrder{
    <#
    .SYNOPSIS
        Creates a new Business Central sales order with the given properties.
    .NOTES
        Returns the sales order object that was created.
    .EXAMPLE
        $NewSalesorderSplat = @{
        CustomerId = 12345678 #note: this is the GUID, not the number
        }
        $SalesOrder = New-BusinessCentralSalesOrder @NewSalesorderSplat
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_salesorder_update
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$OrderId,
        #Optional fields below here
        [string]$Number,
        [string]$BillToName,
        [string]$RequestedDeliveryDate,
        [string]$OrderDate,
        [string]$ExternalDocumentNumber
    )

    #Dynamically create a hashtable with whatever attributes were specified. Have to do this since you can't have a null key value in hashtables and you may not use all params when creating a new object
    $Attributes = @{}
    $Params = $PSBoundParameters
    $Keys = $PsBoundParameters.Keys
    foreach ($Key in $Keys){
        $Attributes.Add($Key,$Params.$Key)
    }

    $Body = $Attributes | ConvertTo-Json

    $Endpoint = "/salesOrders/($OrderId)"

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint -Method Patch -Body $Body

    Return $Request.content | ConvertFrom-Json
}
function Remove-BusinessCentralSalesOrder{
    <#
    .SYNOPSIS
        Deletes a sales order from Business Central.
    .EXAMPLE
        Remove-BusinessCentralSalesOrder -Id 12345678
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_salesorder_delete
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id
    )

    $Endpoint = "/salesOrders($Id)"

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint -Method Delete

    if($Request.StatusCode -ne '204'){
        Write-Error "Failed to delete sales quote $Request"
    }
}
function Get-BusinessCentralItem{
    <#
    .SYNOPSIS
        Gets Business Central items.
    .EXAMPLE
        #Get all items    
        Get-BusinessCentralItem

        #Get specific item by Id
        Get-BusinessCentralItem -Id 12345678
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_item_get
    #>
    param(
        [Parameter(Mandatory = $false)]
        [string]$Id
    )

    If($Id){
        $Endpoint = "/items($Id)"
    }
    else{
        $Endpoint = "/items"
    }

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint

    if($Id){
        Return $Request    
    }
    else{
        Return $Request.value
    }   
}
function New-BusinessCentralItem{
    <#
    .SYNOPSIS
        Updates an item with the given properties.
    .NOTES
        Returns the updated item.
    .EXAMPLE
        $ItemSplat = @{
            Number = 12345678
            DisplayName = "Doohickey"
        }
        $ItemUpdate = New-BusinessCentralItem @ItemSplat
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_item_update
    #>
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Inventory","Service","Non-Inventory")]
        [string]$Type,
        [string]$Number,
        [string]$DisplayName,
        [string]$DisplayName2,
        [string]$ItemCategoryId,
        [string]$itemCategoryCode,
        [bool]$Blocked,
        [string]$Gtin,
        [decimal]$Inventory,
        [decimal]$UnitPrice,
        [decimal]$UnitCost,
        [string]$TaxGroupId,
        [string]$TaxGroupCode,
        [string]$BaseUnitOfMeasureId,
        [string]$BaseUnitOfMeasureCode,
        [string]$GeneralProductPostingGroupId,
        [string]$GeneralProductPostingGroupCode,
        [string]$InventoryPostingGroupId,
        [string]$InventoryPostingGroupCode
    )

    #Dynamically create a hashtable with whatever attributes were specified. Have to do this since you can't have a null key value in hashtables and you may not use all params when creating a new object
    $Attributes = @{}
    $Params = $PSBoundParameters
    $Keys = $PsBoundParameters.Keys
    foreach ($Key in $Keys){
        $Attributes.Add($Key,$Params.$Key)
    }

    $Body = $Attributes | ConvertTo-Json

    $Endpoint = "/items"

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint -Method Post -Body $Body

    Return $Request.content | ConvertFrom-Json
}
function Set-BusinessCentralItem{
    <#
    .SYNOPSIS
        Updates an item with the given properties.
    .NOTES
        Returns the updated item.
    .EXAMPLE
        $ItemUpdateSplat = @{
            Id = 10293845 #Note that this is the GUID
            Number = 12345678
            DisplayName = "Doohickey"
        }
        $ItemUpdate = Set-BusinessCentralItem @ItemUpdateSplat
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_item_update
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id,
        #Optional fields below here
        [string]$Number,
        [string]$DisplayName,
        [string]$DisplayName2,
        [ValidateSet("Inventory","Service","Non-Inventory")]
        [string]$Type,
        [string]$ItemCategoryId,
        [string]$itemCategoryCode,
        [bool]$Blocked,
        [string]$Gtin,
        [decimal]$Inventory,
        [decimal]$UnitPrice,
        [decimal]$UnitCost,
        [string]$TaxGroupId,
        [string]$TaxGroupCode,
        [string]$BaseUnitOfMeasureId,
        [string]$BaseUnitOfMeasureCode,
        [string]$GeneralProductPostingGroupId,
        [string]$GeneralProductPostingGroupCode,
        [string]$InventoryPostingGroupId,
        [string]$InventoryPostingGroupCode
    )

    #Dynamically create a hashtable with whatever attributes were specified. Have to do this since you can't have a null key value in hashtables and you may not use all params when creating a new object
    $Attributes = @{}
    $Params = $PSBoundParameters
    $Keys = $PsBoundParameters.Keys
    foreach ($Key in $Keys){
        $Attributes.Add($Key,$Params.$Key)
    }

    $Body = $Attributes | ConvertTo-Json

    $Endpoint = "/items($Id)"

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint -Method Patch -Body $Body

    Return $Request.content | ConvertFrom-Json
}
function Remove-BusinessCentralItem{
    <#
    .SYNOPSIS
        Deletes an item from Business Central.
    .EXAMPLE
        Remove-BusinessCentralItem -Id 12345678
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_item_delete
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id
    )

    $Endpoint = "/salesQuotes($Id)"

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint -Method Delete

    if($Request.StatusCode -ne '204'){
        Write-Error "Failed to delete sales quote $Request"
    }
}
function Get-BusinessCentralPicture{
    <#
    .SYNOPSIS
        Gets Business Central pictures for a given object type.
    .EXAMPLE
        Get-BusinessCentralPicture -Id 12345678 -Type Item
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_picture_get    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [ValidateSet("Item","Customer","Employee","Vendor","Contact")]
        [string]$Type
    )

    $Endpoint = "/$Type`s($Id)/picture"

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint

    Return $Request    
}
function Set-BusinessCentralPicture{
    <#
    .SYNOPSIS
        Sets a Business Central pictures for a given object.
    .EXAMPLE
        Set-BusinessCentralPicture -Id 12345678 -Type Item
    .NOTES
        The Content parameter takes the image as base64 encoded bytes. 
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_picture_get    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [ValidateSet("Item","Customer","Employee","Vendor","Contact")]
        [string]$Type,
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    try{
        $File = Get-ChildItem -File -Path $FilePath
    }
    catch{
        throw "file not found"
    }
    

    $Endpoint = "/$Type`s($Id)/picture/pictureContent"

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint -Method Upload -Filepath $File.FullName

    Return $Request    
}
function Remove-BusinessCentralPicture{
    <#
    .SYNOPSIS
        Removes the Business Central picture for a given object type and Id.
    .EXAMPLE
        Remove-BusinessCentralPicture -Id 12345678 -Type Item
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_picture_delete
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id,
        [Parameter(Mandatory = $true)]
        [ValidateSet("Item","Customer","Employee","Vendor","Contact")]
        [string]$Type
    )

    $Endpoint = "/$Type`s($Id)/picture"

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint -Method Delete

    if($Request.StatusCode -ne '204'){
        Write-Error "Failed to delete picture $Request"
    }
}

#WIP functions
@'

function Get-BusinessCentralSalesOrderLine{
    <#
    .SYNOPSIS
        Gets Business Central sales order lines.
    .EXAMPLE
        #Get all sales order lines     
        Get-BusinessCentralSalesOrderLine -OrderId 12345678

        #Get specific line
        Get-BusinessCentralSalesOrderLine -OrderId 12345678 -Id 1029384
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_salesorderline_get
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$OrderId,
        [Parameter(Mandatory = $false)]
        [string]$Id
    )

    $Endpoint = "/salesOrders($OrderId)/salesOrderLines"
    if($Id){$Endpoint += "($Id)"}

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint

    if($Id){
        Return $Request    
    }
    else{
        Return $Request.value
    }   
}
function Set-BusinessCentralSalesOrderLine{
    <#
    .SYNOPSIS
        Updates a Business Central sales order line.
    .EXAMPLE
        #Update a line
        Set-BusinessCentralSalesOrderLine -OrderId 12345678 -Id 1029384 -Quantity 10
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/api/dynamics_salesorderline_get
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$OrderId,
        [Parameter(Mandatory = $true)]
        [string]$Id,
        #Optional fields below here
        [string]$Sequence,
        [string]$Description,
        [string]$Description2,
        [decimal]$Quantity
    )

    $Endpoint = "/salesOrders($OrderId)/salesOrderLines($Id)"

    $Attributes = @{}
    $Params = $PSBoundParameters
    $Keys = $PsBoundParameters.Keys
    foreach ($Key in $Keys){
        if($Key -ne "OrderId"){
            $Attributes.Add($Key,$Params.$Key)
        }
    }

    $Body = $Attributes | ConvertTo-Json

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint -Body $Body -Method Patch

    if($Request.StatusCode -ne '200'){
        Write-Error "Failed to update sales order line $Request"
    }
}
function Get-BusinessCentralSubscription{
    <#
    .SYNOPSIS
        Gets a list of subscriptions (registered webhooks).
    .EXAMPLE
        #Get all subscriptions
        Get-BusinessCentralSubscription

    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/dynamics-subscriptions
    #>
    param(
        [Parameter(Mandatory = $false)]
        [string]$Id
    )

    $Endpoint = "/subscriptions"

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint -Mode NoCompanyContext

    Return $Request.value
}
function New-BusinessCentralSubscription{
    <#
    .SYNOPSIS
        Registers a new subscription (webhook) in Business Central.
    .EXAMPLE

    .NOTES
        Subscriptions expire after 3d unless renewed. See the Renew-BusinessCentralSubscription cmdlet for more information.
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/dynamics-subscriptions#register-a-webhook-subscription
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$NotificationUrl,
        [Parameter(Mandatory = $true)]
        [ValidateSet("Customers","SalesOrders","SalesInvoices","SalesQuotes","Vendors")]
        [string]$ObjectType,
        [Parameter(Mandatory = $true)]
        [ValidateSet("Created","Deleted","Updated")]
        [string]$ChangeType
    )

    Write-Host -ForegroundColor Yellow ""

    $Endpoint = "/subscriptions"

    $Company = $ENV:BusinessCentralApiCompanyContext
    $Resource = "/api/v2.0/companies($Company)/$ObjectType"

    $Body = @{
        notificationUrl = $NotificationUrl
        resource = $Resource
    } | ConvertTo-Json

    $Request = InvokeBusinessCentralApi -Endpoint $Endpoint -Method Post -Body $Body -Mode NoCompanyContext

    Return $Request
}
function Renew-BusinessCentralSubscription{
    <#
    .SYNOPSIS
        Renews a subscription in Business Central.
    .EXAMPLE

    .NOTES
        Renewing a subscription (like registering a new one) requires a handshake with the webhook being registered, as such you must ensure your webhook is configured to complete the handshake.
    .LINK
        https://learn.microsoft.com/en-us/dynamics365/business-central/dev-itpro/api-reference/v2.0/dynamics-subscriptions#renewing-the-subscription
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id
    )
}

'@
