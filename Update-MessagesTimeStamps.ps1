function Update-MessageTimeStamps {
    <#
    .SYNOPSIS
        Update message time stamp properties

    .DESCRIPTION
        This function will change the time properties for a message to a designated time or just view items in the mailbox (reporting mode)

    .PARAMETER ClientId
        Azure AD registered application id

    .PARAMETER OauthRequestedirectUri
        Azure AD registered RequestedirectUri id

    .PARAMETER TenantName
        O365 Exchange tenant name

    .PARAMETER ImpersonatedMailboxName
        Mailbox you want to impersonate

    .PARAMETER TargetMailbox
        Mailbox being opened by the impersonation acccount

    .PARAMETER MaxItems
        Max items to view in a mailbox

    .PARAMETER ChangingMessageStamps
        Switch for reporting or modification

    .PARAMETER ExchangeVersion
        Version of EWS for Exchange used for the default connection. Default is "Exchange2013_SP1"

    .PARAMETER UserImpersonation
        User default or impersonation credentials

    .PARAMETER EnableEWSTracing
        Enable or disable EWS debug tracing

    .EXAMPLE
       Update-MessageTimeStamps -ClientID xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxxx -UserImpersonation -TargetDomain YourTenant.onmicrosoft.com -TenantName YourTenant.onmicrosoft.com -TargetMailbox UserMailbox@YourTenant.onmicrosoft.com

        This will log on to the specified tenant using impersonation to a target mailbox and just report the messages in the mailbox

    .EXAMPLE
        Update-MessageTimeStamps -ClientID xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxxx -UserImpersonation -TargetDomain YourTenant.onmicrosoft.com -TenantName YourTenant.onmicrosoft.com -TargetMailbox UserMailbox@YourTenant.onmicrosoft.com -ChangingMessageStamps

        This will log on to the specified tenant using impersonation to a target mailbox and just report the messages in the mailbox and change all the time stamps on a message (CreataionTime, SubmissionsTime, DeliveryTime and LastModificationTime)
    
    .EXAMPLE
        Update-MessageTimeStamps -ClientID xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxxx -TargetDomain YourTenant.onmicrosoft.com -TenantName YourTenant.onmicrosoft.com -TargetMailbox UserMailbox@YourTenant.onmicrosoft.com
        
        This will log on to the specified tenant using not using impersonation and log in to the mailbox that the OAUTh token was generated for.

    .NOTES
        Download - https://www.microsoft.com/en-us/download/confirmation.aspx?id=42951
        # Configuration Impresonation - https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-configure-impersonation?redirectedfrom=MSDN
        # http://blog.icewolf.ch/archive/2021/02/06/exchange-managed-api-and-oauth-authentication.aspx
        #>

    [cmdletbinding(DefaultParameterSetName = 'LoginInfo')]
    param(
        [Parameter(Position = 0, ParameterSetName = "LoginInfo")]
        [string]
        $ClientId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",

        [Parameter(ParameterSetName = "LoginInfo")]
        [string]
        $script:oauthRequestedirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient",

        [Parameter(Position = 1, ParameterSetName = "LoginInfo")]
        [string]
        $TargetDomain = 'tenant.onmicrosoft.com',

        [Parameter(Position = 2, ParameterSetName = "LoginInfo")]
        [string]
        $TenantName = "tenant.onmicrosoft.com",

        [Parameter(Position = 3, ParameterSetName = "LoginInfo")]
        [string]
        $TargetMailbox = 'user@tenant.onmicrosoft.com',
        
        [Parameter(ParameterSetName = "LoginInfo")]
        [ValidateSet("Exchange2007_SP1", "Exchange2010", "Exchange2010_SP1", "Exchange2010_SP2", "Exchange2013", "Exchange2013_SP1")]
        [Object]
        $ExchangeVersion = "Exchange2013_SP1",

        [Parameter(ParameterSetName = "MessageModifications")]
        [Int]
        $MaxItems = 100,

        [Parameter(ParameterSetName = "MessageModifications")]
        [ValidateRange(1, 60)]
        [Int32]
        $Seconds = "15",

        [switch]
        $ChangeMessageStamps,

        [switch]
        $UserImpersonation,

        [switch]
        $EnableEWSTracing
    )

    begin {
        Write-Host -ForegroundColor Green "Starting process!"
        $script:messageCounter = 0
        Update-TypeData -TypeName UpdateMessageTimeStamps -DefaultDisplayPropertySet  '#', Subject, ItemClass, CreateDate, TimeReceived, TimeSent, LastModificationTime -DefaultDisplayProperty Subject -DefaultKeyPropertySet CustomProperties -Force
    }

    process {
        # Download URL
        $url = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=42951"

        #Get Token for EWS
        Write-Host -ForegroundColor Yellow "Checking for presence of the MSAL.PS module"
        if (Get-Module -Name MSAL.PS -ListAvailable) {
            Import-Module -Name MSAL.PS 
            Write-Host -ForegroundColor Green "Module found and imported!"
        }
        else {
            Write-Host -ForegroundColor Yellow "MSAL.PS module not found! Installing and importing module"
            Install-module -Name MSAL.PS -Force -AllowClobber -Scope AllUsers -Repository PSGallery -AcceptLicense
            Import-Module -Name MSAL.PS 
            Write-Host -ForegroundColor Yellow "Installation and importing module complete!"
        }

        if ($TenantName -eq 'tenant.onmicrosoft.com') {
            Write-Host -ForegroundColor Red "No valid tenant specified. Unable to obtain OAUTH token. Domain name must be in the correct format of tenant.onmicrosoft.com. Exiting"
            return
        }
        else {
            Write-Host -ForegroundColor Green "Logging on to $($TenantName) and getting an authentication token from a token server"
            $authority = "https://login.microsoftonline.com/$TenantName"
            $scopes = "EWS.AccessAsUser.All"
        
            $msalProperties = @{
                Clientid    = $clientid
                RedirectUri = $script:oauthRequestedirectUri
                Authority   = $authority
                Interactive = $True
                Scopes      = $scopes
            }
        }

        try {
            if (-NOT ($script:oauthRequest)) {
                $script:oauthRequest = Get-MsalToken @msalProperties
            }
        }
        catch {
            Write-Host -ForegroundColor Red "$_"
            return
        }
        
        # Import EWS Dll
        $ewsdll = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")  
        if (Test-Path $ewsdll) { Import-Module $ewsdll }
        else { 
            Write-Host -ForegroundColor Red "This script requires the EWS Managed API 1.2 or later. Please download and install the current version of the EWS Managed API from $url. `r`nExiting Script"
            return
        } 

        # EWS Service connection parameters
        $ewsClient = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList $ExchangeVersion
        $ewsClient.Url = 'https://outlook.office365.com/EWS/Exchange.asmx'
        
        if ($EnableEWSTracing) {
            $ewsClient.TraceEnabled = $true
        }

        if ($UserImpersonation) {
            $ewsClient.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($script:oauthRequest.AccessToken)
            $ewsClient.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $TargetMailbox)
        }
        else {
            $ewsClient.UseDefaultCredentials = $true
            $ewsClient.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$script:oauthRequest.AccessToken
            
        }

        # Connection Status
        Write-Host -ForegroundColor Green "Connecting to URL:" $ewsClient.Url
        Write-Host -ForegroundColor Green "Time Zone:" $ewsClient.TimeZone.StandardName
        Write-Host -ForegroundColor Green "User Agent String:" $ewsClient.UserAgent

        # Inbox folder object
        try {
            $InboxFolder = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox, $TargetMailbox)
            $Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ewsClient, $InboxFolder) 
            Write-Host -ForegroundColor Green "Obtained inbox folder id: $($Inbox.id)"
        }
        catch {
            Write-Host -ForegroundColor Red "$_`r`nNOTE: If you are not using impersonation please specify the -TargetMailbox as the account you are signing in as."
            return
        }

        Write-Host -ForegroundColor Green "Attempting to bind to inbox successful!"
        if ($ChangeMessageStamps) {
            Write-Host -ForegroundColor Cyan "Building message view to modify messages. Please wait for this to complete!"
            $ItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView($maxItems)
            $filterItems = $ewsClient.FindItems($Inbox.Id, $ItemView)  
            
            # Define the property tags
            $PR_CLIENT_SUBMIT_TIME = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0039, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime)
            $PR_CREATION_TIME = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3007, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime)
            $PR_MESSAGE_DELIVERY_TIME = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3590, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime)
            $PR_LAST_MODIFICATION_TIME = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3008, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime)
            $PR_MESSAGE_FLAGS = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0E07, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
            
            foreach ($message in $filterItems.Items) {  
                # Generate a random date for the messages
                $randomDate = Get-Random -Minimum 1 -Maximum 3650
                $date = (Get-Date).AddDays(-$randomDate)

                # Change the dates on the messages
                $message.SetExtendedProperty($PR_CLIENT_SUBMIT_TIME, $date)
                $message.SetExtendedProperty($PR_CREATION_TIME, $date)
                $message.SetExtendedProperty($PR_MESSAGE_DELIVERY_TIME, $date)
                $message.SetExtendedProperty($PR_LAST_MODIFICATION_TIME, $date)
                $message.SetExtendedProperty($PR_MESSAGE_FLAGS, 0)
                $message.IsRead = $false
                $message.Subject = $message.Subject + "."

                # Update the inbox
                $null = $message.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite) # suppressReadReceipts
                Start-Sleep -Seconds $Seconds
                $script:messageCounter ++
            }

            # Display updated messages
            Get-MailMessages -FolderView $Inbox
        }
        else {
            Get-MailMessages -FolderView $Inbox
        }
    }

    end {
        if ($ChangeMessageStamps) { Write-Host -ForegroundColor Green "Process complete! $($script:messageCounter) messaged modified!" }
        else { Write-Host -ForegroundColor Green "Process complete! $($script:messageCounter) messaged read!" }
    }
}

function Get-MailMessages {
    <#
    .SYNOPSIS
        Get Messages

    .DESCRIPTION
        Will retrieve messages from an Exchange mailbox

    .EXAMPLE
        Get-MailMessages

    .NOTES
        Internal function. Must be called from the main function
    #>
    [cmdletbinding()]
    param(
        [Microsoft.Exchange.WebServices.Data.Folder]
        $FolderView
    )
    
    # Generate a list view table of mailbox items (messages)
    $ItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView($maxItems)
    $filterItems = $ewsClient.FindItems($Inbox.Id, $ItemView) 
    $script:messageCounter = 0
    [System.Collections.ArrayList] $messageList = @()

    # Truncate the message subject in the current view
    foreach ($currentMessage in $filterItems.Items) {
        if ($currentMessage.subject.Length -gt 15) {
            $subject = $currentMessage.subject
            $subject = $subject.SubString(0, 5) + "..." 
        }

        $message = [PSCustomObject]@{
            PSTypeName           = "UpdateMessageTimeStamps"
            '#'                  = $messageCounter
            Subject              = $subject
            ItemClass            = $currentMessage.ItemClass
            CreateDate           = $currentMessage.DateTimeCreated
            TimeReceived         = $currentMessage.DateTimeReceived
            TimeSent             = $currentMessage.DateTimeSent
            LastModificationTime = $currentMessage.LastModifiedTime
        } 
        $null = $messageList.Add($message)
        $script:messageCounter ++
    }
    $messageList | Format-Table
}
