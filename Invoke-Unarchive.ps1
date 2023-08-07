<#
    .SYNOPSIS
    Invoke-Unarchive

    Michel de Rooij
    michel@eightwone.com

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
    ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
    WITH THE USER.

    Version 1.11, August 7th, 2023

    .DESCRIPTION
    This script will process personal archives and reingest contents to their related primary mailbox.
    This can be useful when retention policies archived contents unintentionally, or when organizations want 
    to start adopting a large mailbox policy, abandoning archives archives. Note that all contents is
    moved, including recoverable items, except for data that is part of any hold.

    The script moves contents in the most optimal way:
    1) Folders present in archive but not in primary mailbox, are moved in one operation (Folder.Move)
    2) Folders present in archive and primary are merged
       a) Items in that folder are moved batch-wise (Item.Move)
       b) Subfolders are processed in the same way as 1 and 2 (and repeated when necessary).
       c) Folder should be empty after a) and b), and if so will be removed.
  
    .LINK
    http://eightwone.com

    .NOTES
    - Requires Exchange Server 2013 SP1 (or later) or Exchange Online
    - Microsoft Exchange Web Services (EWS) Managed API 2.21 or up is required. This is available as NuGet package, e.g.
      Install-Package exchange.webservices.managed.api
      See https://eightwone.com/2020/10/05/ews-webservices-managed-api on how to add NuGet as package provider.
    - OAuth requires MSAL library (Microsoft.Identity.Client.dll). This is also available as a NuGet package, e.g.
      Install-Package Microsoft.Identity.Client -ProviderName NuGet
    - Search order for DLL's is script Folder then installed packages.
    - Script has not been tested with mixed locality, e.g. primary mailboxes on-premises with Exchange Online archives.

    Revision History
    --------------------------------------------------------------------------------
    1.0     Initial release
    1.01    Fixed loading of module when using installed NuGet packages
    1.02    Changed check for proper loading of Microsoft.Identity.Client module
    1.03    Added NoSCP switch
            Setting TimeZone when connecting, required for EXO
            Added ExchangeSchema parameter
    1.04    Fixed unarchiving batches instead of whole set of items per folder
            Fixed detection of throttling and honoring backoff period
    1.05    Added progress bar for significant backoff/wait delays
    1.06    Fixed reporting of EWS error status
    1.07    Fixed logic after throttling to reset generic delay
    1.08    Further tuned calculated delays
    1.09    Removed non-functional ExchangeSchema input options
    1.10    Fixed ServerBusy reporting
    1.11    Changed OAuth to use dummy creds to prevent 'Credentials are required to make a service request' issue

    .PARAMETER Identity
    Identity of the Mailbox. Can be CN/SAMAccountName (Exchange on-premises) or e-mail (Exchange on-prem & Exchange Online)

    .PARAMETER Server
    Exchange Client Access Server to use for Exchange Web Services. When ommited, script will use Autodiscover
    to discover the endpoint to use.

    .PARAMETER NoSCP
    Will instruct to skip SCP lookups in Active Directory when using Autodiscover.

    .PARAMETER ExchangeSchema 
    Specify Exchange schema to use when connecting to Exchange server or Exchange Online.
    Options are Exchange2013_SP1, Exchange2015 or Exchange2016. Default is Exchange2013_SP1, except 
    when you specified the server parameter as 'outlook.office365.com', in which case it will be 
    set to Exchange2016 for Exchange Online compatibility reasons.

    .PARAMETER IncludeRecoverableItems
    Instructs script to include moving items back from the deletions in RecoverableItems.

    .PARAMETER Impersonation
    When specified, uses impersonation when accessing the mailbox, otherwise account specified with Credentials is
    used. When using OAuth authentication with a registered app, you don't need to specify Impersonation.
    For details on how to configure impersonation access using RBAC, see this article:
    https://eightwone.com/2014/08/13/application-impersonation-to-be-or-pretend-to-be/

    .PARAMETER Force
    Force moving of items without prompting.

    .PARAMETER NoProgressBar
    Use this switch to prevent displaying a progress bar as folders and items are being processed.

    .PARAMETER TrustAll
    Specifies if all certificates should be accepted, including self-signed certificates.

    .PARAMETER TenantId
    Specifies the identity of the Tenant.

    .PARAMETER ClientId
    Specifies the identity of the application configured in Azure Active Directory.

    .PARAMETER Credentials
    Specify credentials to use with Basic Authentication. Credentials can be set using $Credentials= Get-Credential
    This parameter is mutually exclusive with CertificateFile, CertificateThumbprint and Secret. 

    .PARAMETER CertificateThumbprint
    Specify the thumbprint of the certificate to use with OAuth authentication. The certificate needs
    to reside in the personal store. When using OAuth, providing TenantId and ClientId is mandatory.
    This parameter is mutually exclusive with CertificateFile, Credentials and Secret. 

    .PARAMETER CertificateFile
    Specify the .pfx file containing the certificate to use with OAuth authentication. When a password is required,
    you will be prompted or you can provide it using CertificatePassword.
    When using OAuth, providing TenantId and ClientId is mandatory. 
    This parameter is mutually exclusive with CertificateFile, Credentials and Secret. 

    .PARAMETER CertificatePassword
    Sets the password to use with the specified .pfx file. The provided password needs to be a secure string, 
    eg. -CertificatePassword (ConvertToSecureString -String 'P@ssword' -Force -AsPlainText)

    .PARAMETER Secret
    Specifies the client secret to use with OAuth authentication. The secret needs to be provided as a secure string.
    When using OAuth, providing TenantId and ClientId is mandatory. 
    This parameter is mutually exclusive with CertificateFile, Credentials and CertificateThumbprint. 

    .EXAMPLE
    .\Invoke-Unarchive.ps1 -Identity john@contoso.com -Server outlook.office365.com -Impersonation -Secret $MySecret -TenantId $TenantId -ClientId $ClientId -Verbose

    Invokes unarchiving of contents from personal archive to primary mailbox for john@contoso.com. Authentication is OAuth using provided Tenant ID, Client ID and secret.

    #>
[cmdletbinding(
    SupportsShouldProcess= $true,
    ConfirmImpact= 'High'
)]
param(
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Position= 0, Mandatory= $true, ValueFromPipelineByPropertyName= $true, ParameterSetName= 'BasicAuth')] 
    [alias('Mailbox')]
    [string[]]$Identity,
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [string]$Server,
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [switch]$NoSCP,
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [ValidateSet( 'Exchange2013_SP1', 'Exchange2015', 'Exchange2016' )]
    [string]$ExchangeSchema='Exchange2013_SP1',
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [switch]$Impersonation,
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [switch]$IncludeRecoverableItems,
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [switch]$Force,
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [switch]$NoProgressBar,
    [parameter( Mandatory= $true, ParameterSetName= 'BasicAuth')] 
    [System.Management.Automation.PsCredential]$Credentials,
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecret')] 
    [System.Security.SecureString]$Secret,
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumb')] 
    [String]$CertificateThumbprint,
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFile')] 
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf})]
    [String]$CertificateFile,
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [System.Security.SecureString]$CertificatePassword,
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecret')] 
    [string]$TenantId,
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $true, ParameterSetName= 'OAuthCertSecret')] 
    [string]$ClientId,
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')] 
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')] 
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')] 
    [switch]$TrustAll
)
#Requires -Version 3.0

begin {

    # Process folders these batches
    $script:FolderBatchSize= @{ Min=10; Max=100; Current=100}
    # Process items in these page sizes
    $script:ItemBatchSize= @{ Min=10; Max=100; Current=25}
    # Sleep timers (ms) to backoff EWS operations
    $script:SleepTimer= @{ Min=100; Max=300000; Current= 250}
    # TuningFactors
    $script:Factors= @{ Dec=0.5; Inc=1.5}
    # Variable holding any detected backoff period
    $script:BackOffMilliseconds= 0

    # Error codes
    $ERR_DLLNOTFOUND= 1000
    $ERR_DLLLOADING= 1001
    $ERR_MAILBOXNOTFOUND= 1002
    $ERR_AUTODISCOVERFAILED= 1003
    $ERR_CANTACCESSMAILBOXSTORE= 1004
    $ERR_PROCESSINGMAILBOX= 1005
    $ERR_INVALIDCREDENTIALS= 1007
    $ERR_PROBLEMIMPORTINGCERT= 1008
    $ERR_CERTNOTFOUND= 1009

### HELPER FUNCTIONS ###

    Function Import-ModuleDLL {
        param(
            [string]$Name,
            [string]$FileName,
            [string]$Package
        )

        $AbsoluteFileName= Join-Path -Path $PSScriptRoot -ChildPath $FileName
        If ( Test-Path $AbsoluteFileName) {
            # OK
        }
        Else {
            If( $Package) {
                If( Get-Command -Name Get-Package -ErrorAction SilentlyContinue) {
                    If( Get-Package -Name $Package -ErrorAction SilentlyContinue) {
                        $AbsoluteFileName= (Get-ChildItem -ErrorAction SilentlyContinue -Path (Split-Path -Parent (get-Package -Name $Package | Sort-Object -Property Version -Descending | Select-Object -First 1).Source) -Filter $FileName -Recurse).FullName
                    }
                }
            }
        }

        If( $absoluteFileName) {
            $ModLoaded= Get-Module -Name $Name -ErrorAction SilentlyContinue
            If( $ModLoaded) {
                Write-Verbose ('Module {0} v{1} already loaded' -f $ModLoaded.Name, $ModLoaded.Version)
            }
            Else {
                Write-Verbose ('Loading module {0}' -f $absoluteFileName)
                try {
                    Import-Module -Name $absoluteFileName -Global -Force
                    Start-Sleep 1
                }
                catch {
                    Write-Error ('Problem loading module {0}: {1}' -f $Name, $error[0])
                    Exit $ERR_DLLLOADING
                }
                $ModLoaded= Get-Module -Name $Name -ErrorAction SilentlyContinue
                If( $ModLoaded) {
                    Write-Verbose ('Module {0} v{1} loaded' -f $ModLoaded.Name, $ModLoaded.Version)
                }
                If(!( Get-Module -Name $Name -ErrorAction SilentlyContinue)) {
                    Write-Error ('Problem loading module {0}: {1}' -f $Name, $_.Exception.Message)
                    Exit $ERR_DLLLOADING
                }
            }
        }
        Else {
            Write-Verbose ('Required module {0} could not be located' -f $FileName)
            Exit $ERR_DLLNOTFOUND
        }
    }

    Function Set-SSLVerification {
        param(
            [switch]$Enable,
            [switch]$Disable
        )

        Add-Type -TypeDefinition  @"
            using System.Net.Security;
            using System.Security.Cryptography.X509Certificates;
            public static class TrustEverything
            {
                private static bool ValidationCallback(object sender, X509Certificate certificate, X509Chain chain,
                    SslPolicyErrors sslPolicyErrors) { return true; }
                public static void SetCallback() { System.Net.ServicePointManager.ServerCertificateValidationCallback= ValidationCallback; }
                public static void UnsetCallback() { System.Net.ServicePointManager.ServerCertificateValidationCallback= null; }
        }
"@
        If($Enable) {
            Write-Verbose ('Enabling SSL certificate verification')
            [TrustEverything]::UnsetCallback()
        }
        Else {
            Write-Verbose ('Disabling SSL certificate verification')
            [TrustEverything]::SetCallback()
        }
    }

    Function iif( $eval, $tv= '', $fv= '') {
        If ( $eval) { return $tv } else { return $fv}
    }

    Function Get-EmailAddress {
        param(
            [string]$Identity
        )
        $address= [regex]::Match([string]$Identity, ".*@.*\..*", "IgnoreCase")
        if ( $address.Success ) {
            return $address.value.ToString()
        }
        Else {
            # Use local AD to look up e-mail address using $Identity as SamAccountName
            $ADSearch= New-Object DirectoryServices.DirectorySearcher( [ADSI]"")
            $ADSearch.Filter= "(|(cn=$Identity)(samAccountName=$Identity)(mail=$Identity))"
            $Result= $ADSearch.FindOne()
            If ( $Result) {
                $objUser= $Result.getDirectoryEntry()
                return $objUser.mail.toString()
            }
            else {
                return $null
            }
        }
    }

    Function Optimize-OperationalParameters {
        param(
            [bool]$previousResultSuccess= $false
        )
        if ( $previousResultSuccess) {
            If ( $script:SleepTimer['Current'] -gt $script:SleepTimer['Min']) {
                $script:SleepTimer['Current']= [int]([math]::Max( [int]($script:SleepTimer['Current'] * $script:Factors['Dec']), $script:SleepTimer['Min']))
                $script:FolderBatchSize['Current']= [int]([math]::Min( ($script:FolderBatchSize['Current'] * $script:Factors['Inc']), $script:FolderBatchSize['Max']))
                $script:ItemBatchSize['Current']= [int]([math]::Min( ($script:ItemBatchSize['Current'] * $script:Factors['Inc']), $script:ItemBatchSize['Max']))
            }
            $waitMs= $script:SleepTimer['Current']
            If( $waitMs -gt 5000) {
                Write-Verbose ('Waiting for {0:N0}s to be nice to the back-end' -f ($waitMs/1000))
            }
        }
        Else {
            # Previous operation failed, see if we're throttled or need to use calculated delay
            $waitMs= [int]($script:BackOffMilliseconds)
            If( $waitMs -eq 0) {
                If( $script:SleepTimer['Current'] -lt $script:SleepTimer['Max']) {
                    $script:SleepTimer['Current']= [int]([math]::Min( ($script:SleepTimer['Current'] * $script:Factors['Inc']), $script:SleepTimer['Max']))
                    $script:FolderBatchSize['Current']= [int]([math]::Max( [int]($script:FolderBatchSize['Current'] * $script:Factors['Dec']), $script:FolderBatchSize['Min']))
                    $script:ItemBatchSize['Current']= [int]([math]::Max( [int]($script:ItemBatchSize['Current'] * $script:Factors['Dec']), $script:ItemBatchSize['Min']))
                }
                $waitMs= $script:SleepTimer['Current']
                Write-Warning ('Previous EWS operation failed, waiting for {0:N0}s' -f ($waitMs/1000))
            }
            Else {
                Write-Warning ('Throttling detected; server requested to backoff for {0:N0}s' -f ($waitMs/1000))
            }
        }
        If( $waitMS -ge 5000 -and !( $NoProgressBar)) {
            # When waiting for >10s, show a progress bar
            $WaitUnit= [uint]($waitMS/10)
            1..10 | ForEach-Object {
                Write-Progress -Id 4 -Activity 'Waiting' -Status 'Waiting for back-end' -PercentComplete ($_ * 10)
                Start-Sleep -Milliseconds $WaitUnit
            }
            Write-Progress -Id 4 -Activity 'Waiting' -Status 'Done' -Completed
        }
        Else {
            Start-Sleep -Milliseconds $waitMs
        }
        $script:BackOffMilliseconds= 0
    }

    Function myEWSFind-Folders {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [Microsoft.Exchange.WebServices.Data.FolderId]$FolderId,
            [Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection]$FolderSearchCollection,
            [Microsoft.Exchange.WebServices.Data.FolderView]$FolderView
        )
        $OpSuccess= $false
        $CritErr= $false
        Do {
            Try {
                $res= $EwsService.FindFolders( $FolderId, $FolderSearchCollection, $FolderView)
                $OpSuccess= $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess= $false
                Write-Warning ('EWS operation failed ({0}), will retry later' -f $_.Exception.ErrorCode)
                $script:BackOffMilliseconds= $_.Exception.BackOffMilliseconds
            }
            catch {
                $OpSuccess= $false
                $critErr= $true
                Write-Warning ('Error performing operation FindFolders with Search options in {0}. Error: {1}' -f $FolderId.FolderName, $_.Exception.InnerException.Message)
            }
            finally {
                If ( !$critErr) { Optimize-OperationalParameters $OpSuccess }
            }
        } while ( !$OpSuccess -and !$critErr)
        Write-Output -NoEnumerate $res
    }

    Function myEWSMove-Folder {
        param(
            [Microsoft.Exchange.WebServices.Data.Folder]$SourceFolder,
            [Microsoft.Exchange.WebServices.Data.Folder]$TargetFolder
        )
        $OpSuccess= $false
        $CritErr= $false
        Do {
            Try {
                $null= $SourceFolder.Move( $TargetFolder.Id)
                $OpSuccess= $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess= $false
                Write-Warning ('EWS operation failed ({0}), will retry later' -f $_.Exception.ErrorCode)
                $script:BackOffMilliseconds= $_.Exception.BackOffMilliseconds
            }
            catch [Microsoft.Exchange.WebServices.Data.ServiceRequestException] {
                $OpSuccess= $false
                Write-Warning ('EWS operation ({0}) failed, will retry later' -f $_.Exception.InnerException.Status)
            }
            catch {
                $OpSuccess= $false
                $critErr= $true
                Write-Warning ('Error performing operation MoveFolder on {0}. Error: {1}' -f $SourceFolder.DisplayName, $_.Exception.InnerException.Message)
            }
            finally {
                If ( !$critErr) { Optimize-OperationalParameters $OpSuccess }
            }
        } while ( !$OpSuccess -and !$critErr)
    }

    Function myEWSDelete-Folder {
        param(
            [Microsoft.Exchange.WebServices.Data.Folder]$Folder,
            [Microsoft.Exchange.WebServices.Data.DeleteMode]$DeleteMode
        )
        $OpSuccess= $false
        $CritErr= $false
        $res= $false
        Do {
            Try {
                $Folder.Delete( $DeleteMode)
                $res= $true
                $OpSuccess= $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess= $false
                Write-Warning ('EWS operation failed ({0}), will retry later' -f $_.Exception.ErrorCode)
                $script:BackOffMilliseconds= $_.Exception.BackOffMilliseconds
            }
            catch {
                If( $_.Exception.InnerException.ErrorCode -eq 'ErrorDeleteDistinguishedFolder') {
                    $OpSuccess= $true
                    Write-Host ('{0} is a non-removable Distinguished Folder, skipping deletion' -f $Folder.DisplayName)
                }
                Else {
                    $OpSuccess= $false
                    $critErr= $true
                    Write-Warning ('Error deleting folder {0}: {1}' -f $Folder.DisplayName, $_.Exception.InnerException.Message)
                }
           }
           finally {
                If ( !$critErr) { Optimize-OperationalParameters $OpSuccess }
           }
        } while ( !$OpSuccess -and !$critErr)

        Write-Output -NoEnumerate $res
    }

    Function myEWSMove-Items {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            $ItemIds,
            [Microsoft.Exchange.WebServices.Data.Folder]$Folder
        )
        $OpSuccess= $false
        $CritErr= $false
        Do {
            Try {
                $res= $EwsService.MoveItems( $ItemIds, $Folder.Id)
                $OpSuccess= $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess= $false
                Write-Warning ('EWS operation failed ({0}), will retry later' -f $_.Exception.ErrorCode)
                $script:BackOffMilliseconds= $_.Exception.BackOffMilliseconds
            }
            catch [Microsoft.Exchange.WebServices.Data.ServiceRequestException] {
                $OpSuccess= $false
                Write-Warning ('EWS operation failed ({0}), will retry later' -f $_.Exception.InnerException.Status)
            }
            catch {
                $OpSuccess= $false
                $critErr= $true
                Write-Warning ('Error performing operation MoveItems on {0}. Error: {1}' -f $Folder.DisplayName, $_.Exception.InnerException.Message)
            }
            finally {
                If ( !$critErr) { Optimize-OperationalParameters $OpSuccess }
            }
        } while ( !$OpSuccess -and !$critErr)

        Write-Output -NoEnumerate $res
    }   

    Function myEWSFind-FoldersNoSearch {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [Microsoft.Exchange.WebServices.Data.Folder]$Folder,
            [Microsoft.Exchange.WebServices.Data.FolderView]$FolderView
        )
        $OpSuccess= $false
        $CritErr= $false
        Do {
            Try {
                $res= $EwsService.FindFolders( $Folder.Id, $FolderView)
                $OpSuccess= $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess= $false
                Write-Warning ('EWS operation failed ({0}), will retry later' -f $_.Exception.ErrorCode)
            }
            catch {
                $OpSuccess= $false
                $critErr= $true
                Write-Warning ('Error performing operation FindFolders without Search options in {0}. Error: {1}' -f $Folder.Id.FolderName, $_.Exception.InnerException.Message)
            }
            finally {
                If ( !$critErr) { Optimize-OperationalParameters $OpSuccess }
            }
        } while ( !$OpSuccess -and !$critErr)
        Write-Output -NoEnumerate $res
    }

    Function myEWSFind-ItemsNoSearch {
        param(
            [Microsoft.Exchange.WebServices.Data.Folder]$Folder,
            [Microsoft.Exchange.WebServices.Data.ItemView]$ItemView
        )
        $OpSuccess= $false
        $CritErr= $false
        Do {
            Try {
                $res= $Folder.FindItems( $ItemView)
                $OpSuccess= $true
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess= $false
                Write-Warning ('EWS operation failed ({0}), will retry later' -f $_.Exception.ErrorCode)
                $script:BackOffMilliseconds= $_.Exception.BackOffMilliseconds
            }
            catch {
                $OpSuccess= $false
                $critErr= $true
                Write-Warning ('Error performing operation FindItems without Search options in {0}: {1}' -f $Folder.DisplayName, $_.Exception.InnerException.Message)
            }
            finally {
                If ( !$critErr) { Optimize-OperationalParameters $OpSuccess }
            }
        } while ( !$OpSuccess -and !$critErr)
        Write-Output -NoEnumerate $res
    }

    Function myEWSBind-WellKnownFolder {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [string]$WellKnownFolderName,
            [string]$emailAddress,
            [switch]$ShowVersion
        )
        $OpSuccess= $false
        $critErr= $false
        Do {
            Try {
                $explicitFolder= New-Object -TypeName Microsoft.Exchange.WebServices.Data.FolderId( [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$WellKnownFolderName, $emailAddress)  
                $res= [Microsoft.Exchange.WebServices.Data.Folder]::Bind( $EwsService, $explicitFolder)
                $OpSuccess= $true
                if( $ShowVersion) {
                    # Show Exchange build when connecting to a primary/archive/pf mailbox
                    Write-Verbose ('Detected Exchange Server version {0}.{1}.{2}.{3} ({4}, requested schema {5})' -f $EwsService.ServerInfo.MajorVersion, $EwsService.ServerInfo.MinorVersion, $EwsService.ServerInfo.MajorBuildNumber, $EwsService.ServerInfo.MinorBuildNumber, $EwsService.ServerInfo.VersionString, $ExchangeSchema)
                }
            }
            catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
                $OpSuccess= $false
                Write-Warning ('EWS operation failed ({0}), will retry later' -f $_.Exception.ErrorCode)
                $script:BackOffMilliseconds= $_.Exception.BackOffMilliseconds
            }
            catch {
                $OpSuccess= $false
                $critErr= $true
                Write-Warning ('Cannot bind to {0}: {1}' -f $WellKnownFolderName, $Error[0])
            }
            finally {
                If ( !$critErr) { Optimize-OperationalParameters $OpSuccess }
            }
        } while ( !$OpSuccess -and !$critErr)

        $res
    }

    Function Get-SubFolders {
        param(
            [Microsoft.Exchange.WebServices.Data.Folder]$Folder,
            $CurrentPath
        )
        $FoldersToProcess= [System.Collections.ArrayList]@()
        $FolderView= New-Object Microsoft.Exchange.WebServices.Data.FolderView( $script:FolderBatchSize['Current'])
        $FolderView.Traversal= [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow
        $FolderView.PropertySet= New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
            [Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties )
        Do {
            $FolderSearchResults= myEWSFind-FoldersNoSearch -EwsService $EwsService -Folder $Folder -FolderView $FolderView
            ForEach ( $FolderItem in $FolderSearchResults) {
                $FolderPath= Join-Path -Path $CurrentPath -ChildPath $FolderItem.DisplayName 
                $Obj= New-Object -TypeName PSObject -Property @{
                    Path    = $FolderPath
                    Folder  = $FolderItem
                }
                Write-Debug ('Located {0}' -f $FolderPath)
                $null= $FoldersToProcess.Add( $Obj)
            }
            $FolderView.Offset += $FolderSearchResults.Folders.Count
        } While ($FolderSearchResults.MoreAvailable)
        Write-Output -NoEnumerate $FoldersToProcess
    }

    Function Move-MailboxContents {
        [CmdletBinding(SupportsShouldProcess=$true)]
        param(
            [string]$Identity,
            [Microsoft.Exchange.WebServices.Data.Folder]$Target,
            [Microsoft.Exchange.WebServices.Data.Folder]$Source,
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            $emailAddress,
            [string]$CurrentPath
        )

        $ProcessingOK= $True
        $FoldersFound= 0
        $FoldersProcessed= 0
        $ItemsFound= 0
        $ItemsProcessed= 0

        # Do make 'root' show up as \
        $prettyCurrentPath= iif -eval ([string]::IsNullOrEmpty( $CurrentPath)) -tv '\' -fv $CurrentPath

        Write-Verbose ('Collecting folders to unarchive in {0}' -f $prettyCurrentPath)

        # Build list of folders to process
        $FoldersToProcess= Get-SubFolders -EwsService $EwsService -Folder $Source -CurrentPath $prettyCurrentPath

        # Build list of folders in target mailbox, so we can determine what we can just move and what we need to 'merge'
        $ExistingFolders= Get-SubFolders -EwsService $EwsService -Folder $Target -CurrentPath $prettyCurrentPath

        $ItemView= New-Object Microsoft.Exchange.WebServices.Data.ItemView( $script:ItemBatchSize['Current'])
        $ItemView.Traversal= [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Shallow
        $ItemView.PropertySet= New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
            [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
        $ItemType= ("System.Collections.Generic.List" + '`' + "1") -as 'Type'
        $ItemType= $ItemType.MakeGenericType([Microsoft.Exchange.WebServices.Data.ItemId] -as 'Type')

        # First, collect all Item Ids
        $ItemsToProcess= [Activator]::CreateInstance( $ItemType)
        $ItemResults= myEWSFind-ItemsNoSearch -EwsService $EwsService -Folder $Source -ItemView $ItemView
        If( $ItemResults.Items.Count -gt 0) {

            Write-Verbose ('Retrieving items to unarchive from {0} ..' -f $prettyCurrentPath)
            While( $ItemResults.Items.Count -gt 0) {
                $ItemsFound+= $ItemResults.Items.Count
                If (!$NoProgressBar) {
                    Write-Progress -Id 3 -Activity ('Unarchiving items in folder {0}' -f $prettyCurrentPath) -Status ('Discovered {0} items' -f $ItemsFound)
                }
                ForEach ( $Item in $ItemResults) {
                    $ItemsToProcess.Add( $Item.Id)
                }
                $ItemView.Offset += $ItemResults.Items.Count
                $ItemResults= myEWSFind-ItemsNoSearch -EwsService $EwsService -Folder $Source -ItemView $ItemView
            }

            If( $ItemsFound -gt 0) {

                Write-Verbose ('Discovered {0} items in {1}' -f $ItemsFound, $prettyCurrentPath)

                # Chop set of items in chunks to process
                $ItemBatch= [Activator]::CreateInstance( $ItemType)
                $ItemsProcessed=0
                ForEach( $Item in $ItemsToProcess) {
                    $ItemBatch.Add( $Item)
                    $ItemsProcessed++
        
                    # When cut-off for items or last item in batch reached, process the subset
                    If( $ItemBatch.Count -ge $script:ItemBatchSize['Current'] -or $ItemsProcessed -eq $ItemsFound) {
                        If (!$NoProgressBar) {
                                Write-Progress -Id 3 -Activity ('Processing folder {0}' -f $prettyCurrentPath) -Status ('Unarchived {0} items of {1}' -f $ItemsProcessed, $ItemsFound) -PercentComplete ( $ItemsProcessed / $ItemsFound * 100)
                        }
                        If ( $Force -or $PSCmdlet.ShouldProcess( ('Unarchiving {1} item(s) from folder {0}' -f $prettyCurrentPath, $ItemBatch.Count))) {
                            try {
                                Write-Verbose ('Unarchiving {1}/{2} item(s) from {0}' -f $prettyCurrentPath, $ItemsProcessed, $ItemsFound)
                                $null= myEWSMove-Items -EwsService $EwsService -ItemIds $ItemBatch -Folder $Target
                            }
                            catch {
                                Write-Error ('Problem unarchiving items from folder {0}: {1}' -f $prettyCurrentPath, $_.Exception.InnerException.Message)
                            }
                        }
                        $ItemBatch= [Activator]::CreateInstance( $ItemType)
                    }
                } 
            }
        }

        If (!$NoProgressBar) {
            Write-Progress -Id 3 -Activity ('Processing folder {0}' -f $prettyCurrentPath) -Status 'Finished unarchiving items.' -Completed
        }

        $FoldersFound= $FoldersToProcess.Count
        Write-Verbose ('Located {0} folders in {1}' -f $FoldersFound, $prettyCurrentPath)

        ForEach ( $SubFolder in $FoldersToProcess) {

            $FoldersProcessed++

            If (!$NoProgressBar) {
                Write-Progress -Id 1 -Activity ('Processing {0}' -f $Identity) -Status ('Unarchiving folder {0} of {1}' -f $FoldersProcessed, $FoldersFound) -PercentComplete ( $FoldersProcessed / $FoldersFound * 100)
            }

            $MatchingTarget= $ExistingFolders | Where-Object { $_.Path -eq $SubFolder.Path }
            If( $MatchingTarget) {

                # First, see if we need to unarchive contents of this folder first
                If($SubFolder.Folder.TotalCount -gt 0 -or $SubFolder.Folder.ChildFolderCount -gt 0) {
                    Write-Host ('Folder {0} exists, merging contents to unarchive {1} item(s) and {2} folder(s)' -f $SubFolder.Path, $SubFolder.Folder.TotalCount, $SubFolder.Folder.ChildFolderCount)
                    Move-MailboxContents -Identity $Identity -Source $SubFolder.Folder -Target $MatchingTarget.Folder -EwsService $EwsService -emailAddress $emailAddress -CurrentPath $SubFolder.Path

                }

                # Refresh folder to get proper counts
                $SubFolder.Folder.Load()

                # Second, see if we need to/can remove this folder
                If($SubFolder.Folder.TotalCount -eq 0 -and $SubFolder.Folder.ChildFolderCount -eq 0) {
                    Write-Host ('Folder {0} in personal archive is empty' -f $SubFolder.Path)
                    If( myEWSDelete-Folder -Folder $SubFolder.Folder -DeleteMode 'HardDelete') {
                        Write-Host ('Folder {0} in personal archive has been removed' -f $SubFolder.Path)
                    }
                }
                Else {
                    Write-Warning ('Folder {0} in personal archive is not empty, not removing' -f $SubFolder.Path)
                }
            }
            Else {

                If ( $Force -or $PSCmdlet.ShouldProcess( ('Unarchiving folder {0}' -f $SubFolder.Path))) {
                    try {
                        Write-Host ('Unarchiving folder {0}' -f $SubFolder.Path)
                        $null= myEWSMove-Folder -SourceFolder $SubFolder.Folder -TargetFolder $Target
                    }
                    catch {
                        Write-Error ('Problem unarchiving folder {0}: {1}' -f $SubFolder.Path, $_.Exception.Message)
                    }
                }
            }
            If (!$NoProgressBar) {
                Write-Progress -Id 2 -Activity ('Processing folder {0}' -f $SubFolder.Path) -Status 'Finished unarchiving folder.' -Completed
            }

        } # ForEach SubFolder

        If (!$NoProgressBar) {
            Write-Progress -Id 1 -Activity ('Processing {0}' -f $Identity) -Status 'Finished unarchiving.' -Completed
        }

        Return $ProcessingOK
    }

    ### MAIN ROUTINE ###

    Import-ModuleDLL -Name 'Microsoft.Exchange.WebServices' -FileName 'Microsoft.Exchange.WebServices.dll' -Package 'Exchange.WebServices.Managed.Api'
    Import-ModuleDLL -Name 'Microsoft.Identity.Client' -FileName 'Microsoft.Identity.Client.dll' -Package 'Microsoft.Identity.Client'

    $TZ= [System.TimeZoneInfo]::Local
    # Override ExchangeSchema when connecting to EXO
    If( $Server -eq 'outlook.office365.com') {
        $ExchangeSchema= 'Exchange2016'
    } 
    Try {
        $EwsService= [Microsoft.Exchange.WebServices.Data.ExchangeService]::new( [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$ExchangeSchema, [System.TimeZoneInfo]::FindSystemTimeZoneById( $TZ.Id))
    }
    Catch {
        Throw( 'Problem initializing Exchange Web Services using schema {0} and TimeZone {1}' -f $ExchangeSchema, $TZ.Id)
    }

    If( $Credentials) {
        try {
            Write-Verbose ('Using credentials {0}' -f $Credentials.UserName)
            $EwsService.Credentials= [System.Net.NetworkCredential]::new( $Credentials.UserName, [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $Credentials.Password )))
        }
        catch {
            Write-Error ('Invalid credentials provided: {0}' -f $_.Exception.Message)
            Exit $ERR_INVALIDCREDENTIALS
        }
    }
    Else {
        # Use OAuth (and impersonation/X-AnchorMailbox always set)
        $Impersonation= $true

        # Dummy creds to prevent "Credentials are required to make a service request" issue
        $EwsService.Credentials= [System.Net.NetworkCredential]::new( '', ( ConvertTo-SecureString -String 'dummy' -AsPlainText -Force))

        If( $CertificateThumbprint -or $CertificateFile) {
            If( $CertificateFile) {
                
                # Use certificate from file using absolute path to authenticate
                $CertificateFile= (Resolve-Path -Path $CertificateFile).Path
                
                Try {
                    If( $CertificatePassword) {
                        $X509Certificate2= [System.Security.Cryptography.X509Certificates.X509Certificate2]::new( $CertificateFile, [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $CertificatePassword)))
                    }
                    Else {
                        $X509Certificate2= [System.Security.Cryptography.X509Certificates.X509Certificate2]::new( $CertificateFile)
                    }
                }
                Catch {
                    Write-Error ('Problem importing PFX: {0}' -f $_.Exception.Message)
                    Exit $ERR_PROBLEMIMPORTINGCERT
                }
            }
            Else {
                # Use provided certificateThumbprint to retrieve certificate from My store, and authenticate with that
                $CertStore= [System.Security.Cryptography.X509Certificates.X509Store]::new( [Security.Cryptography.X509Certificates.StoreName]::My, [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
                $CertStore.Open( [System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly )
                $X509Certificate2= $CertStore.Certificates.Find( [System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint, $CertificateThumbprint, $False) | Select-Object -First 1
                If(!( $X509Certificate2)) {
                    Write-Error ('Problem locating certificate in My store: {0}' -f $error[0])
                    Exit $ERR_CERTNOTFOUND
                }
            }
            Write-Verbose ('Will use certificate {0}, issued by {1} and expiring {2}' -f $X509Certificate2.Thumbprint, $X509Certificate2.Issuer, $X509Certificate2.NotAfter)
            $App= [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create( $ClientId).WithCertificate( $X509Certificate2).withTenantId( $TenantId).Build()
               
        }
        Else {
            # Use provided secret to authenticate
            Write-Verbose ('Will use provided secret to authenticate')
            $PlainSecret= [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR( $Secret))
            $App= [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create( $ClientId).WithClientSecret( $PlainSecret).withTenantId( $TenantId).Build()
        }
        $Scopes= New-Object System.Collections.Generic.List[string]
        $Scopes.Add( 'https://outlook.office365.com/.default')
        Try {
            $Response=$App.AcquireTokenForClient( $Scopes).executeAsync()
            $Token= $Response.Result
            $EwsService.Credentials= [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$Token.AccessToken
            Write-Verbose ('Authentication token acquired')
        }
        Catch {
            Write-Error ('Problem acquiring token: {0}' -f $error[0])
            Exit $ERR_INVALIDCREDENTIALS
        }
    }

    $EwsService.EnableScpLookup= if( $NoSCP) { $false } else { $true }
    $EwsService.Timeout= $script:SleepTimer['Max']

    If( $TrustAll) {
        Set-SSLVerification -Disable
    }
}

Process {

    ForEach ( $CurrentIdentity in $Identity) {

        $EmailAddress= get-EmailAddress -Identity $CurrentIdentity
        If ( !$EmailAddress) {
            Write-Error ('Specified mailbox {0} not found' -f $EmailAddress)
            Exit $ERR_MAILBOXNOTFOUND
        }

        Write-Host ('Processing mailbox {0} ({1})' -f $EmailAddress, $CurrentIdentity)

        If( $Impersonation) {
            Write-Verbose ('Using {0} for impersonation' -f $EmailAddress)
            $EwsService.ImpersonatedUserId= [Microsoft.Exchange.WebServices.Data.ImpersonatedUserId]::new( [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $EmailAddress)
            $EwsService.HttpHeaders.Clear()
            $EwsService.HttpHeaders.Add( 'X-AnchorMailbox', $EmailAddress)
        }
            
        If ($Server) {
            $EwsUrl= 'https://{0}/EWS/Exchange.asmx' -f $Server
            Write-Verbose ('Using Exchange Web Services URL {0}' -f $EwsUrl)
            $EwsService.Url= $EwsUrl
        }
        Else {
            Write-Verbose ('Looking up EWS URL using Autodiscover for {0}' -f $EmailAddress)
            try {
                # Set script to terminate on all errors (autodiscover failure isn't) to make try/catch work
                $ErrorActionPreference= 'Stop'
                $EwsService.autodiscoverUrl( $EmailAddress, {$true})
            }
            catch {
                Write-Error ('Autodiscover failed: {0}' -f $_.Exception.Message)
                Exit $ERR_AUTODISCOVERFAILED
            }
            $ErrorActionPreference= 'Continue'
            Write-Verbose 'Using EWS endpoint {0}' -f $EwsService.Url
        } 

        $PrimaryRootFolder= myEWSBind-WellKnownFolder $EwsService 'MsgFolderRoot' $EmailAddress -ShowVersion
        If ($null -ne $PrimaryRootFolder) {
            $ArchiveRootFolder= myEWSBind-WellKnownFolder $EwsService 'ArchiveMsgFolderRoot' $EmailAddress
            If ($null -ne $ArchiveRootFolder) {
                Write-Verbose ('Unarchiving personal archive contents to primary mailbox of {0} ({1})' -f $EmailAddress, $CurrentIdentity)
                If (! ( Move-MailboxContents -Identity $CurrentIdentity -Target $PrimaryRootFolder -Source $ArchiveRootFolder -EwsService $EwsService -emailAddress $emailAddress -CurrentPath '')) {
                    Write-Error ('Problem unarchiving of {0} ({1})' -f $EmailAddress, $CurrentIdentity)
                    Exit $ERR_PROCESSINGMAILBOX
                }
            }
            Else {
                Write-Error ('Cannot access personal archive of {0}: {1}' -f $EmailAddress, $_.Exception.Message)
                Exit $ERR_CANTACCESSMAILBOXSTORE
            }
            }
        Else {
            Write-Error ('Cannot access primary mailbox of {0}: {1}' -f $EmailAddress, $_.Exception.Message)
            Exit $ERR_CANTACCESSMAILBOXSTORE
        }

        If($IncludeRecoverableItems.IsPresent) {
            $RecoverableItemsDeletions= myEWSBind-WellKnownFolder $EwsService 'RecoverableItemsDeletions' $EmailAddress
            If ($null -ne $RecoverableItemsDeletions) {
                $ArchiveRecoverableItemsDeletions= myEWSBind-WellKnownFolder $EwsService 'ArchiveRecoverableItemsDeletions' $EmailAddress
                If ($null -ne $ArchiveRecoverableItemsDeletions) {
                    Write-Verbose ('Unarchiving recoverable deleted items from personal archive to primary mailbox of {0} ({1})' -f $EmailAddress, $CurrentIdentity)
                    If (! ( Move-MailboxContents -Identity $CurrentIdentity -Target $RecoverableItemsDeletions -Source $ArchiveRecoverableItemsDeletions -EwsService $EwsService -emailAddress $emailAddress -CurrentPath '')) {
                        Write-Error ('Problem unarchiving recoverable items of {0} ({1})' -f $EmailAddress, $CurrentIdentity)
                        Exit $ERR_PROCESSINGRECOVERABLEITEMS
                    }
                }
                Else {
                    Write-Error ('Cannot access personal archive recoverable items of {0}: {1}' -f $EmailAddress, $_.Exception.Message)
                    Exit $ERR_CANTACCESSMAILBOXSTORE
                }
                }
            Else {
                Write-Error ('Cannot access recoverable items of {0}: {1}' -f $EmailAddress, $_.Exception.Message)
                Exit $ERR_CANTACCESSMAILBOXSTORE
            }
        }
        Write-Verbose ('Processing {0} finished' -f $EmailAddress)
    }
}
End {
    If( $TrustAll) {
        Set-SSLVerification -Enable
    }
}
