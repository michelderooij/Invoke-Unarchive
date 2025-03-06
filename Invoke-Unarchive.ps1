<#
    .SYNOPSIS
    Invoke-Unarchive

    Michel de Rooij
    michel@eightwone.com

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE
    ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS
    WITH THE USER.

    Version 1.32, March 6th, 2025

    .DESCRIPTION
    This script will process personal archives and reingest contents to their related primary mailbox.
    This can be useful when retention policies archived contents unintentionally, or when organizations want
    to start adopting a large mailbox policy, abandoning archives. Note that all contents is
    moved, including recoverable items, except for data that is part of any hold.

    The script moves contents by recreating folders if needed, and moving items back to the
    primary mailbox batch-wise. After processing, emptied regular folders in the archive are removed.

    .LINK
    http://eightwone.com

    .NOTES
    - Requires Exchange Server 2013 SP1 (or later) or Exchange Online
    - Microsoft Exchange Web Services (EWS) Managed API 2.2 or up is required. This is available as NuGet package, e.g.
      Install-Package microsoft.exchange.webservices -ProviderName NuGet
    - OAuth requires MSAL library. This is also available as a NuGet package, e.g.
      Install-Package Microsoft.Identity.Client -ProviderName NuGet
    - Search order for DLL's is script folder, then installed packages.
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
    1.12    Changed [uint] accelerator to [uint32] for PS5.1 compatibility
    1.20    Added IncludeFolder and ExcludeFolder parameters
            Switched to Microsoft.Exchange.WebServices package detection/loading
            Fixed loading of DLLs when they were installed using package manager
    1.30    Rewritten folder move to recreate folder + content move due to EWS MoveFolder issue
            Because of MoveFolder change, added empty folder cleanup
            Rewritten folder collection for better performance
            Refactored module loading
            Bumped required PowerShell version to 5
    1.31    Fixed processing folders of more than 2 levels deep
            Minor tweaks
    1.32    Removed obsolete function
            Functions now use approved verbs
            Put Garbage Collection in place
            Minor tweaks
            Updated description in synopsis

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

    .PARAMETER IncludeFolders
    Specify one or more names of folder(s) to include, e.g. 'Projects'. You can use wildcards
    around or at the end to include folders containing or starting with this string, e.g.
    'Projects*' or '*Project*'. To match folders and subfolders, add a trailing \*,
    e.g. Projects\*. This will include folders named Projects and all subfolders.
    To match from the top of the structure, prepend using '\'. Matching is case-insensitive.

    .PARAMETER ExcludeFolders
    Specify one or more folder(s) to exclude. Usage of wildcards and well-known folders identical to IncludeFolders.
    Note that ExcludeFolders criteria overrule IncludeFolders when matching folders.

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
    .\Invoke-Unarchive.ps1 -Identity john@contoso.com -Server outlook.office365.com -Secret $MySecret -TenantId $TenantId -ClientId $ClientId -Verbose

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
    [string[]]$IncludeFolders,
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertThumb')]
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertFile')]
    [parameter( Mandatory= $false, ParameterSetName= 'OAuthCertSecret')]
    [parameter( Mandatory= $false, ParameterSetName= 'BasicAuth')]
    [string[]]$ExcludeFolders,
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
#Requires -Version 5.0

begin {

    # Process folders these batches
    $script:FolderBatchSize= @{ Min=20; Max=250; Current=100}
    # Process items in these page sizes
    $script:ItemBatchSize= @{ Min=50; Max=250; Current=100}
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

        # Get type to see if already loaded
        $typeLoaded = [System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.ManifestModule -ieq $FileName }
        if ($typeLoaded) {
            $Version= If( $typeLoaded -match 'Version=[\d\.]+') { $matches[0] } else { 'Unknown' }
            Write-Verbose ('Assembly {0} ({1}) already loaded' -f $Name, $Version)
            Return
        }
        $ModLoaded= Get-Module -Name $Name -ErrorAction SilentlyContinue
        If( $ModLoaded) {
            Write-Verbose ('Module {0} v{1} already loaded' -f $ModLoaded.Name, $ModLoaded.Version)
            Return
        }

        # See if module present in local folder
        $AbsoluteFileName = Join-Path -Path $PSScriptRoot -ChildPath $FileName
        If (Test-Path $AbsoluteFileName) {
            # OK
        }
        Else {
            # See if module present in installed packages
            If ($Package) {
                If (Get-Command -Name Get-Package -ErrorAction SilentlyContinue) {
                    If (Get-Package -Name $Package -ErrorAction SilentlyContinue) {
                        $pkg = Get-Package -Name $Package | Sort-Object -Property Version -Descending | Select-Object -First 1
                        $AbsoluteFileName = Join-Path -Path $pkg.Source -ChildPath $FileName
                    }
                }
            }
        }

        If( $absoluteFileName) {
            Write-Verbose ('Loading assembly {0}' -f $absoluteFileName)
            try {
                Add-Type -Path $absoluteFileName
                $typeLoaded = [System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.ManifestModule -ieq $FileName }
                if ($typeLoaded) {
                    $Version= If( $typeLoaded -match 'Version=[\d\.]+') { $matches[0] } else { 'Unknown' }
                    Write-Verbose ('Assembly {0} ({1}) loaded' -f $Name, $Version)
                }
            }
            catch {
                Write-Verbose ('Assembly {0} loading issue:{1}' -f $Name, $Error[0].Exception.Message)
                Exit $ERR_DLLLOADING
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
            $WaitUnit= [uint32]($waitMS/10)
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

    Function Remove-myEWSFolder {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
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

    Function New-myEWSFolder {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [Microsoft.Exchange.WebServices.Data.Folder]$Folder,
            [string]$DisplayName,
            [string]$FolderClass= 'IPF.Note'
        )
        $OpSuccess= $false
        $CritErr= $false
        $res= $null
        Do {
            Try {
                $res = [Microsoft.Exchange.WebServices.Data.Folder]::new( $EwsService)
                $res.DisplayName = $DisplayName
                $res.FolderClass = $FolderClass
                $null = $res.Save( $Folder.Id)
                Write-Debug ('Created folder {0} ({1})' -f $DisplayName, $FolderClass)
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
                Write-Warning ('Error creating folder {0}: {1}' -f $DisplayName, $_.Exception.InnerException.Message)
           }
           finally {
                If ( !$critErr) { Optimize-OperationalParameters $OpSuccess }
           }
        } while ( !$OpSuccess -and !$critErr)

        Write-Output  $res
    }

    Function Move-myEWSItems {
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

    Function Find-myEWSFoldersNoSearch {
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
                Write-Warning ('Error performing operation FindFolders in {0}. Error: {1}' -f $Folder.Id.FolderName, $_.Exception.InnerException.Message)
            }
            finally {
                If ( !$critErr) { Optimize-OperationalParameters $OpSuccess }
            }
        } while ( !$OpSuccess -and !$critErr)
        Write-Output -NoEnumerate $res
    }

    Function Find-myEWSItemsNoSearch {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [Microsoft.Exchange.WebServices.Data.Folder]$Folder,
            [Microsoft.Exchange.WebServices.Data.ItemView]$ItemView
        )
        $OpSuccess= $false
        $CritErr= $false
        Do {
            Try {
                $res= $EwsService.FindItems( $Folder.Id, $ItemView)
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
                Write-Warning ('Error performing operation FindItems in {0}: {1}' -f $Folder.DisplayName, $_.Exception.InnerException.Message)
            }
            finally {
                If ( !$critErr) { Optimize-OperationalParameters $OpSuccess }
            }
        } while ( !$OpSuccess -and !$critErr)
        Write-Output -NoEnumerate $res
    }

    Function Get-myEWSBindWellKnownFolder {
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

    Function New-FolderFilter {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [string[]]$Folders,
            [string]$emailAddress
        )
        If ( $Folders) {
            $FolderFilterSet= [System.Collections.ArrayList]::Synchronized(@())
            ForEach ( $Folder in $Folders) {
                # Convert simple filter to (simple) regexp
                $Parts= $Folder -match '^(?<root>\\)?(?<keywords>.*?)?(?<sub>\\\*)?$'
                If ( !$Parts) {
                    Write-Error ('Invalid regular expression matching against {0}' -f $Folder)
                }
                Else {
                    $Keywords= Search-ReplaceWellKnownFolderNames -EwsService $EwsService -Criteria ($Matches.keywords) -EmailAddress $emailAddress
                    $EscKeywords= [Regex]::Escape( $Keywords) -replace '\\\*', '.*'
                    $Pattern= iif -eval $Matches.Root -tv '^\\' -fv '^\\(.*\\)*'
                    $Pattern += iif -eval $EscKeywords -tv $EscKeywords -fv ''
                    $Pattern += iif -eval $Matches.sub -tv '(\\.*)?$' -fv '$'
                    $Obj= [pscustomobject]@{
                        'Pattern'    = [string]$Pattern
                        'IncludeSubs'= [bool]$Matches.Sub
                        'OrigFilter' = [string]$Folder
                    }
                    $null= $FolderFilterSet.Add( $Obj)
                    Write-Debug ($Obj -join ',')
                }
            }
        }
        Else {
            $FolderFilterSet= $null
        }
        return $FolderFilterSet
    }

    Function Search-ReplaceWellKnownFolderNames {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [string]$criteria= '',
            [string]$emailAddress
        )
        $AllowedWKF= 'Inbox', 'Calendar', 'Contacts', 'Notes', 'SentItems', 'Tasks', 'JunkEmail', 'DeletedItems', 'Drafts'
        # Construct regexp to see if allowed WKF is part of criteria string
        ForEach ( $ThisWKF in $AllowedWKF) {
            If ( $criteria -match '#{0}#') {
                $criteria= $criteria -replace ('#{0}#' -f $ThisWKF), (Get-myEWSBindWellKnownFolder -EwsService $EwsService -WellKnownFolderName $ThisWKF -EmailAddress $emailAddress).DisplayName
            }
        }
        return $criteria
    }

    Function Get-SubFolders {
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [Microsoft.Exchange.WebServices.Data.Folder]$Folder,
            $IncludeFilter= $null,
            $ExcludeFilter= $null
        )
        $FoldersToProcess= [System.Collections.ArrayList]@()

        # Lookup table for full path creation - add 'root folder' initially
        $PathCache= @{}
        $ParentFolderCache= @{}
        $PathCache[$Folder.Id.UniqueId]= '\'
        $ParentFolderCache[ $Folder.ParentFolderId.UniqueId]= [Microsoft.Exchange.WebServices.Data.Folder]::Bind( $EwsService, $Folder.ParentFolderId)
        $Obj= New-Object -TypeName PSObject -Property @{
            Folder= $Folder
            FolderPath= '\'
            ParentFolder= $ParentFolderCache[ $Folder.ParentFolderId.UniqueId]
            ParentFolderPath= $null
        }
        $null= $FoldersToProcess.Add( $Obj)

        # Collect all folder ids and construct paths and parent relationships
        $FolderView= New-Object Microsoft.Exchange.WebServices.Data.FolderView( $script:FolderBatchSize['Current'])
        $FolderView.Traversal= [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
        $PropertySet= @(
            [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,
            [Microsoft.Exchange.WebServices.Data.FolderSchema]::ParentFolderId,
            [Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass)
        $FolderView.PropertySet= New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
            [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly, $PropertySet)
        Do {
            $FolderSearchResults= Find-myEWSFoldersNoSearch -EwsService $EwsService -Folder $Folder -FolderView $FolderView
            ForEach ( $FolderItem in $FolderSearchResults) {
                If( -not $ParentFolderCache[ $FolderItem.ParentFolderId.UniqueId]) {
                    $ParentFolderCache[ $FolderItem.ParentFolderId.UniqueId]= [Microsoft.Exchange.WebServices.Data.Folder]::Bind( $EwsService, $FolderItem.ParentFolderId)
                }
                $ParentFolder= $ParentFolderCache[ $FolderItem.ParentFolderId.UniqueId]
                $FolderPath= Join-Path -Path $PathCache[ $FolderItem.ParentFolderId.Uniqueid] -ChildPath $FolderItem.DisplayName
                $PathCache[$FolderItem.Id.Uniqueid]= $FolderPath
                $Obj= New-Object -TypeName PSObject -Property @{
                    Folder= $FolderItem
                    FolderPath= $FolderPath
                    ParentFolder= $ParentFolder
                    ParentFolderPath= $PathCache[ $FolderItem.ParentFolderId.UniqueId]
                }
                $null= $FoldersToProcess.Add( $Obj)
            }
            $FolderView.Offset += $FolderSearchResults.Folders.Count
        } While ($FolderSearchResults.MoreAvailable)

        If ( $IncludeFilter) {
            $FilteredFolders= [System.Collections.ArrayList]@()
            ForEach( $FolderItem in $FoldersToProcess) {
                $Keep= $false
                ForEach ( $Filter in $IncludeFilter) {
                    If ( $FolderItem.FolderPath -match $Filter.Pattern) {
                        $Keep= $true
                    }
                }
                If( $Keep) {
                    Write-Debug ( 'Including folder {0}' -f $FolderItem.FolderPath)
                    $null= $FilteredFolders.Add( $FolderItem)
                }
            }
        }
        Else {
            # Default is to include all folders
            $FilteredFolders= $FoldersToProcess
        }

        If ( $ExcludeFilter) {
            ForEach( $FolderItem in $FilteredFolders) {
                $Remove= $false
                ForEach ( $Filter in $ExcludeFilter) {
                    If ( $FolderItem.FolderPath -match $Filter.Pattern) {
                        $Remove= $true
                    }
                }
                If( $Remove) {
                    Write-Debug ( 'Removing excluded folder {0}' -f $FolderItem.FolderPath)
                    $null= $FilteredFolders.Remove( $FolderItem)
                }
            }
        }
        Else {
            # Default is to not exclude any folder
        }
        Write-Output -NoEnumerate $FilteredFolders
    }

    Function Move-MailboxContents {
        [CmdletBinding(SupportsShouldProcess=$true)]
        param(
            [Microsoft.Exchange.WebServices.Data.ExchangeService]$EwsService,
            [string]$Identity,
            [Microsoft.Exchange.WebServices.Data.Folder]$Source,
            [Microsoft.Exchange.WebServices.Data.Folder]$Target,
            $IncludeFilter,
            $ExcludeFilter,
            $emailAddress
        )

        $FoldersProcessed= 0
        $Result= $True

        # Build list of matching (sub)folders to process
        $FoldersToProcess= Get-SubFolders -EwsService $EwsService -Folder $Source -IncludeFilter $IncludeFilter -ExcludeFilter $ExcludeFilter
        Write-Debug ('Located {0} archive folders' -f $FoldersToProcess.Count)

        $ExistingFolders= Get-SubFolders -EwsService $EwsService -Folder $Target
        Write-Debug ('Collected {0} mailbox folders for destination mapping' -f $ExistingFolders.Count)

        If (!$NoProgressBar) {
            Write-Progress -Id 1 -Activity ('Processing {0}' -f $Identity) -Status 'Unarchiving folders and items'
        }

        ForEach ( $SubFolder in $FoldersToProcess) {

            If (!$NoProgressBar) {
                Write-Progress -Id 2 -Activity ('Processing {0}' -f $Identity) -Status ('Unarchiving folder {0} of {1}' -f $FoldersProcessed, $FoldersToProcess.Count) -PercentComplete ( $FoldersProcessed / $FoldersToProcess.Count * 100)
            }

            $FoldersProcessed++
            If ( $Force -or $PSCmdlet.ShouldProcess( ('Unarchiving folder {0}' -f $SubFolder.FolderPath))) {

                # Load first class properties (stats)
                $SubFolder.Folder.Load()

                Write-Host ('Checking {0} with {1} item(s) and {2} subfolder(s)' -f $SubFolder.FolderPath, $SubFolder.Folder.TotalCount, $SubFolder.Folder.ChildFolderCount)

                If($SubFolder.Folder.TotalCount -ne 0 -or $SubFolder.Folder.ChildFolderCount -ne 0) {

                    # Determine if we need recreate or merge
                    $MatchingFolder= $ExistingFolders | Where-Object { $_.FolderPath -eq $SubFolder.FolderPath } | Select-Object -First 1

                    If( -not $MatchingFolder) {
                        # No match, create target in folder with matching parent folder path (breath-first, should exist)
                        $TargetParentFolder= $ExistingFolders | Where-Object { $_.FolderPath -eq $SubFolder.ParentFolderPath } | Select-Object -First 1
                        $TargetFolder= New-myEWSFolder -EwsService $EwsService -Folder $TargetParentFolder.Folder -DisplayName $SubFolder.Folder.DisplayName -FolderClass $SubFolder.Folder.FolderClass

                        # Add newly created target folder to list of folders in mailbox
                        $NewExistingFolder= New-Object -TypeName PSObject -Property @{
                            FolderPath      = Join-Path -Path $SubFolder.ParentFolderPath -ChildPath $SubFolder.Folder.DisplayName
                            Folder          = $TargetFolder
                            ParentFolderPath= $TargetParentFolder.FolderPath
                            ParentFolder    = $TargetParentFolder.ParentFolder
                        }
                        $null= $ExistingFolders.Add( $NewExistingFolder)
                    }
                    Else {
                        # Match found, use existing target folder
                        $TargetFolder= $MatchingFolder.Folder
                    }

                    # Move items plus folders to the matching or new target
                    If( $TargetFolder) {

                        $SubFolder.Folder.Load()
                        If( $SubFolder.Folder.TotalCount -ne 0 ) {

                            # Move all items from folder in source to target
                            $ItemsFound= 0
                            $ItemView= New-Object Microsoft.Exchange.WebServices.Data.ItemView( $script:ItemBatchSize['Current'])
                            $ItemView.Traversal= [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Shallow
                            $ItemView.PropertySet= New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
                                [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
                            $ItemsToProcess = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]@())

                            # First, collect all Item Ids in current folder
                            Write-Verbose ('Looking for items to unarchive in {0} ..' -f $SubFolder.FolderPath)
                            Do {
                                $ItemResults= Find-myEWSItemsNoSearch -EwsService $EwsService -Folder $SubFolder.Folder -ItemView $ItemView
                                $ItemsFound+= $ItemResults.Items.Count
                                If (!$NoProgressBar) {
                                    Write-Progress -Id 3 -Activity ('Unarchiving items in folder {0}' -f $SubFolder.FolderPath) -Status ('Discovered {0} items' -f $ItemsFound)
                                }
                                ForEach ( $Item in $ItemResults) {
                                    $ItemsToProcess.Add( $Item.Id)
                                }
                                $ItemView.Offset += $ItemResults.Items.Count

                            } While( $ItemResults.MoreAvailable)

                            If( $ItemsFound -gt 0) {
                                Write-Verbose ('Located {0} item(s) in {1}' -f $ItemsFound, $SubFolder.FolderPath)
                            }

                            # Chop set of items in chunks to process
                            $ItemBatchIds = [System.Collections.Concurrent.ConcurrentBag[Microsoft.Exchange.WebServices.Data.ItemId]]::new()
                            $ItemsProcessed=0
                            ForEach( $Item in $ItemsToProcess) {
                                $ItemBatchIds.Add( $Item)
                                $ItemsProcessed++

                                # When cut-off for items or last item in batch reached, process the remainder
                                If( $ItemBatchIds.Count -ge $script:ItemBatchSize['Current'] -or $ItemsProcessed -eq $ItemsFound) {
                                    If (!$NoProgressBar) {
                                        Write-Progress -Id 3 -Activity ('Processing folder {0}' -f $SubFolder.FolderPath) -Status ('Unarchived {0} items of {1}' -f $ItemsProcessed, $ItemsFound) -PercentComplete ( $ItemsProcessed / $ItemsFound * 100)
                                    }
                                    If ( $Force -or $PSCmdlet.ShouldProcess( ('Unarchiving {1} item(s) from folder {0}' -f $SubFolder.FolderPath, $ItemBatch.Count))) {
                                        try {
                                            Write-Verbose ('Unarchiving {1} of {2} item(s) from {0}' -f $SubFolder.FolderPath, $ItemsProcessed, $ItemsFound)
                                            $null= Move-myEWSItems -EwsService $EwsService -ItemIds $ItemBatchIds -Folder $TargetFolder
                                        }
                                        catch {
                                            Write-Error ('Problem unarchiving items from folder {0}: {1}' -f $SubFolder.FolderPath, $_.Exception.Message)
                                            $Result= $False
                                        }
                                    }
                                    $ItemBatchIds.Clear()
                                }
                            }

                            # Perform some garbage collection
                            [System.GC]::Collect()
                            $null = [System.GC]::WaitForPendingFinalizers()
                        }
                        Else {
                            # No items to process in this folder
                        }
                        If (!$NoProgressBar) {
                            Write-Progress -Id 3 -Activity ('Processing folder {0}' -f $SubFolder.FolderPath) -Status 'Finished unarchiving items.' -Completed
                        }
                    }
                    Else {
                        # Folder to move has no known parent folder (should not happen)
                        Write-Error ('Cannot locate parent folder {0} for {1}' -f $SubFolder.ParentFolderPath, $SubFolder.FolderPath)
                        $Result= $False
                    }
                }
                Else {
                    # Empty folder, clean up later
                }
            }

            If (!$NoProgressBar) {
                Write-Progress -Id 2 -Activity ('Processing archive folder {0}' -f $SubFolder.FolderPath) -Status 'Finished unarchiving folder.' -Completed
            }

        } # ForEach SubFolder

        $FoldersProcessed= 0
        # Process folders but now start at the end (bottom up and depth first)
        $FoldersToClean= $FoldersToProcess.Clone()
        $FoldersToClean.Remove( $FoldersToClean[0]) # Do not try cleaning up root folder
        $FoldersToClean.Reverse()
        ForEach ( $SubFolder in $FoldersToClean) {

            If (!$NoProgressBar) {
                Write-Progress -Id 2 -Activity ('Empty archive folder cleanup for {0}' -f $Identity) -Status ('Cleaning up folder {0} of {1}' -f $FoldersProcessed, $FoldersToProcess.Count) -PercentComplete ( $FoldersProcessed / $FoldersToClean.Count * 100)
            }

            $FoldersProcessed++

            # After processing, see if archived folder is empty and we can remove it
            $SubFolder.Folder.Load()
            If($SubFolder.Folder.TotalCount -eq 0 -and $SubFolder.Folder.ChildFolderCount -eq 0) {
                If ( $Force -or $PSCmdlet.ShouldProcess( ('Removing empty archive folder {0}' -f $SubFolder.FolderPath))) {
                    Try {
                        If( Remove-myEWSFolder -EwsService $EwsService -Folder $SubFolder.Folder -DeleteMode 'HardDelete') {
                            Write-Verbose ('Empty archive folder {0} removed' -f $SubFolder.FolderPath)
                        }
                    }
                    Catch {
                        Write-Error ('Problem removing archive folder {0}: {1}' -f $SubFolder.FolderPath, $_.Exception.Message)
                    }
                }
            }

            If (!$NoProgressBar) {
                Write-Progress -Id 2 -Activity ('Empty archive folder cleanup for {0}' -f $Identity) -Status 'Finished' -Completed
            }

        }
        If (!$NoProgressBar) {
            Write-Progress -Id 1 -Activity ('Processing {0}' -f $Identity) -Status 'Finished unarchiving.' -Completed
        }

        Write-Output -NoEnumerate $Result
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
        Throw( 'Problem initializing Exchange Web Services using schema {0} and timezone {1}' -f $ExchangeSchema, $TZ.Id)
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
    $EwsService.KeepAlive= $true
    $EwsService.AcceptGzipEncoding= $true

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

        Write-Verbose 'Constructing folder matching rules'
        $IncludeFilter= New-FolderFilter $EwsService $IncludeFolders $EmailAddress
        $ExcludeFilter= New-FolderFilter $EwsService $ExcludeFolders $EmailAddress

        $PrimaryRootFolder= Get-myEWSBindWellKnownFolder $EwsService 'MsgFolderRoot' $EmailAddress -ShowVersion
        If ($null -ne $PrimaryRootFolder) {
            $ArchiveRootFolder= Get-myEWSBindWellKnownFolder $EwsService 'ArchiveMsgFolderRoot' $EmailAddress
            If ($null -ne $ArchiveRootFolder) {
                Write-Verbose ('Unarchiving contents to mailbox of {0} ({1})' -f $EmailAddress, $CurrentIdentity)
                If (! ( Move-MailboxContents -Identity $CurrentIdentity -Target $PrimaryRootFolder -Source $ArchiveRootFolder -EwsService $EwsService -emailAddress $emailAddress -IncludeFilter $IncludeFilter -ExcludeFilter $ExcludeFilter)) {
                    Write-Error ('Problem unarchiving of {0} ({1})' -f $EmailAddress, $CurrentIdentity)
                    Exit $ERR_PROCESSINGMAILBOX
                }
            }
            Else {
                Write-Error ('Cannot access archive of {0}: {1}' -f $EmailAddress, $_.Exception.Message)
                Exit $ERR_CANTACCESSMAILBOXSTORE
            }
            }
        Else {
            Write-Error ('Cannot access mailbox of {0}: {1}' -f $EmailAddress, $_.Exception.Message)
            Exit $ERR_CANTACCESSMAILBOXSTORE
        }

        If($IncludeRecoverableItems.IsPresent) {
            $RecoverableItemsDeletions= Get-myEWSBindWellKnownFolder $EwsService 'RecoverableItemsDeletions' $EmailAddress
            If ($null -ne $RecoverableItemsDeletions) {
                $ArchiveRecoverableItemsDeletions= Get-myEWSBindWellKnownFolder $EwsService 'ArchiveRecoverableItemsDeletions' $EmailAddress
                If ($null -ne $ArchiveRecoverableItemsDeletions) {
                    Write-Verbose ('Unarchiving recoverable deleted items from archive to mailbox of {0} ({1})' -f $EmailAddress, $CurrentIdentity)
                    If (! ( Move-MailboxContents -Identity $CurrentIdentity -Target $RecoverableItemsDeletions -Source $ArchiveRecoverableItemsDeletions -EwsService $EwsService -emailAddress $emailAddress -IncludeFilter $IncludeFilter -ExcludeFilter $ExcludeFilter)) {
                        Write-Error ('Problem unarchiving recoverable items of {0} ({1})' -f $EmailAddress, $CurrentIdentity)
                        Exit $ERR_PROCESSINGRECOVERABLEITEMS
                    }
                }
                Else {
                    Write-Error ('Cannot access archive recoverable items of {0}: {1}' -f $EmailAddress, $_.Exception.Message)
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
