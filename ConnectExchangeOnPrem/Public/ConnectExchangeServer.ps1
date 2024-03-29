<#
.SYNOPSIS
Connect to and import a implicit remoting session from an On-Premises Exchange deployment.

.DESCRIPTION
This command makes connection to an On-Premises exchange simple. It is intended to be an On-Premises equivalent of the Connect-ExchangeOnline command.

It will attempt to discover existing exchange servers in the current domain. It will then attempt to connect to each exchange server and import the first sucessful session.
The search order is determined by the server version, it will try newer exchange servers first.

.PARAMETER ComputerName
Specify the hostname of the exchange server to connect to.

Using this parameter will disabled the domain search for servers.

On PS 5.0+ this agument supports auto-complete from the list of Exchange servers discovered from AD.

.PARAMETER Authentication
Use to override the Authentication method use to connect to the server.

.PARAMETER Credential
Credential to connect to the Exchange server. You should only need this if you change the authentication method away from the default Kerberos.

.PARAMETER Prefix
Prefix commands from the implicit module with the specified string.

This comes from the module import so will Prefix to the Noun of the command.

.PARAMETER VersionString
Only connect to servers that have the given admin version.
The version string is the same format as the admin version from Get-ExchangeServer ie "Version 15.0 (Build 1234.56)"

This parameter supports wildcards so "Version 15.2* will only connect to Exchange 2019 servers."

.EXAMPLE
Connect-ExchangeOnPrem

With no parameters, the command will attempt to lookup a working Exchange server in AD. It will then connect using the default Kerberos authentication.

.EXAMPLE
Connect-ExchangeOnPrem -VersionString "Version 15.0*"

Lookup server list from AD, but only try to connect to a server who's Admin Version matches the specified version.
In this example it will only connect to Exchange 2013 servers.

.Link
https://github.com/purplemonkeymad/ConnectExchangeOnPrem/tree/master/docs/Connect-ExchangeOnPrem.md

#>

function Connect-ExchangeOnPrem {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [Alias("Computer")]
        [Alias("ServerName")]
        [Alias("Server")]
        [string]
        $ComputerName,
        [Parameter()]
        [ValidateSet("Kerberos","Default","Basic","Credssp","Digest","Negotiate")]
        [string]
        $Authentication,
        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential,
        [Parameter()]
        [string]
        $Prefix,
        [Parameter()]
        [string]
        $VersionString
    )
    
    begin {}
    process {}
    
    end {
        if (-not $ComputerName) {
            # lets try to lookup an exchange server
            try {
                $DNC = ([adsi]'LDAP://RootDSE').Get("Defaultnamingcontext")
                $Searcher = [adsisearcher]'(objectclass=msExchExchangeServer)'
                $Searcher.SearchRoot = "LDAP://CN=configuration,$DNC"
                $ExchangeServers = ([array]$Searcher.FindAll()) | Sort-Object {$_.Properties['serialnumber']} -Descending
            } catch {
                Write-Error "Unable to contact domain, Please specify a ComputerName." -Exception $_
                return
            }

            # if we are filtering the version, find only servers that match that wildcard

            if ($VersionString) {
                $ExchangeServers = $ExchangeServers | Where-Object {$_.Properties['serialnumber'] -like $VersionString}
            }

        
            # Check if we Got any results
            if ($ExchangeServers.count -eq 0){
                $PSCmdlet.ThrowTerminatingError(
                    ( New-Object System.Management.Automation.ErrorRecord -ArgumentList @(
                        [System.Management.Automation.RuntimeException]"No Exchange Servers could be found in $DNC",
                        'ExchangeOnPrem.LookupError',
                        [System.Management.Automation.ErrorCategory]::ObjectNotFound,
                        '(objectclass=msExchExchangeServer)'
                        )
                    )
                )
                return
            }
            Write-Verbose "Found $($ExchangeServers.Count) Servers"

            # Ask ad for all the matching objects, we create a custom query here so we only have to
            # make a single call to ldap.

            $LDAPServerQueryTemplate = "(&(objectclass=computer)(|{0}))"
            [array]$AllHostname = foreach ($ExchangeServer in $ExchangeServers){
                $CanonicalName = $ExchangeServer.Properties.name # properties are case sensitive
                $LDAPServerQuery = $LDAPServerQueryTemplate -f "(name=$CanonicalName)"
                ([adsisearcher]$LDAPServerQuery).FindAll().Properties.dnshostname | Where-Object {$_}
            }

            # Try to create a session with each of the servers.
            $index = 0
            do {
                
                $ComputerName = $AllHostname[$index]
                if (-not $ComputerName){
                    continue
                }

                $SessionSplat = @{
                    ConfigurationName = "Microsoft.Exchange"
                    ConnectionUri = "http://$ComputerName/PowerShell/"
                }
                if ($Credential){
                    $SessionSplat.Credential = $Credential
                }
                if ($Authentication){
                    $SessionSplat.Authentication =  $Authentication
                }
                $session = New-PSSession @SessionSplat
                $connectedServer = $ComputerName
                $index++
            } while ( (-not $session) -and ($index -lt $AllHostname.Count) )

        } else { # servername supplied

            $SessionSplat = @{
                ConfigurationName = "Microsoft.Exchange"
                ConnectionUri = "http://$ComputerName/PowerShell/"
            }
            if ($Credential){
                $SessionSplat.Credential = $Credential
            }
            if ($Authentication){
                $SessionSplat.Authentication =  $Authentication
            }
            $session = New-PSSession @SessionSplat
            $connectedServer = $ComputerName

        }

        # Check for a valid connection
        if ($session){
            # this sometimes causes issues with the import due to scope exposure.
            Remove-Variable Authentication
            $PrefixParam = @{}
            if ($Prefix){
                $PrefixParam['Prefix'] = $Prefix
            }
            Import-module (Import-PSSession -AllowClobber -Session $session  @PrefixParam) -Global -DisableNameChecking @PrefixParam
            $Global:ExchConnectedServer = $connectedServer
        } else {
            Write-Error "Failed to connect to any of the Exchange Servers."
            return
        }
    }

}

if ($PSVersionTable.psversion -gt ([version]'5.0')){

    Register-ArgumentCompleter -CommandName Connect-ExchangeOnPrem -ParameterName ComputerName -ScriptBlock {
        param($Command, $Parameter, $WordToComplete, $CommandAst, $FakeBoundParams)
        $DNC = ([adsi]'LDAP://RootDSE').Get("Defaultnamingcontext")
        $Searcher = [adsisearcher]'(objectclass=msExchExchangeServer)'
        $Searcher.SearchRoot = "LDAP://CN=configuration,$DNC"
        $ExchangeServers = [array]$Searcher.FindAll()
        $PossibleServers = $ExchangeServers.Properties.admindisplayname
        $LDAPServerQueryTemplate = "(&(objectclass=computer)(|{0}))"
        $NameQueryList = foreach ($CanonicalName in $PossibleServers){
            "(name=$CanonicalName)"
        }
        $LDAPServerQuery = $LDAPServerQueryTemplate -f ($NameQueryList -join '')
        [array]$PossibleDNSNames = ([adsisearcher]$LDAPServerQuery).FindAll().Properties.dnshostname
        if ($WordToComplete){
            $PossibleDNSNames = $PossibleDNSNames.where({
                $_ -like "$WordToComplete*"
            })
        }
        return $PossibleDNSNames
    }
}