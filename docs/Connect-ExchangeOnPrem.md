# Connect-ExchangeOnPrem

Connect to and import a implicit remoting session from an On-Premises Exchange deployment.

|Skip To|
|-------|
|[Syntax](#syntax) [Description](#description) [Parameters](#parameters) [Examples](#examples) [Links](#links)|

## Syntax

```powershell
Connect-ExchangeOnPrem [[-ComputerName] <String>] [[-Authentication] <String>] [[-Credential] <PSCredential>] [[-Prefix] <String>] [[-VersionString] <String>]
```

## Description

This command makes connection to an On-Premises exchange simple. It is intended to be an On-Premises equivalent of the Connect-ExchangeOnline command.

It will attempt to discover existing exchange servers in the current domain. It will then attempt to connect to each exchange server and import the first sucessful session.
The search order is determined by the server version, it will try newer exchange servers first.

## Parameters

This command provides the following parameters: [ComputerName](#computername) [Authentication](#authentication) [Credential](#credential) [Prefix](#prefix) [VersionString](#versionstring)

### ComputerName

     -ComputerName <String>

Specify the hostname of the exchange server to connect to.

Using this parameter will disabled the domain search for servers.

On PS 5.0+ this agument supports auto-complete from the list of Exchange servers discovered from AD.

|Type|required|pipelineInput|position|Aliases|
|---|---|---|---|---|
|String|false|false|1||

### Authentication

     -Authentication <String>

Use to override the Authentication method use to connect to the server.

|Type|required|pipelineInput|position|Aliases|
|---|---|---|---|---|
|String|false|false|2||

### Credential

     -Credential <PSCredential>

Credential to connect to the Exchange server. You should only need this if you change the authentication method away from the default Kerberos.

|Type|required|pipelineInput|position|Aliases|
|---|---|---|---|---|
|PSCredential|false|false|3||

### Prefix

     -Prefix <String>

Prefix commands from the implicit module with the specified string.

This comes from the module import so will Prefix to the Noun of the command.

|Type|required|pipelineInput|position|Aliases|
|---|---|---|---|---|
|String|false|false|4||

### VersionString

     -VersionString <String>

Only connect to servers that have the given admin version.
The version string is the same format as the admin version from Get-ExchangeServer ie "Version 15.0 (Build 1234.56)"

This parameter supports wildcards so "Version 15.2* will only connect to Exchange 2019 servers."

|Type|required|pipelineInput|position|Aliases|
|---|---|---|---|---|
|String|false|false|5||

## Examples

### -------------------------- EXAMPLE 1 --------------------------

    Connect-ExchangeOnPrem

With no parameters, the command will attempt to lookup a working Exchange server in AD. It will then connect using the default Kerberos authentication.

### -------------------------- EXAMPLE 2 --------------------------

    Connect-ExchangeOnPrem -VersionString "Version 15.0*"

Lookup server list from AD, but only try to connect to a server who's Admin Version matches the specified version.
In this example it will only connect to Exchange 2013 servers.

## Links

https://github.com/purplemonkeymad/ConnectExchangeOnPrem/tree/master/docs/Connect-ExchangeOnPrem.md
