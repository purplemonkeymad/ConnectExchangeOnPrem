# ConnectExchangeOnPrem
A Powershell Module with a command for quick connects to Exchange On-Prem deployments.

## About

I wanted an easy way to just connect to an on-prem setup without having to specify URIs and server names all the time. As I might be connecting within a large number of setups. I wrote this command to auto-discover all the required information and give me an Exchange management environment with only a single command.

## Requirements

The command expects the computer to be part of a domain. You will also have to have at least one On-Prem Exchange server active on that domain. By default it will use Kerberos based authentication.

## Install

Save the module folder to a location in your PSModulePath environment variable. The per user location defaults to your `Documents\WindowsPowershell\Modules` folder. 

## Usage

After install you can just call the command. Or you can manually import the module using:

    Import-Module ConnectExchangeOnPrem

If you don't care which server you are connecting to you can run the command without any parameters.

    Connect-ExchangeOnPrem

If you want to connect to a specific Exchange server then you can specify the server name using the ComputerName parameter.

    Connect-ExchangeOnPrem -ComputerName exchange01.ad.contoso.com

In PS 5.0+ this parameter can be autocompleted from the list of discovered servers from AD.

If you need to connect to multiple different exchange environments, you can use the parameter Prefix to differentiate commands from each session. With the following command `Get-Mailbox` will be proxied as `Get-OnPremMailbox`.

    Connect-ExchangeOnPrem -Prefix OnPrem

## Author

https://github.com/purplemonkeymad

## Source

The current and future sources will be available at:

https://github.com/purplemonkeymad/ConnectExchangeOnPrem