# Requires -Version 5.0
# Requires -Module ActiveDirectory

<#
.SYNOPSIS
Imports a cert from WASC renewal into Exchange services, and enables it for SMTP. It also runs commands needed to configure
connectors and the hybrid configuration object.

.DESCRIPTION
Note that this script is intended to be run via the install script plugin from WASC via the batch script wrapper. As such, we use positional parameters to avoid issues with using a dash in the cmd line.
It is intended to be run on a single server. If you have multiple servers, you will need to run this script on each server that has Exchange installed and is part of the hybrid configuration.

Documentation referenced from https://technet.microsoft.com/en-us/library/aa997231(v=exchg.160).aspx

Proper information should be available here

https://github.com/simple-acme/simple-acme/wiki/Install-Script

or more generally, here

https://github.com/simple-acme/simple-acme/wiki/Example-Scripts

.PARAMETER NewCertThumbprint
The exact thumbprint of the cert to be enabled.

.PARAMETER ReceiveConnectorName
By default the script will configure the "Default Frontend <hostname>" receive connector for the current server. If you have
a different receive connector that you want to configure, you can specify it here. The script will attempt to find a receive
connector with this name and configure it with the new certificate.:

.PARAMETER SendConnectorName
By default the script will configure the send connector with the name starting with 'Outbound to Office 365*'. If you have a
different send connector that you want to configure, you can specify it here. The script will attempt to find a send connector
with this name and configure it with the new certificate.

.PARAMETER LeaveOldExchangeCerts
A bool (as an int, since bools are difficult to pass through as parameters) to determine if old exchange certs with the same CN should be deleted.

1 - Leaves old Exchange certs
0 - Deletes old Exchange certs

.PARAMETER DebugOn
Include this switch parameter to write debug outputs for troubleshooting

.EXAMPLE

ImportExchangeHybridSMTP.ps1 -NewCertThumbprint <certThumbprint>

Enables the new certificate for SMTP and configures the default send and receive connectors for the current server with the new certificate.

.EXAMPLE

ImportExchangeHybridSMTP.ps1 -NewCertThumbprint <certThumbprint> -ReceiveConnectorName "Inbound from Office 365"

Enables the new certificate for SMTP and configures the specified receive connector and the default send connector for the current server with the new certificate.

.EXAMPLE

ImportExchangeHybridSMTP.ps1 -NewCertThumbprint <certThumbprint> -ReceiveConnectorName "Inbound from Office 365" -SendConnectorName "Outbound to Office 365" -LeaveOldExchangeCerts 1

Enables the new certificate for SMTP and configures the specified receive and send connectors for the current server with the new certificate. Leaves old Exchange certs with the same CN.

.NOTES

#>

param(
	[Parameter(Position=0,Mandatory=$true)]
	[string]$NewCertThumbprint,
	[Parameter(Position=1,Mandatory=$false)]
	[string]$ReceiveConnectorName = "$($env:COMPUTERNAME)\Default Frontend $($env:COMPUTERNAME)",
	[Parameter(Position=2,Mandatory=$false)]
	[string]$SendConnectorName = '*Outbound to Office 365*',
	[Parameter(Position=3,Mandatory=$false)]
	[int]$LeaveOldExchangeCerts = 1,
	[Parameter(Position=4,Mandatory=$false)]
	[switch]$DebugOn
)

if($DebugOn){
	$DebugPreference = "Continue"
}

If($OSVersion -eq "Windows Server 2008 R2 Standard")
{
	Write-Error "Please upgrade Powershell version. See this URL for details: https://github.com/simple-acme/simple-acme/issues/1104"
	exit
}

# Print debugging info to make sure the parameters arrived
Write-Host "NewCertThumbprint: $NewCertThumbprint"
Write-Host "ReceiveConnectorName: $ReceiveConnectorName"
Write-Host "SendConnectorName: $SendConnectorName"
Write-Host "LeaveOldExchangeCerts: $LeaveOldExchangeCerts"

# Load Exchange Management shell
# https://github.com/simple-acme/simple-acme/issues/1372
$ServerFQDN = ([System.Net.Dns]::GetHostByName(($env:computerName))).Hostname
$ConnectionUri = "http://$ServerFQDN/PowerShell/"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri
Import-PSSession $Session

# Test if all Cmdlets are there now
$RequiredCmdlets = @("Enable-ExchangeCertificate", "Get-ExchangeCertificate", "Get-SendConnector", "Set-SendConnector", "Get-ReceiveConnector", "Set-ReceiveConnector", "Set-HybridConfiguration")
foreach ($Cmdlet in $RequiredCmdlets)
{
	$Command = Get-Command $Cmdlet -errorAction SilentlyContinue
	if ($null -eq $Command)
	{
		Write-Error "Cmdlet '$Cmdlet' not found. This cmdlet is required for the script to function. Exiting script."
		Exit
	}
}

try
{
	# Check if the new certificate is recognized by Exchange
	# Because this will only give back certificates that have a private key and are in the local machine store we make sure
	# the certificate is available and usable by Exchange.
	$ExchangeCert = Get-ExchangeCertificate -Thumbprint $NewCertThumbprint -ErrorAction SilentlyContinue
	if ($null -eq $ExchangeCert)
	{
		Write-Error "Certificate with thumbprint '$NewCertThumbprint' not found in Exchange. Exiting script."
		Exit
	}

	# Retrieve the current default SMTP certificate for later restoration.
	$DN = (Get-Exchangeserver $env:computername).distinguishedname
	$CurrentSMTPCert = (Get-ADObject -Identity $DN -Properties msExchServerInternalTLSCert).msExchServerInternalTLSCert
	$CurrentThumbprint = (New-Object system.security.cryptography.x509certificates.x509certificate2 -ArgumentList @(,$CurrentSMTPCert)).Thumbprint

	try
	{
		# Enable the new certificate for SMTP. This is required for Exchange to accept the new certificate for
		# connector configuration. The -Force parameter is used to avoid prompts for confirmation. However, this
		# will overwrite the current default SMTP certificate, so we will restore it later.
		Write-Host "Enabling Certificate $($ExchangeCert.Thumbprint) for Exchange services $($ExchangeServices)..."
		Enable-ExchangeCertificate -Services SMTP -Thumbprint $ExchangeCert.Thumbprint -Force -ErrorAction Stop

		# Restore the previous default SMTP certificate. This is required to avoid issues with Exchange services that rely on the default SMTP certificate.
		# Our new certificate will still be enabled for SMTP.
		Write-Host "Restoring previous default SMTP certificate with Thumbprint $($CurrentThumbprint)..."
		Enable-ExchangeCertificate -Services SMTP -Thumbprint $CurrentThumbprint -Force -ErrorAction Stop

		# Check if the new certificate is now enabled for SMTP. If not, we will throw an error and exit the script.
		$NewCertEnabledForSMTP = (Get-ExchangeCertificate -Thumbprint $NewCertThumbprint).Services -contains "SMTP"
		if (-not $NewCertEnabledForSMTP)
		{
			Write-Error "New certificate with thumbprint '$NewCertThumbprint' is not enabled for SMTP. Exiting script."
			Exit
		}
		# Remove the old certificate from Exchange services. This is required to avoid issues with Exchange services that rely on the default SMTP certificate.
		Write-Host "Removing old certificate with Thumbprint $($CurrentThumbprint) from Exchange..."

		# First we remove the certificate from send/receive connectors and hybrid configuration. If the new certificate would contain
		# the same CN and Issuer as the old certificate, Exchange will block removing the old one. By removing the certificate
		# string from connectors and hybrid configuration object, we are able to remove the old certificate.
		Write-Host "Removing certificate configuration string from send connector $($SendConnectorName)"
        Get-SendConnector $SendConnectorName | Set-SendConnector -TlsCertificateName $Null
		Write-Host "Removing certificate configuration string from receive connector $($ReceiveConnectorName)"
        Get-ReceiveConnector $ReceiveConnectorName | Set-ReceiveConnector -TlsCertificateName $Null
		Write-Host "Removing certificate configuration string from hybrid configuration object"
        Set-HybridConfiguration -TlsCertificateName $Null

		# Since we are changing attributes in AD, we need to wait for replication to occur before we can update the send/receive connectors
		# and hybrid configuration with the new certificate. We will wait for 60 seconds to allow for replication to occur in the same AD site.
		Write-Host "Waiting 60 seconds for AD replication to occur..."
		Start-Sleep -Seconds 60

		# Preparing the new 'TlsCertificateName' for the send/receive connectors and hybrid configuration. This is a string in the format of "<Issuer><Subject>".
        $TLSCertName = "<I>$($ExchangeCert.Issuer)<S>$($ExchangeCert.Subject)"
		Write-Host "New 'TlsCertificateName' for send/receive connectors and hybrid configuration object: $($TLSCertName)"
		# Now we can update the send/receive connectors and hybrid configuration with the new certificate.
		Write-Host "Updating the send connector $($SendConnectorName) with the value of $($TLSCertName)"
        Get-SendConnector $SendConnectorName | Set-SendConnector -TlsCertificateName $TLSCertName
		Write-Host "Updating the receive connector $($ReceiveConnectorName) for the current server with the value of $($TLSCertName)"
        Get-ReceiveConnector $ReceiveConnectorName | Set-ReceiveConnector -TlsCertificateName $TLSCertName
		Write-Host "Updating the hybrid configuration object with the value of $($TLSCertName)"
        Set-HybridConfiguration -TlsCertificateName $TLSCertName
		Write-Host "Certificate set"
	}

	catch
	{
		Write-Error "Error occurred when applying certificate to Exchange: $_"
		throw
	}

	if ($LeaveOldExchangeCerts -ne 1)
	{
		Write-Host "Old Exchange certificates being cleaned up"
		try
		{
			Get-ExchangeCertificate -DomainName $ExchangeCert.Subject.split(",")[0].split("=")[1] `
				| Where-Object -FilterScript {
					$_.Thumbprint -ne $NewCertThumbprint
				} `
			| Remove-ExchangeCertificate -Confirm:$false
		}

		catch
		{
			Write-Error "Error cleaning up old certificates Get-ExchangeCertificate/Remove-ExchangeCertificate: $_"
			throw
		}
	}
}

catch
{
    Write-Error "Script hasn't completed."
}
