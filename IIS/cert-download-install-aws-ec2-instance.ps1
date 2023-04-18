#
# This script will download a PFX file from S3, extract with password from SSM Parameter store,
# compare with currently installed certificate and apply it to IIS and other services if it`s newer.
# Before using you must prepare an S3 bucket and parameter in SSM.
#

[Net.ServicePointManager]::SecurityProtocol = "TLS12"

#########################################################################################################

# Set variables (change to yours)

# Certificate domain
$CertCN = "domain.com"

# Name of S3 bucket
$S3BucketName = "company-pfx-files"

# AWS Region where everything is happening
$S3Region = "eu-west-1"

# Name of PFX file in S3 bucket
$FileName = "domain_com_2023.pfx"

# SSM Parameter name 
$PfxPassParameter = "domain_com_pfx_pass"

# Days before expiration, no any action will be performed in case of too early launch
$DaysToExpiration = 31

# Where to store log file
$LogFile = "$PSScriptRoot\$CertCN-install-log.txt"

# Slack Webhook URL
$SlackUri = "https://hooks.slack.com/services/XXXXXXXXXX"

#########################################################################################################

$PfxFullPath = "$PSScriptRoot\$FileName"
$PfxPassword = (Get-SSMParameter -Name $PfxPassParameter -WithDecryption $true -Region $S3Region).Value
$CertStore = "Cert:\LocalMachine"

$expirationDate = (Get-Date).AddDays($DaysToExpiration)

# Define how to send messages with Slack
Function Send-MessageToSlack {
	Param (
	[Parameter(Mandatory=$true)]
	[string] $SendSubject,

	[Parameter(Mandatory=$true)]
        [string] $SendBody
	)
 
	$body = ConvertTo-Json @{
		pretext = "$SendSubject"
		text = "$SendBody"
    }
	
	Invoke-RestMethod -Uri $SlackUri -Method Post -Body $body -ContentType 'application/json' | Out-Null
}

# Define how to write log
Function WriteLog {
	Param (
	[Parameter(Mandatory=$true)]
	[string] $Message
	)

	"$(Get-Date) - $Message" | Out-File $LogFile -Append -Encoding UTF8
}

# Define how to convert values from binary to hex
Function Convert-ByteArrayToHex {

    [cmdletbinding()]

    param(
        [parameter(Mandatory=$true)]
        [Byte[]]
        $Bytes
    )

    $HexString = New-Object -TypeName "System.Text.StringBuilder" ($WMSVCCertBinaryHash.Length * 2)

    ForEach($byte in $Bytes){
        $HexString.AppendFormat("{0:x2}", $byte) | Out-Null
    }

    $HexString.ToString()
}

# Get current installed certificate for $CertCN and select the latest one
$CertCurrent = (Get-ChildItem -Path $CertStore\* -Exclude "Remote Desktop" | Where-Object {$_.Subject -match $CertCN} | Sort-Object -Property NotAfter -Descending | Select-Object -first 1)
if ($CertCurrent -eq $null) {
	Write "Current certificate for $CertCN not found, nothing to update"
	Return
	}

# Stop if it`s too early to check
if ( $CertCurrent.NotAfter -lt $expirationDate) { 
	Write "Time to check for new certificate!"
	} else {
	Write "More than $DaysToExpiration days left to expire, no need to worry."
	Return
}

try {
# Download PFX from S3 and save it locally
Copy-S3Object -BucketName $S3BucketName -Key $FileName -LocalFile $PfxFullPath -Region $S3Region

# Importing modules
Import-Module WebAdministration

# Get current binded certificates and compare with installed
try {
	$AllCerts = Get-WebBinding -Protocol "https" | ForEach-Object { $_.certificateHash}
	$CertBinded = Get-ChildItem -Path $CertStore\* -Exclude "Remote Desktop" | Where {$AllCerts -contains $_.Thumbprint -and $_.Subject -match $CertCN -and $_.Thumbprint -ne $CertCurrent.Thumbprint}
	if ($CertBinded -ne $null) {
		 Send-MessageToSlack -SendSubject "Warning!" -SendBody "Latest certificate for $($CertCN) on $($env:ComputerName) is not matched with binded in IIS! It will expire on $($CertBinded.NotAfter), latest - on $($CertCurrent.NotAfter)."
		 
		# Rebind cert to the latest. If uncomment - change Slack message above.
		# Get-WebBinding | Where-Object { $_.certificateHash -eq $CertBinded.Thumbprint} | ForEach-Object {
		# $_.AddSslCertificate($CurrentCert.Thumbprint, 'My')
		# }  -ErrorAction Stop
		}
	} catch {
	Send-MessageToSlack -SendSubject "Error comparing certificates at $($env:ComputerName)" -SendBody "$_.Exception.Message"
	Return
}

# Get the expire date of certificate in PFX
$CertNew = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
$CertNew.Import($PfxFullPath, $PfxPassword, "PersistKeySet")

	} catch {
	WriteLog $_.Exception.Message
	Send-MessageToSlack -SendSubject "Failed to check new certificate at $($env:ComputerName)" -SendBody "$_.Exception.Message"
	Return
}

# Compare expire dates and if downloaded cert is newer - import it into local store
if ( $CertNew.NotAfter -gt $CertCurrent.NotAfter ) {
    
	try {

    # We are using Import-PfxCertificate here and not the .NET function for reason! Don`t even try to save cert using System.Security.Cryptography.X509Certificates.X509Certificate2 - it will fail with completely unrelated error!
    $PfxSecurePassword = $PfxPassword | ConvertTo-SecureString -AsPlainText -Force
    Import-PfxCertificate -Filepath $PfxFullPath -CertStoreLocation $CertStore\My -Password $PfxSecurePassword

    $NewCertThumbprint = $CertNew.Thumbprint
    $CurrentCertThumbprint = $CertCurrent.Thumbprint

    # Try to apply new cert for IIS sites with specific binding
    #$WebBinding = Get-WebBinding -Protocol "https" | Where-Object {$_.bindingInformation -match $CertCN}
    #if ($WebBinding -ne $null) {$WebBinding.AddSslCertificate($NewCertThumbprint, "My")}
    Get-WebBinding | Where-Object { $_.certificateHash -eq $CurrentCertThumbprint} | ForEach-Object {
        $_.AddSslCertificate($NewCertThumbprint, 'My')
        }  -ErrorAction Stop

    # Try to apply new cert for RDGW, if exist
    $RDSModule = Get-Module -ListAvailable -Name RemoteDesktopServices
    if ($RDSModule -ne $null) {
	Import-Module $RDSModule
		if ((Get-Item -Path RDS:\GatewayServer\SSLCertificate\Thumbprint).CurrentValue -eq $CurrentCertThumbprint) {
			Set-Item -Path RDS:\GatewayServer\SSLCertificate\Thumbprint -Value $NewCertThumbprint
		}
	}

    # Try to apply new cert for Web Management service (WebDeploy), if exist
	$WMSVCRegPath = "HKLM:\SOFTWARE\Microsoft\WebManagement\Server"
    if (Test-Path $WMSVCRegPath) {
		$WMSVCCertBinaryHash = (Get-ItemProperty -Path $WMSVCRegPath -Name "SslCertificateHash").SslCertificateHash
		$WMSVCCertThumbprint = Convert-ByteArrayToHex $WMSVCCertBinaryHash
	
		if ($WMSVCCertThumbprint -eq $CurrentCertThumbprint) {
			$WebDeployService = Get-Service -Name WMSVC
			try {
				$WebDeployService.Stop()
				$WebDeployService.WaitForStatus("Stopped", "00:00:30")
			} catch {
				Stop-Process -Name "wmsvc" -Force
			}

			$webManagementPort = (Get-ItemProperty -Path $WMSVCRegPath -Name "Port").Port
			$webManagementIP = (Get-ChildItem -Path IIS:\SslBindings | Where-Object Port -eq $webManagementPort).IPAddress.IPAddressToString
			Get-ChildItem -Path IIS:\SslBindings | Where-Object Port -eq $webManagementPort | Where-Object IPAddress -eq $webManagementIP | Remove-Item -ErrorAction Stop
			Get-Item -Path $CertStore\My\$NewCertThumbprint | New-Item -Path IIS:\SslBindings\$webManagementIP!$webManagementPort -ErrorAction Stop

			$bytes = for($i = 0; $i -lt $NewCertThumbprint.Length; $i += 2) { [convert]::ToByte($NewCertThumbprint.SubString($i, 2), 16) }
			Set-ItemProperty -Path $WMSVCRegPath -Name SslCertificateHash -Value $bytes -ErrorAction Stop

			Start-Service $WebDeployService -ErrorAction Stop
		}
	}

 
    # Write success message to log
    WriteLog "New certificate for $($CertCN) imported and applied successfully."
    Send-MessageToSlack -SendSubject "New cert at $($env:ComputerName)" -SendBody "Certificate for $CertCN imported and applied successfully."
	} catch {
	WriteLog $_.Exception.Message
	Send-MessageToSlack -SendSubject "Failed to apply new certificate at $($env:ComputerName)" -SendBody "$_.Exception.Message"
		}
    }
else  {
    # Do nothing (except logging) in case the newer cert is no newer than currently installed
    WriteLog "No new certificate for $($CertCN)"
	}

# Cleanup expired certificates (uncomment if needed)
#$AllMyCerts = Get-ChildItem -Path $CertStore -Recurse
#	Foreach($Cert in $AllMyCerts) {
#	if($Cert.NotAfter -lt (Get-Date)) { $Cert | Remove-Item }
# }

# Remove downloaded files
if (Test-Path $PfxFullPath) {Remove-Item $PfxFullPath}
