# OneDrive Maintenance Script

Scripts for OneDrive maintenance and inventory:

* OneDriveInventory.ps1 takes inventory of OneDrive storage usage
* SetOneDriveConditionalAccessPolicy.ps1 sets CA policy to AllowLimitedAccess for all OneDrives

In order to authenticate even when legacy authentication is disabled (LegacyAuthProtocolsEnabled = false), we create and App registration and use certificate authentication.

* Create an App registration in Azure AD and grant necessary SPO permissions (Sites.FullControl.All). This has to be application permissions and not delegated permission.
* Generate a new certificate with New-PnPAzureCertificate and Upload the certficate to the App registration.
* Extract the certificate and the private key as PEM
* Paste the PEM formatted certificate into the config file (PEMCertificate)
* Encrypt the private key with DPAPI as the user running the script and paste it into the config file (PEMPrivateKey)

```PowerShell
Get-Clipboard | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Set-Clipboard
```
