# huntEmail API configuration template

Use this file as a local reference for the Microsoft Graph app registration that powers the toolkit.
Do not commit production IDs, thumbprints, passwords, secrets, or certificate material to source control.

## App registration

- **Application (client) ID:** `<APPLICATION_CLIENT_ID>`
- **Object ID:** `<SERVICE_PRINCIPAL_OBJECT_ID>`
- **Directory (tenant) ID:** `<DIRECTORY_TENANT_ID>`
- **Certificate thumbprint:** `<CERTIFICATE_THUMBPRINT>`

## Suggested local environment variables

```powershell
$env:HUNTEMAIL_TENANT_ID = "<DIRECTORY_TENANT_ID>"
$env:HUNTEMAIL_CLIENT_ID = "<APPLICATION_CLIENT_ID>"
$env:HUNTEMAIL_CERT_THUMBPRINT = "<CERTIFICATE_THUMBPRINT>"
```

## Import cert to personal store

```powershell
$PfxPath = "C:\certs\huntEmail.pfx"
$Password = Read-Host "Enter PFX password" -AsSecureString
Import-PfxCertificate -FilePath $PfxPath -CertStoreLocation Cert:\CurrentUser\My -Password $Password
Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq "<CERTIFICATE_THUMBPRINT>" }
```

## Connect once cert has been imported

```powershell
Connect-MgGraph -TenantId "<DIRECTORY_TENANT_ID>" -ClientId "<APPLICATION_CLIENT_ID>" -CertificateThumbprint "<CERTIFICATE_THUMBPRINT>" -NoWelcome
```

## Remove a cert from the cert store

```powershell
Remove-Item -Path "Cert:\CurrentUser\My\<CERTIFICATE_THUMBPRINT>"
```
