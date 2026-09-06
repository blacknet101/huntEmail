# huntEmail — Certificate Authentication Setup

Use this guide to create a self-signed certificate and configure your Entra app registration for certificate-based auth (no client secret).

**Project root:** the folder containing `huntEmail.ps1`  
**Cert output folder:** `C:\certs` (recommended) or `cert\` under the project root.

**App details** (from `documentation/api_info.md` — replace with your own values):

- **Application (client) ID:** `<APPLICATION_CLIENT_ID>`
- **Directory (tenant) ID:** `<DIRECTORY_TENANT_ID>`

---

## Step 1: Create the self-signed certificate (PowerShell)

### Method 1 — New-SelfSignedCertificate (may prompt for smart card)

Run this **once** on the machine where you will run the Email Threat Response toolkit (or on a secure machine, then copy the `.pfx` to the target).

```powershell
# Optional: create a folder for the cert files
$CertFolder = "C:\certs"
if (-not (Test-Path $CertFolder)) { New-Item -Path $CertFolder -ItemType Directory -Force }

$CertName = "huntEmail-cert"
$CertPassword = Read-Host "Enter a password for the PFX file" -AsSecureString

$Cert = New-SelfSignedCertificate `
    -Subject "CN=$CertName" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec KeyExchange `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -NotAfter (Get-Date).AddYears(2)

# Export the PUBLIC certificate (.cer) — you upload this to Entra
$CerPath = Join-Path $CertFolder "huntEmail.cer"
Export-Certificate -Cert $Cert -FilePath $CerPath -Type CERT

# Export the PRIVATE key as PFX — you keep this secure and import on the machine that runs the toolkit
$PfxPath = Join-Path $CertFolder "huntEmail.pfx"
Export-PfxCertificate -Cert $Cert -FilePath $PfxPath -Password $CertPassword

Write-Host "Created: $CerPath (upload to Entra)" -ForegroundColor Green
Write-Host "Created: $PfxPath (import on workstation/jump box)" -ForegroundColor Green
Write-Host "Thumbprint: $($Cert.Thumbprint)" -ForegroundColor Cyan
```

If you get **`SCARD_W_CANCELLED_BY_USER`** or a smart card dialog even though you don’t use smart cards, Windows is still trying to use a smart card CSP for the key. Use **Method 2** below instead.

---

### Method 2 — CertificateRequest (avoids smart card / software-only key)

Use this when Method 1 fails with a smart card–related error. This creates the key in software only (no CSP/smart card), then exports `.cer` and `.pfx` to your folder. The cert is **not** automatically added to the cert store; you import the PFX in Step 3.

Prefer running the maintained script:

```powershell
.\cert\createCert.ps1
```

Or the equivalent inline form (BSTR allocated once and freed in `finally`):

```powershell
$CertFolder = "C:\certs"
if (-not (Test-Path $CertFolder)) { New-Item -Path $CertFolder -ItemType Directory -Force }

$CertName = "huntEmail-cert"
$CerPath = Join-Path $CertFolder "huntEmail.cer"
$PfxPath = Join-Path $CertFolder "huntEmail.pfx"

# Create key and cert in memory (software only — no smart card)
$rsa = [System.Security.Cryptography.RSA]::Create(2048)
$request = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
    "CN=$CertName",
    $rsa,
    [System.Security.Cryptography.HashAlgorithmName]::SHA256,
    [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
)
$notBefore = Get-Date
$notAfter = $notBefore.AddYears(2)
$cert = $request.CreateSelfSigned($notBefore, $notAfter)

# Export .cer (public only — upload to Entra)
[System.IO.File]::WriteAllBytes($CerPath, $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))

# Export .pfx (password required)
$password = Read-Host "Enter a password for the PFX file" -AsSecureString
$bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
try {
    $passwordPlain = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    $pfxBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pkcs12, $passwordPlain)
    [System.IO.File]::WriteAllBytes($PfxPath, $pfxBytes)
} finally {
    if ($bstr -ne [IntPtr]::Zero) {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

Write-Host "Created: $CerPath (upload to Entra)" -ForegroundColor Green
Write-Host "Created: $PfxPath (import on workstation/jump box)" -ForegroundColor Green
Write-Host "Thumbprint: $($cert.Thumbprint)" -ForegroundColor Cyan
```

- **huntEmail.cer** → Upload to Entra (Step 2).
- **huntEmail.pfx** → Import on the machine where you run the toolkit (Step 3); use the thumbprint shown for Connect-MgGraph (Step 4).

If you already created the cert and only need the thumbprint later (cert in store from Method 1):

```powershell
Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -like "*huntEmail*" } | Select-Object Thumbprint, Subject, NotAfter
```

---

## Step 2: Upload the certificate to Entra

1. Open **Microsoft Entra admin center** → **Applications** → **App registrations**.
2. Open your **huntEmail** app registration (`<APPLICATION_CLIENT_ID>`).
3. Go to **Certificates & secrets**.
4. Click **Upload certificate**.
5. Choose the **huntEmail.cer** file you exported in Step 1.
6. Give it a description (e.g. `huntEmail self-signed – expires YYYY-MM`) and click **Add**.

Entra now trusts this certificate for authentication. You do **not** upload the `.pfx` to Entra.

---

## Step 3: Import the PFX on the machine that runs the toolkit

Do this on the workstation or jump box where you run `huntEmail`, `purgeEmail`, and `timelineEmail`.

**Option A — Import helper script**

```powershell
$pwd = Read-Host "Enter PFX password" -AsSecureString
.\Import-HuntEmailCert.ps1 -PfxPath "C:\certs\huntEmail.pfx" -Password $pwd
```

**Option B — Import manually**

```powershell
$PfxPath = "C:\certs\huntEmail.pfx"
$Password = Read-Host "Enter PFX password" -AsSecureString
Import-PfxCertificate -FilePath $PfxPath -CertStoreLocation Cert:\CurrentUser\My -Password $Password
```

**Option C — Cert is already in Current User\My from Step 1**

If you ran Step 1 on this same machine, the cert is already in `Cert:\CurrentUser\My`. You only need to export the PFX for backup or for importing on another machine (e.g. jump box). No need to import again on that same machine.

After import, get the thumbprint:

```powershell
Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -like "*huntEmail*" } | Select-Object Thumbprint, Subject, NotAfter
```

---

## Step 4: Connect to Microsoft Graph with the certificate

Use the **thumbprint** from Step 1 or Step 3 (they are the same for the same cert).

```powershell
$env:HUNTEMAIL_TENANT_ID = "<DIRECTORY_TENANT_ID>"
$env:HUNTEMAIL_CLIENT_ID = "<APPLICATION_CLIENT_ID>"
$env:HUNTEMAIL_CERT_THUMBPRINT = "<CERTIFICATE_THUMBPRINT>"

.\Connect-HuntEmailGraph.ps1
```

Or connect directly:

```powershell
$TenantId   = "<DIRECTORY_TENANT_ID>"
$ClientId   = "<APPLICATION_CLIENT_ID>"
$Thumbprint = "<CERTIFICATE_THUMBPRINT>"

Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $Thumbprint -NoWelcome
```

Verify:

```powershell
Get-MgContext
# Should show your ClientId and AuthType Certificate
```

Then run the toolkit as usual, e.g.:

```powershell
. .\huntEmail.ps1
huntEmail -Sender "test@example.com" -StartDate (Get-Date).AddDays(-1)
```

---

## Optional: Add to your PowerShell profile

If you want the huntEmail connection to be one command, add a function (do **not** put the PFX password in the profile). Only the thumbprint is in the profile; the private key stays in the cert store.

```powershell
function Connect-HuntEmailGraph {
    $TenantId   = "<DIRECTORY_TENANT_ID>"
    $ClientId   = "<APPLICATION_CLIENT_ID>"
    $Thumbprint = "<CERTIFICATE_THUMBPRINT>"
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $Thumbprint -NoWelcome
}
```

Then: `Connect-HuntEmailGraph`

Prefer the shipped `Connect-HuntEmailGraph.ps1` plus environment variables so secrets stay out of your profile script.

---

## Checklist

- [ ] Create self-signed cert; export `.cer` and `.pfx` (Step 1).
- [ ] Upload **huntEmail.cer** in Entra → your app → Certificates & secrets (Step 2).
- [ ] Import **huntEmail.pfx** on the machine that runs the toolkit (Step 3).
- [ ] Grant admin consent for **Mail.Read**, **Mail.ReadWrite**, **User.Read.All** on the app (if not already done).
- [ ] Connect with `Connect-HuntEmailGraph.ps1` / `Connect-MgGraph -CertificateThumbprint` and test hunt/purge/timeline (Step 4).
- [ ] Store PFX and password securely; never commit `.pfx` / `.cer` to git.
- [ ] Set a calendar reminder to renew the cert before it expires (e.g. 2 years from creation).

---

## Rotating the certificate later

1. Create a new self-signed cert (repeat Step 1 with a new name or date).
2. Upload the new `.cer` in Entra (Step 2). You can have two certs at once.
3. Import the new PFX where you run the toolkit and switch your connection script/profile/env vars to the new thumbprint.
4. Test; then remove the old certificate from Entra and delete the old PFX from disk.
