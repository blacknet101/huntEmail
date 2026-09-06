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
