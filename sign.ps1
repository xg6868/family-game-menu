param([string[]]$Targets)
$ErrorActionPreference = "Stop"
$subject = "XueZhiXia GameMenu"

# 0. Clean up old code-signing certs (incl. mojibake ones) from My/Root/TrustedPublisher
foreach ($loc in @("CurrentUser")) {
    foreach ($storeName in @("My","Root","TrustedPublisher")) {
        $st = New-Object System.Security.Cryptography.X509Certificates.X509Store($storeName,$loc)
        $st.Open("ReadWrite")
        $old = @($st.Certificates | Where-Object { $_.Extensions | Where-Object {
            $_.Oid.Value -eq "2.5.29.37" -and $_.Format($false) -match "1.3.6.1.5.5.7.3.3" } })
        foreach ($c in $old) {
            try { $st.Remove($c); Write-Output "Removed old cert $($c.Thumbprint) from $storeName" } catch {}
        }
        $st.Close()
    }
}

# 1. Create code-signing cert
$cert = New-SelfSignedCertificate -Type CodeSigningCert `
    -Subject "CN=$subject" `
    -KeyUsage DigitalSignature `
    -KeyAlgorithm RSA -KeyLength 2048 `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.3") `
    -NotAfter (Get-Date).AddYears(10)
Write-Output "Created cert: $($cert.Thumbprint)"

# 2. Trust for current user
foreach ($storeName in @("Root","TrustedPublisher")) {
    $st = New-Object System.Security.Cryptography.X509Certificates.X509Store($storeName,"CurrentUser")
    $st.Open("ReadWrite")
    $st.Add($cert)
    $st.Close()
    Write-Output "Trusted in $storeName"
}

# 3. Sign
foreach ($t in $Targets) {
    if (-not (Test-Path $t)) { Write-Output "Skip (missing): $t"; continue }
    $r = Set-AuthenticodeSignature -FilePath $t -Certificate $cert -HashAlgorithm SHA256
    Write-Output ("Signed {0}: {1}" -f (Split-Path $t -Leaf), $r.Status)
}
