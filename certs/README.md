# Create a sef-signed certificate for the installer

First, use the assembly information to create the certificate:

```bash
New-SelfSignedCertificate -Type Custom -Subject "CN=Insert image in Document, O=Luighi Viton-Zorrilla, C=US" -KeyUsage DigitalSignature -FriendlyName "Insert image to Document certificate" -KeySpec KeyExchange -CertStoreLocation "Cert:\CurrentUser\My" -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.3", "2.5.29.19={text}")
```

The `-KeySpec KeyExchange` option is important to sign the assemblies, otherwise it will fail (See https://stackoverflow.com/a/55800580).

Then, export this certificate to the .pfx format:
```bash
$password = ConvertTo-SecureString -String iid123 -Force -AsPlainText
Export-PfxCertificate -cert "Cert:\CurrentUser\My\<Certificate Thumbprint>" -FilePath cert_iid.pfx -Password $password
```

Export the signtool in Windows Development Kit to the system PATH. It is located under `C:\Program Files (x86)\Windows Kits\10\bin\10.0.22000.0\x64`.

Then, use this tool to sign the installer under the `setup\Installs` folder:
```bash
signtool sign /debug /f "..\..\certs\cert_iid.pfx" /p iid123 /d "Installer for IID Software" /fd SHA256 /t http://timestamp.digicert.com /v .\IIDInstaller-1.0.0.4-Release-x64.msi
```

## Install the certificate

You must issue the following command from the cert directory to sign assemblies.
```bash
sn -i cert_iid.pfx VS_KEY_2EC106B2AEE315EC
```

## Delete old certificates

To delete the certificates you can launch the Certificates Manager from Windows+R and writing  `certmgr`. Then, in the list of certificates you can delete the desired certificate.

## References

- https://docs.microsoft.com/en-us/windows/msix/package/create-certificate-package-signing
- https://docs.microsoft.com/en-us/windows/win32/seccrypto/signtool
- https://stackoverflow.com/a/40748646