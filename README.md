# How it works
* The script `OpenAndSaveExcelFile.ps1` will check each 3 seconds if the file `File.xslx` exists AND if the file `success.flag` does not exist.
* If so, the file `File.xlsx` is opened, the Excel formulas are recalculated and the file is saved and closed.
* Afterwards, the file `success.flag` is created.

How can another process use this?
* Another process manipulates some cells in an Excel file but is unable to recalculate the formulas.
* Thus, it saves the Excel file as `File.xlsx`.
* It then deletes the file `success.flag` (if existent).
* It waits until the file `success.flag` appears.
* Then, the process can read out the cells of interest from the Excel file (that have been recalculated).

# Start the script
Execute the script `OpenAndSaveExcelFile.ps1` with PowerShell ISE. It will run in an indefinite loop.

# Making the PowerShell script executable
You can either, **open the PowerShell ISE** and execute the script in there. For that purpose, right-click the file `OpenAndSaveExcelFile.ps1` -> Edit.
Otherwise, if you want call the script from somewhere else, say, cmd, then follow the steps below.

## Change ExecutionPolicy to "Unrestricted"
This does not work in all environments because your computers execution policy may be overruled by your organizations execution policy.
```ps1
Set-ExecutionPolicy -ExecutionPolicy Unrestricted
```

## ExecutionPolicy bypass
Another proposed solution is to call PowerShell as follows:
```cmd
powershell.exe -file .\OpenAndSaveExcelFile.ps1 -executionpolicy bypass 
```

## Certify the sript
If the upper commands do not achieve the goal because there are rules in place that only allow for certified scripts to be exectued, then follow the steps described below ([source](https://adamtheautomator.com/how-to-sign-PowerShell-script/)).

### Certification preparations - Do this only once
```ps1
# Generate a self-signed Authenticode certificate in the local computer's personal certificate store.
 $authenticode = New-SelfSignedCertificate -Subject "ATA Authenticode" -CertStoreLocation Cert:\LocalMachine\My -Type CodeSigningCert

# Add the self-signed Authenticode certificate to the computer's root certificate store.
## Create an object to represent the LocalMachine\Root certificate store.
 $rootStore = [System.Security.Cryptography.X509Certificates.X509Store]::new("Root","LocalMachine")
## Open the root certificate store for reading and writing.
 $rootStore.Open("ReadWrite")
## Add the certificate stored in the $authenticode variable.
 $rootStore.Add($authenticode)
## Close the root certificate store.
 $rootStore.Close()
 
# Add the self-signed Authenticode certificate to the computer's trusted publishers certificate store.
## Create an object to represent the LocalMachine\TrustedPublisher certificate store.
 $publisherStore = [System.Security.Cryptography.X509Certificates.X509Store]::new("TrustedPublisher","LocalMachine")
## Open the TrustedPublisher certificate store for reading and writing.
 $publisherStore.Open("ReadWrite")
## Add the certificate stored in the $authenticode variable.
 $publisherStore.Add($authenticode)
## Close the TrustedPublisher certificate store.
 $publisherStore.Close()
 
 # Confirm if the self-signed Authenticode certificate exists in the computer's Personal certificate store
 Get-ChildItem Cert:\LocalMachine\My | Where-Object {$_.Subject -eq "CN=ATA Authenticode"}
# Confirm if the self-signed Authenticode certificate exists in the computer's Root certificate store
 Get-ChildItem Cert:\LocalMachine\Root | Where-Object {$_.Subject -eq "CN=ATA Authenticode"}
# Confirm if the self-signed Authenticode certificate exists in the computer's Trusted Publishers certificate store
 Get-ChildItem Cert:\LocalMachine\TrustedPublisher | Where-Object {$_.Subject -eq "CN=ATA Authenticode"}
```

### Certifiy specific script
This has to be done once to get the script to work.  
If you apply changes to the script, this step has to be repeated.

```ps1
# Get the code-signing certificate from the local computer's certificate store with the name *ATA Authenticode* and store it to the $codeCertificate variable.
 $codeCertificate = Get-ChildItem Cert:\LocalMachine\My | Where-Object {$_.Subject -eq "CN=ATA Authenticode"}

# Sign the PowerShell script
# PARAMETERS:
# FilePath - Specifies the file path of the PowerShell script to sign, eg. C:\ATA\myscript.ps1.
# Certificate - Specifies the certificate to use when signing the script.
# TimeStampServer - Specifies the trusted timestamp server that adds a timestamp to your script's digital signature. Adding a timestamp ensures that your code will not expire when the signing certificate expires.
 Set-AuthenticodeSignature -FilePath ".\OpenAndSaveExcelFile.ps1" -Certificate $codeCertificate -TimeStampServer http://timestamp.digicert.com
```
