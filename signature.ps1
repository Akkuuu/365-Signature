#This goes the beginning again to detect
$signatureFolderRoot = Join-Path -Path $env:APPDATA -ChildPath "Microsoft\Signatures\"
$signatureFolder = Join-Path -Path $signatureFolderRoot -ChildPath "YOURCOMPANY2023 ($($Mail))_files"
$folderNamePattern = "*YOURCOMPANY2023*"
$matchingFolders = Get-ChildItem -Path $signatureFolderRoot -Directory -Filter $folderNamePattern
if ($matchingFolders -eq $null) {
$PassCode = 'YOUR PASS CODE'
# Get the current user's domain and username from 'whoami' command
$domainUser = whoami

# Split the domain and username
$domain, $userName = $domainUser -split '\\'

# Print the username without the domain
Write-Host "Current user's username: $userName"

if ($userName -contains "Admin" -or $userName -contains "Workstation"){
exit 0
}

# Make API call
$Uri = "https://YOUR-AMACON-FUNCTION.net/api/YOURCOMPANYSignatureT?displayName=$($userName)&PassCode=$($PassCode)"
write-host $Uri
$Headers = @{
    "x-functions-key" = "FUNCTION HEADER KEY"
}

$RandomSeconds = Get-Random -Minimum 1 -Maximum 10
Start-Sleep -Seconds $RandomSeconds
$Response = Invoke-RestMethod -Method Post -Uri $Uri -Headers $Headers -ContentType "application/json"

if ($Response -ne $null) {
    $DisplayName = $Response.DisplayName
    $JobTitle = $Response.JobTitle
    $Department = $Response.Department
    $Mail = $Response.Mail
    $City = $Response.City
    Write-Host "User's Display Name: $($Response.DisplayName)"
    Write-Host "User's Job Title: $($Response.JobTitle)"
    Write-Host "User's Department: $($Response.Department)"

    # Process the values as needed
} else {
    Write-Host "User not found"
}

# Define local paths
$downloadPath = "C:\YOURCOMPANYIT\Signatures.zip"
$extractPath = "$env:APPDATA\Microsoft\Signatures\YOURCOMPANY2023 ($($Mail))_files"

# Check if the YOURCOMPANY2023 folder and the HTML file exist
$signatureFolderRoot = Join-Path -Path $env:APPDATA -ChildPath "Microsoft\Signatures\"
$signatureFolder = Join-Path -Path $signatureFolderRoot -ChildPath "YOURCOMPANY2023 ($($Mail))_files"
$signatureFilePath = Join-Path -Path $signatureFolder -ChildPath "YOURCOMPANY2023.htm ($($Mail))"


# Download the zip file from Azure Blob Storage
$RandomSeconds = Get-Random -Minimum 1 -Maximum 10
Start-Sleep -Seconds $RandomSeconds
#This part downloads and opens image/signatures files to the path/
$signatureFiles = "YOUR ZIP LINK" 
Invoke-WebRequest -Uri $signatureFiles -OutFile $downloadPath

# Extract the zip file to the specific location
Expand-Archive -Path $downloadPath -DestinationPath $extractPath

# Vancouver
# Define the email signature in HTML format using the user information
$signatureContent = @"your signature HTML"@

# Locate the Outlook signature folder
$appData = [Environment]::GetFolderPath([Environment+SpecialFolder]::ApplicationData)
$signatureFolder = Join-Path -Path $appData -ChildPath "Microsoft\Signatures\YOURCOMPANY2023 ($($Mail))_files"
$signatureFolderRoot = Join-Path -Path $appData -ChildPath "Microsoft\Signatures\"

# Check if the folder exists and create it if necessary
if (-not (Test-Path -Path $signatureFolder)) {
    New-Item -Path $signatureFolder -ItemType Directory | Out-Null
}

# Create the HTML file with the signature content
$signatureFilePath = Join-Path -Path $signatureFolderRoot -ChildPath "YOURCOMPANY2023 ($($Mail)).htm"

# Change signature based on city if needed.
if ($City -eq "Vancouver"){
Set-Content -Path $signatureFilePath -Value $signatureContent -Encoding UTF8
}
if ($City -eq "Toronto"){
Set-Content -Path $signatureFilePath -Value $signatureToronto -Encoding UTF8
}
if ($City -eq "Istanbul"){
Set-Content -Path $signatureFilePath -Value $signatureEdmonton -Encoding UTF8
}
if ($City -eq "Bangkok"){
Set-Content -Path $signatureFilePath -Value $signatureDenver -Encoding UTF8
}

# Change registry to make it default signature.
$registryPath = "HKCU:\Software\Microsoft\Office\Outlook\Settings\Data"
$valueName = "$($Mail)_Roaming_New_Signature"
$SecondValueName = "$($Mail)_Roaming_Reply_Signature"
$ThirdValueName = "$($Mail)_reply_signaturehtml"
$FourthValueName = "$($Mail)_signaturehtml"
$FifthValueName = "$($Mail)_signaturetext"
$SixthValueName = "$($Mail)_reply_signaturetext"
$profileInfo = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676\00000002" #Get Profile and change new signature
$profileValue = "New Signature"
$profileReplyValue = "Reply-Forward Signature"

# Get the current value and second.
$originalValue = (Get-ItemProperty -Path $registryPath -Name $valueName).$valueName
$SecondValue = (Get-ItemProperty -Path $registryPath -Name $SecondValueName).$SecondValueName
$ThirdValue = (Get-ItemProperty -Path $registryPath -Name $ThirdValueName).$ThirdValueName
$FourthValue = (Get-ItemProperty -Path $registryPath -Name $FourthValueName).$FourthValueName
$FifthValue = (Get-ItemProperty -Path $registryPath -Name $FifthValueName).$FifthValueName
$SixthValue = (Get-ItemProperty -Path $registryPath -Name $SixthValueName).$SixthValueName

# Convert the JSON string to a PowerShell object
$jsonValue = $originalValue | ConvertFrom-Json
$SecondjsonValue = $SecondValue | ConvertFrom-Json
$ThirdjsonValue = $ThirdValue | ConvertFrom-Json
$FourthjsonValue = $FourthValue | ConvertFrom-Json
$FifthjsonValue = $FourthValue | ConvertFrom-Json
$SixthjsonValue = $FourthValue | ConvertFrom-Json

# Update the 'value' property
$newValue = "YOURCOMPANY2023" # Replace this with the new value you want to set
$newValuewithMail = "YOURCOMPANY2023 ($Mail)"
$jsonValue.value = $newValue

# Convert the updated object back to a JSON string
$updatedJsonValue = $jsonValue | ConvertTo-Json -Compress
$SecondupdatedJsonValue = $SecondjsonValue | ConvertTo-Json -Compress
$ThirdupdatedJsonValue = $ThirdjsonValue | ConvertTo-Json -Compress
$FourthupdatedJsonValue = $FourthjsonValue | ConvertTo-Json -Compress
$FifthupdatedJsonValue = $FifthjsonValue | ConvertTo-Json -Compress
$SixthupdatedJsonValue = $SixthjsonValue | ConvertTo-Json -Compress

# Update the registry value
Set-ItemProperty -Path $registryPath -Name $valueName -Value $updatedJsonValue
Set-ItemProperty -Path $registryPath -Name $SecondValueName -Value $SecondupdatedJsonValue
Set-ItemProperty -Path $registryPath -Name $ThirdValueName -Value $ThirdupdatedJsonValue
Set-ItemProperty -Path $registryPath -Name $FourthValueName -Value $FourthupdatedJsonValue
Set-ItemProperty -Path $registryPath -Name $FifthValueName -Value $FifthupdatedJsonValue
Set-ItemProperty -Path $registryPath -Name $SixthValueName -Value $SixthupdatedJsonValue
Set-ItemProperty -Path $profileInfo -Name $profileValue -Value $newValuewithMail
Set-ItemProperty -Path $profileInfo -Name $profileReplyValue -Value $newValuewithMail


Write-Output "Signature Inserted Successfully"
exit 1
}
else {
    Write-Output "YOURCOMPANY2023 folder and HTML file already exist. Exiting script."
    exit 0
}
