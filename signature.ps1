    # Define the root signature folder path
    $signatureFolderRoot = Join-Path -Path $env:APPDATA -ChildPath "Microsoft\Signatures\"

    # Define the signature folder path
    $signatureFolder = Join-Path -Path $signatureFolderRoot -ChildPath "YOURCOMPANY ($($Mail))_files"

    # Define the folder name pattern
    $folderNamePattern = "*YOURCOMPANY*"

    # Get all folders matching the pattern
    $matchingFolders = Get-ChildItem -Path $signatureFolderRoot -Directory -Filter $folderNamePattern

    # Check if there are no matching folders
    if ($matchingFolders -eq $null) {
    # Define the passcode
    $PassCode = 'DEFINE-A-RANDOM-PASSWORD'
    
    # Get the current user's domain and username from 'whoami' command
    $domainUser = whoami

    # Split the domain and username
    $domain, $userName = $domainUser -split '\\'

    # Print the username without the domain
    Write-Host "Current user's username: $userName"

    # Check if the username contains 'Admin' or 'Workstation'
    if ($userName -contains "Admin" -or $userName -contains "Workstation"){
        exit 0
    }

    # Make API call to get user information
    $Uri = "https://*YOUR-FUNCTION-URL*/api/*FUNCTIONNAME*?displayName=$($userName)&PassCode=$($PassCode)"
    write-host $Uri
    $Headers = @{
        "x-functions-key" = "***"
    }
    #Puts random interval to decrease load on the function/svr.
    $RandomSeconds = Get-Random -Minimum 1 -Maximum 10
    Start-Sleep -Seconds $RandomSeconds
    $Response = Invoke-RestMethod -Method Post -Uri $Uri -Headers $Headers -ContentType "application/json"

    # Check if a response was received
    if ($Response -ne $null) {
        # Extract user information from response
        $DisplayName = $Response.DisplayName
        $JobTitle = $Response.JobTitle
        $Department = $Response.Department
        $Mail = $Response.Mail
        $City = $Response.City
        $Location = $Response.Location

        # Print user information
        Write-Host "User's Display Name: $($Response.DisplayName)"
        Write-Host "User's Job Title: $($Response.JobTitle)"
        Write-Host "User's Department: $($Response.Department)"
        Write-Host "User's Location: $($Response.Location)"

        # Process the values as needed
    } else {
        Write-Host "User not found"
    }

    # Define local paths for downloading and extracting signature files
    $downloadPath = "C:\SIGNATURE-PATH\Signatures.zip"
    $extractPath = "$env:APPDATA\Microsoft\Signatures\YOURCOMPANY ($($Mail))_files"

    # Check if the Amacon2023 folder and the HTML file exist
    $signatureFolderRoot = Join-Path -Path $env:APPDATA -ChildPath "Microsoft\Signatures"


    # Define the signature folder path
    $signatureFolder = Join-Path -Path $signatureFolderRoot -ChildPath "YOURCOMPANY ($($Mail))_files"

    # Define the signature file path
    $signatureFilePath = Join-Path -Path $signatureFolder -ChildPath "YOURCOMPANY.htm ($($Mail))"

    # Download the zip file from Azure Blob Storage
    $RandomSeconds = Get-Random -Minimum 1 -Maximum 10
    Start-Sleep -Seconds $RandomSeconds
    $credential = "https://YOUR-LINK-TO-ZIP-TO-EXTRACT-IMAGES-ETC/public/amacon.zip"
    Invoke-WebRequest -Uri $credential -OutFile $downloadPath

    # Extract the zip file to the specific location
    Expand-Archive -Path $downloadPath -DestinationPath $extractPath

    # Vancouver
    # Define the email signature in HTML format using the user information

$signatureContent = @"YOUR HTML"@

    # Locate the Outlook signature folder
    $appData = [Environment]::GetFolderPath([Environment+SpecialFolder]::ApplicationData)
    $signatureFolder = Join-Path -Path $appData -ChildPath "Microsoft\Signatures\YOURCOMPANY ($($Mail))_files"
    $signatureFolderRoot = Join-Path -Path $appData -ChildPath "Microsoft\Signatures\"

    <#
    ==============================
     Check if the folder exists and create it if necessary
    ==============================
    #>
    if (-not (Test-Path -Path $signatureFolder)) {
        New-Item -Path $signatureFolder -ItemType Directory | Out-Null
    }

    <#
    =========================================
     Create the HTML file with the signature content
    =========================================
    #>

    <# If you want some conditions
    $signatureFilePath = Join-Path -Path $signatureFolderRoot -ChildPath "YOURCOMPANY ($($Mail)).htm"
    if ($Location -contains "Vancouver"){
    Set-Content -Path $signatureFilePath -Value $signatureContent -Encoding UTF8
    }
    if ($Location -contains "Toronto"){
    Set-Content -Path $signatureFilePath -Value $signatureToronto -Encoding UTF8
    }
    Write-Host "City is : " $City
    Write-Host "Location is : " $Location
    #>
    <#

    ====================================
     Change registry to make it default signature.
    ====================================
    #>

    $registryPath = "HKCU:\Software\Microsoft\Office\Outlook\Settings\Data"
    $valueName = "$($Mail)_Roaming_New_Signature"
    $SecondValueName = "$($Mail)_Roaming_Reply_Signature"
    $ThirdValueName = "$($Mail)_reply_signaturehtml"
    $FourthValueName = "$($Mail)_signaturehtml"
    $FifthValueName = "$($Mail)_signaturetext"
    $SixthValueName = "$($Mail)_reply_signaturetext"
    $profileInfo = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\XXX\00000002" #Get Profile and change new signature
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
    $newValue = "Amacon2023" # Replace this with the new value you want to set
    $newValuewithMail = "Amacon2023 ($Mail)"
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
    }

    Write-Output "Signature Inserted Successfully"
    exit 1
    }
    else {
        Write-Output "YOURCOMPANY folder and HTML file already exist. Exiting script."
        exit 0
    }
