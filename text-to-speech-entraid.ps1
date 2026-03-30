# =============================================================================
# text-to-speech-entraid.ps1
# Azure AI Speech – Text-to-Speech (TTS) using Microsoft Entra ID authentication
# =============================================================================
# Prerequisites:
#   - Az.Accounts PowerShell module: Install-Module Az.Accounts -AllowClobber
#   - Logged in: Connect-AzAccount
#   - Custom domain enabled on the Speech resource
#   - RBAC role "Cognitive Services Speech User" assigned to your account
#
# Usage:
#   1. Set $resourceName below to your Azure AI Speech resource name
#   2. Place SSML XML files in the .\ssml\ folder
#   3. Run: .\text-to-speech-entraid.ps1
#   4. WAV files will be created in the .\audio\ folder
# =============================================================================

Import-Module Az.Accounts

# ---------- Configuration ----------
$resourceName = "<YOUR-SPEECH-RESOURCE-NAME>"   # e.g. "my-speech-resource"

$inDir  = ".\ssml"
$outDir = ".\audio"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

# ---------- Acquire Entra ID token ----------
$tokenResult = Get-AzAccessToken -ResourceUrl "https://cognitiveservices.azure.com"

# Az.Accounts v3+ returns Token as SecureString — convert to plain text
if ($tokenResult.Token -is [System.Security.SecureString]) {
    $accessToken = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR(
        [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenResult.Token))
} else {
    $accessToken = $tokenResult.Token
}

# ---------- Process SSML files ----------
$files = @(Get-ChildItem $inDir -Filter *.xml)
$total = $files.Count
$overwriteAll = $false
$skipAll = $false

for ($i = 0; $i -lt $total; $i++) {
    $file = $files[$i]
    $ssml    = Get-Content $file.FullName -Raw
    $outFile = Join-Path $outDir ($file.BaseName + ".wav")

    Write-Progress -Activity "Generating audio" -Status "$($file.Name) ($($i+1)/$total)" `
        -PercentComplete (($i / $total) * 100)

    # Check for existing file
    if (Test-Path $outFile) {
        if ($skipAll) { Write-Host "Skipped: $($file.BaseName).wav"; continue }
        if (-not $overwriteAll) {
            $choice = Read-Host "$($file.BaseName).wav already exists. [S]kip / [O]verwrite / Skip [A]ll / Overwrite A[l]l"
            switch ($choice.ToUpper()) {
                'S' { Write-Host "Skipped: $($file.BaseName).wav"; continue }
                'A' { $skipAll = $true; Write-Host "Skipped: $($file.BaseName).wav"; continue }
                'L' { $overwriteAll = $true }
            }
        }
    }

    # Entra ID auth requires the custom-domain endpoint (not the regional endpoint)
    $uri = "https://$resourceName.cognitiveservices.azure.com/tts/cognitiveservices/v1"
    $headers = @{
        "Authorization"            = "Bearer $accessToken"
        "Content-Type"             = "application/ssml+xml"
        "X-Microsoft-OutputFormat" = "riff-16khz-16bit-mono-pcm"
        "User-Agent"               = "ppt-narration"
    }

    Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $ssml -OutFile $outFile
    Write-Host "Created: $outFile"
}
Write-Progress -Activity "Generating audio" -Completed
Write-Host "Done: $total files processed."
