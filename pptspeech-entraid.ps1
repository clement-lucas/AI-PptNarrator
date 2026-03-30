# =============================================================================
# pptspeech-entraid.ps1
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
#   3. Run: .\pptspeech-entraid.ps1
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
Get-ChildItem $inDir -Filter *.xml | ForEach-Object {
    $ssml    = Get-Content $_.FullName -Raw
    $outFile = Join-Path $outDir ($_.BaseName + ".wav")

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

Write-Host "Done."
