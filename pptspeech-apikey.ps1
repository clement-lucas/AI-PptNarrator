# =============================================================================
# pptspeech-apikey.ps1
# Azure AI Speech – Text-to-Speech (TTS) using API Key authentication
# =============================================================================
# Usage:
#   1. Set $region and $apiKey below
#   2. Place SSML XML files in the .\ssml\ folder
#   3. Run: .\pptspeech-apikey.ps1
#   4. WAV files will be created in the .\audio\ folder
# =============================================================================

# ---------- Configuration ----------
$region = "<YOUR-REGION>"       # e.g. "swedencentral", "eastus", "westeurope"
$apiKey = "<YOUR-API-KEY>"      # Azure AI Speech resource key (Key 1 or Key 2)

$inDir  = ".\ssml"
$outDir = ".\audio"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

# ---------- Process SSML files ----------
Get-ChildItem $inDir -Filter *.xml | ForEach-Object {
    $ssml    = Get-Content $_.FullName -Raw
    $outFile = Join-Path $outDir ($_.BaseName + ".wav")

    $uri = "https://$region.tts.speech.microsoft.com/cognitiveservices/v1"
    $headers = @{
        "Ocp-Apim-Subscription-Key" = $apiKey
        "Content-Type"               = "application/ssml+xml"
        "X-Microsoft-OutputFormat"   = "riff-16khz-16bit-mono-pcm"
        "User-Agent"                 = "ppt-narration"
    }

    Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $ssml -OutFile $outFile
    Write-Host "Created: $outFile"
}

Write-Host "Done."
