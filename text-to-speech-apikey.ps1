# =============================================================================
# text-to-speech-apikey.ps1
# Azure AI Speech – Text-to-Speech (TTS) using API Key authentication
# =============================================================================
# Usage:
#   1. Set $region and $apiKey below
#   2. Place SSML XML files in the .\ssml\ folder
#   3. Run: .\text-to-speech-apikey.ps1
#   4. WAV files will be created in the .\audio\ folder
# =============================================================================

# ---------- Configuration ----------
$region = "<YOUR-REGION>"       # e.g. "swedencentral", "eastus", "westeurope"
$apiKey = "<YOUR-API-KEY>"      # Azure AI Speech resource key (Key 1 or Key 2)

$inDir  = ".\ssml"
$outDir = ".\audio"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

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
Write-Progress -Activity "Generating audio" -Completed
Write-Host "Done: $total files processed."
