function getPartialContent($path, $extension, $partial) {
    try {
        $content = Get-Content -Path "$PSScriptRoot\Signatures\$($path)_partials\$($partial)$($extension)"
        return $content
    }
    catch {
        return ""
    }
}


Start-Transcript -Path "$($env:TEMP)\IntuneSignatureManagerForOutlook-log.txt" -Force

$userPrincipalName = whoami -upn

if (-not (Test-Path "$($env:APPDATA)\Microsoft\Signatures")) {
    $null = New-Item -Path "$($env:APPDATA)\Microsoft\Signatures" -ItemType Directory
}

# Get all signature files
$signatureFiles = Get-ChildItem -Path "$PSScriptRoot\Users\$userPrincipalName"
try {
    Get-ChildItem -Path $signatureFiles -Recurse -Filter "thumbs.db" | Remove-Item
    Get-ChildItem -Path "$($env:APPDATA)\Microsoft\Signatures\" -Recurse -Filter "thumbs.db" | Remove-Item 
}
catch {
}

foreach ($signatureFile in $signatureFiles) {

    try {
        Copy-Item -Path $signatureFile.FullName -Destination "$($env:APPDATA)\Microsoft\Signatures\" -Recurse -Force
    }
    catch {

    }
    
}

Stop-Transcript
