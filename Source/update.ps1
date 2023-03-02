function getPartial($file, $partial) {
	try {
		$content = Get-Content -Path "$PSScriptRoot\Signatures\$($file.BaseName)_partials\$($partial)$($file.extension)"
		return $content
	} catch {
		return ""
	}
}
function getPartialWithContent($file, $partial, $value) {
	try {
		$content = getPartial $file $partial
		if ($value) {
			$content = $content -replace "%$Partial%", $value
		} else {
			return ""
		}
		return $content
	} catch {
		return ""
	}
}

# Win32 app runs PowerShell in 32-bit by default. AzureAD module requires PowerShell in 64-bit, so we are going to trigger a rerun in 64-bit.
if ($env:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
	try {
		& "$env:WINDIR\SysNative\WindowsPowershell\v1.0\PowerShell.exe" -File $PSCommandPath
	}
	catch {
		throw "Failed to start $PSCommandPath"
	}
	exit
}

Start-Transcript -Path "$($env:TEMP)\IntuneSignatureManagerForOutlook-log.txt" -Force

# Install AzureAD module to retrieve the user information
#Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force

# Leverage Single Sign-on to sign into the AzureAD PowerShell module
$userPrincipalName = whoami -upn
$usersRootDirectory = Resolve-Path -Path "$PSScriptRoot\Users"
$signaturesRootDirectory = Get-ChildItem -Path "$PSScriptRoot\Signatures"

Connect-MgGraph -Scopes 'User.Read.All'
$users = Get-MgUser -All

foreach ($userObject in $users) {

	$signatureDestinationDirectory = Join-Path $usersRootDirectory $userObject.UserPrincipalName
	# Create signatures folder if not exists
	try {	
		if (-not (Test-Path $signatureDestinationDirectory)) {
			New-Item -ItemType Directory -Path $signatureDestinationDirectory
		}
	} catch { }
	

	foreach ($signatureFile in $signaturesRootDirectory) {
		if ($signatureFile.Name -like "*.htm" -or $signatureFile.Name -like "*.rtf" -or $signatureFile.Name -like "*.txt") {
			# Get file content with placeholder values
			$signatureFileContent = Get-Content -Path $signatureFile.FullName
			#Write-Host ($userObject | Format-List -Force | Out-String)
						# Replace placeholder values
			$signatureFileContent = $signatureFileContent -replace "%DisplayName%", $userObject.DisplayName
			$signatureFileContent = $signatureFileContent -replace "%Mail%", $userObject.Mail
			$signatureFileContent = $signatureFileContent -replace "%TelephoneNumber%", $userObject.BusinessPhones[0]
			$signatureFileContent = $signatureFileContent -replace "%JobTitle%", $userObject.JobTitle

			$signatureFileContent = $signatureFileContent -replace "%Mobile%", (getPartialWithContent $signatureFile "Mobile" $userObject.Mobile)
			# $signatureFileContent = $signatureFileContent -replace "%Tradeshows%", (getPartial $signatureFile "Tradeshows")
			$signatureFileContent = $signatureFileContent -replace "%Tradeshows%", ""

			if ($userObject.Department -like "Sales") {
				$partial = getPartial $signatureFile "Department"
				$signatureFileContent = $signatureFileContent -replace "%Department%", $partial
			} else {
				$signatureFileContent = $signatureFileContent -replace "%Department%", ""
			}

			# Set file content with actual values in $env:APPDATA\Microsoft\Signatures
			Set-Content -Path "$signatureDestinationDirectory\$($signatureFile.Name)" -Value $signatureFileContent -Force
		} elseif ($signatureFile.getType().Name -eq 'DirectoryInfo') {
			try {
				Copy-Item -Path $signatureFile.FullName -Destination "$signatureDestinationDirectory" -Recurse -Force
			} catch { }
		}
	}
}

try {
	Get-ChildItem -Path $signaturesRootDirectory -Recurse -Filter "thumbs.db" | Remove-Item
	Get-ChildItem -Path $usersRootDirectory -Recurse -Filter "thumbs.db" | Remove-Item
} catch {}
Stop-Transcript
