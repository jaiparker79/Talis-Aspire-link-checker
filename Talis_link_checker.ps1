# Begin Talis Aspire link checking script
# Function to check if a URL is broken or redirected to a domain
function Test-Url {
    param (
        [string]$url
    )

    # Check for specific URL patterns first.  The ones listed below are QUT-specific, please edit these for your own libraries needs.
    if ($url -match "^https://web\.p\.ebscohost\.com") {
        return "EBSCOhost"
    } elseif ($url -match "^https://www\.clickview\.net") {
        return "ClickView"
    } elseif ($url -match "^https://edu\.digitaltheatreplus\.com") {
        return "Digital Theatre+"
    } elseif ($url -match "^https://learning\.oreilly\.com") {
        return "O'Reilly"
    } elseif ($url -match "^https://viewer\.books24x7\.com") {
        return "Skillsoft"
    } elseif ($url -match "^https://anzlaw\.thomsonreuters\.com") {
        return "Westlaw Australia"
    } elseif ($url -match "^https://1\.next\.westlaw\.com") {
        return "Westlaw International"
    } elseif ($url -match "^https://uk\.westlaw\.com") {
        return "Westlaw UK"
    } elseif ($url -match "^https://ovidsp\.[a-z0-9]+\.ovid\.com") {
        return "Ovid"
    } elseif ($url -match "^https://global-factiva-com\.eu1\.proxy\.openathens\.net") {
        return "Factiva"
    }

    $maxRetries = 1
    $retryCount = 0
    $errorCode = $null

    while ($retryCount -lt $maxRetries -and $errorCode -eq $null) {
        try {
            $response = Invoke-WebRequest -Uri $url -Method Head -TimeoutSec 30 -Headers @{"User-Agent"="Mozilla/5.0"} -MaximumRedirection 5 -ErrorAction Stop

            # Handle redirections
            if ($response.StatusCode -ge 300 -and $response.StatusCode -lt 400) {
                $finalUrl = $response.Headers.Location
                if ($finalUrl -match "^https?://[^/]+/?$") {
                    return "$response.StatusCode - Redirected to domain"
                }
            }

            if ($response.StatusCode -eq 404) {
                return $response.StatusCode
            } elseif ($response.StatusCode -eq 418) {
                return "I'm a teapot"
            }
        } catch {
            if ($_.Exception.Response.StatusCode -eq 404) {
                return $_.Exception.Response.StatusCode
		# these elseif statements which attempt to catch inner exceptions are not working at present. Attempting to work out if PowerShell's Invoke-WebRequest can even do this. Jai 11-05-2025
            } elseif ($_.Exception.InnerException -match "The remote name could not be resolved") {
                return "DNS Lookup Failed"
            } elseif ($_.Exception.InnerException -match "The operation has timed out") {
                return "Timeout"
            } elseif ($_.Exception.InnerException -match "The underlying connection was closed") {
                return "Connection Closed"
            } else {
                $errorCode = $null
            }
        }
        $retryCount++
        Start-Sleep -Seconds 5
    }
    return $errorCode
}

# Function to display menu and get user selection
function Show-Menu {
    param (
        [array]$files
    )
    Write-Host "Select a CSV file to check:"
    for ($i = 0; $i -lt $files.Length; $i++) {
        Write-Host "$($i + 1). $($files[$i].Name)"
    }
	    Write-Host ""  # Blank line
    $selection = Read-Host "Enter the number of the file you want to check"
    return $files[$selection - 1]
}

# Get list of CSV files
$csvFiles = Get-ChildItem -Path . -Filter "all_list_items_*.csv"

if ($csvFiles.Length -eq 0) {
    Write-Host "No CSV files found with the specified pattern." -ForegroundColor Red
    exit
}

    Write-Host "#################################" -ForegroundColor DarkYellow
    Write-Host "Talis Aspire link checking script (beta)" -ForegroundColor DarkYellow
    Write-Host "#################################" -ForegroundColor DarkYellow
    Write-Host ""  # Blank line

# Show menu and get user selection
$inputFilename = Show-Menu -files $csvFiles
$outputFilename = "broken-links-$($inputFilename.BaseName).csv"

try {
    Write-Host ""  # Blank line
    
    if ($inputFilename) {
        Write-Host "Checking $($inputFilename.Name)" -ForegroundColor Magenta
        Write-Host ""  # Blank line
        $csv = Import-Csv -Path $inputFilename.FullName
        $output = @()

        $lineCount = 0
        foreach ($row in $csv) {
            $lineCount++
            Write-Host "Processing line $lineCount"

            # Check URLs in columns AL and O
            $columns = @("Online Resource Web Address", "DOI")
            foreach ($column in $columns) {
                $url = $row.$column
                if ($url) {
                    # Prepend https://doi.org/ if the DOI value starts with 10.
                    if ($url -match "^10\..*") {
                        $url = "https://doi.org/$url"
                    }
                    Write-Host "Checking URL: $url"
                    $errorCode = Test-Url -url $url
                    if ($errorCode -and $errorCode -ne $null) {
                        Write-Host "Broken link detected: $url - Status Code: $errorCode" -ForegroundColor Red
                        $output += [pscustomobject]@{
                            "Item Link"      = $row."Item Link"
                            "HTTP Error Code" = $errorCode
                        }
                        Write-Host ""  # Blank line
                        break
                    } else {
                        Write-Host "URL OK: $url" -ForegroundColor Green
                    }
                    Write-Host ""  # Blank line
                }
            }
            Start-Sleep -Seconds 1  # Add a delay between requests to avoid rate limiting
        }

        # Export the results to a new CSV file
        $output | Export-Csv -Path $outputFilename -NoTypeInformation

        Write-Host "Link checking complete. Please open $outputFilename" -ForegroundColor Green
    } else {
        Write-Host "No CSV file found with the specified pattern."
    }
} catch {
    Write-Host "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
}

# Keep the PowerShell window open
Read-Host -Prompt "Press Enter to exit"
# End QUT Readings link checking script
