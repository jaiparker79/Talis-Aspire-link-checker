# Begin Talis Aspire link checking script

# Function to check if a URL is broken or redirected to a domain

function Test-Url {
    param (
        [string]$url,
        [array]$cancelledItems
    )

    # Check against cancelled items
    foreach ($cancelled in $cancelledItems) {
        if ($url -like "*$cancelled*") {
            return "Cancelled item"
        }
    }

    # Check for specific URL patterns first
    if ($url -match "^https://web\.[a-z0-9]+\.ebscohost\.com") {
        return "https://www.library.qut.edu.au/search/status/linking/#other"
    } elseif ($url -match "^https://www\.clickview\.net/videos/") {
        return "https://www.library.qut.edu.au/search/status/linking/#other"
    } elseif ($url -match "^https://launch\.clickview\.net") {
        return "https://www.library.qut.edu.au/search/status/linking/#other"
    } elseif ($url -match "^https://online\.clickview\.com\.au") {
        return "https://www.library.qut.edu.au/search/status/linking/#other"
    } elseif ($url -match "^https://edu\.digitaltheatreplus\.com") {
        return "https://www.library.qut.edu.au/search/status/linking/digitaltheatre/"
    } elseif ($url -match "^https://learning\.oreilly\.com") {
        return "https://www.library.qut.edu.au/search/status/linking/oreilly/"
    } elseif ($url -match "^https://anzlaw\.thomsonreuters\.com") {
        return "https://www.library.qut.edu.au/search/status/linking/westlaw/"
    } elseif ($url -match "^https://1\.next\.westlaw\.com") {
        return "https://www.library.qut.edu.au/search/status/linking/westlaw/"
    } elseif ($url -match "^https://uk\.westlaw\.com") {
        return "https://www.library.qut.edu.au/search/status/linking/westlaw/"
    } elseif ($url -match "^https://ovidsp\.[a-z0-9]+\.ovid\.com") {
        return "Ovid. Check metadata has valid DOI and switch linking to OpenURL"
    } elseif ($url -match "^https://global-factiva-com\.eu1\.proxy\.openathens\.net") {
        return "https://www.library.qut.edu.au/search/status/linking/factiva/"
    } elseif ($url -match "^https://dj-factiva-com\.eu1\.proxy\.openathens\.net") {
        return "https://www.library.qut.edu.au/search/status/linking/factiva/"
    } elseif ($url -match "^https://qut\.eblib\.com") {
        return "EBL. Change URL to ProQuest Ebook Central"
    } elseif ($url -match "^https://qut\.eblib\.com\.au") {
        return "EBL. Change URL to ProQuest Ebook Central"
    } elseif ($url -match "eu1\.proxy\.openathens\.net") {
        return "OpenAthens Proxied"
    } elseif ($url -match "ezp01\.library\.qut\.edu\.au") {
        return "EZproxy"
    } elseif ($url -match "gateway\.library\.qut\.edu\.au") {
        return "EZproxy"
    } elseif ($url -match "c=UERG") {
        return "Ebook Central PDF"
    } elseif ($url -match "^https://iview\.abc\.net\.au") {
        return "ABC iView, replace with copy from ClickView or EduTV"
    } elseif ($url -match "^https://www\.sbs\.com\.au/ondemand/") {
        return "SBS On Demand, replace with copy from ClickView or EduTV"
    } elseif ($url -match "bloomsburycollections\.com") {
        return "Bloomsbury, switch to OpenURL"
    } elseif ($url -match "bloomsburyfashoncentral\.com") {
        return "Bloomsbury, switch to OpenURL"
    } elseif ($url -match "bloomsburymusicandsound\.com") {
        return "Bloomsbury, switch to OpenURL"
    } elseif ($url -match "bloomsburyvideolibrary\.com") {
        return "Bloomsbury, switch to OpenURL"
    } elseif ($url -match "bloomsburyvisualarts\.com") {
        return "Bloomsbury, switch to OpenURL"
    } elseif ($url -match "dramaonlinelibrary\.com") {
        return "Bloomsbury, switch to OpenURL"
    } elseif ($url -match "screenstudies\.com") {
        return "Bloomsbury, switch to OpenURL"
    } elseif ($url -match "storyboxhub\.com/stories.*") {
        return "https://www.library.qut.edu.au/search/status/linking/storyboxhub/"
    }

    # $maxRetries is now set based on user selection below
    $retryCount = 0
    $errorCode = $null

    while ($retryCount -lt $maxRetries -and $errorCode -eq $null) {
        try {
            # -MaximumRedirection 0 is unreliable across PowerShell/.NET versions
            # (some treat 0 as "unlimited" rather than "stop before following").
            # -MaximumRedirection 1 reliably follows exactly one hop, which is enough
            # to inspect whether that hop was a 301 to a bare domain.
            $response = Invoke-WebRequest -UseBasicParsing -Uri $url -Method Head -TimeoutSec 90 `
                -Headers @{"User-Agent" = "Mozilla/5.0"} -MaximumRedirection 1 -ErrorAction Stop

            # ----- STEP 1: 301 redirected to a bare domain -----
            # No exception was thrown, meaning the redirect (if any) was followed
            # and landed on a working (2xx) page. Detect whether a redirect actually
            # happened by comparing host+path (trailing slash/case insensitive),
            # then flag only if it landed on a bare top-level domain.
            $requestedUri = [System.Uri]$url
            $landedUri = $response.BaseResponse.ResponseUri
            $redirectOccurred = ($requestedUri.Host -ne $landedUri.Host) -or
                                 ($requestedUri.AbsolutePath.TrimEnd('/') -ne $landedUri.AbsolutePath.TrimEnd('/'))

            if ($redirectOccurred -and $landedUri.AbsoluteUri -match "^https?://[^/]+/?$") {
                return "301 - Redirected to domain"
            }

            # ----- STEP 2: 404, 405, 500 -----
            if ($response.StatusCode -eq 404) {
                return $response.StatusCode
            } elseif ($response.StatusCode -eq 405) {
                return $response.StatusCode
            } elseif ($response.StatusCode -eq 500) {
                return "Server Error $($response.StatusCode)"
            }

        } catch {
            $statusCode = $null
            if ($_.Exception.Response -ne $null) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            # ----- STEP 1: 301 redirected to a bare domain -----
            # A thrown 3xx here means a second redirect occurred beyond the one hop
            # we allowed; PowerShell surfaces that as an exception.
            if ($statusCode -eq 301) {
                $locationHeader = $_.Exception.Response.Headers["Location"]
                $finalUrl = [string]$locationHeader

                if ($finalUrl -match "^https?://[^/]+/?$") {
                    return "301 - Redirected to domain"
                }
                # 301 to a specific path, not a bare domain - not flagged per Step 1.
            }
            # ----- STEP 2: 400, 404, 418, 500 -----
            elseif ($statusCode -eq 400) {
                return $statusCode
            } elseif ($statusCode -eq 404) {
                return $statusCode
            } elseif ($statusCode -eq 405) {
                return $statusCode
            } elseif ($statusCode -eq 500) {
                return "Server Error $statusCode"
            }
            # ----- STEP 2: named connection/DNS errors -----
            elseif ($_.Exception -match "The remote name could not be resolved") {
                return "DNS Lookup Failed"
            } elseif ($_.Exception -match "The operation has timed out") {
                return "Timeout"
            } elseif ($_.Exception -match "The underlying connection was closed") {
                return "Connection Closed"
            } elseif ($_.Exception -match "NXDOMAIN") {
                return "NXDOMAIN Error"
            } else {
                $errorCode = $null
            }
        }

        $retryCount++
        Start-Sleep -Seconds 1
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
    Write-Host ""
    $selection = Read-Host "Enter the number of the file you want to check"
    return $files[$selection - 1]
}

# Load cancelled items from Excel
$cancelledItems = @()
if (Test-Path ".\cancelled.xlsx") {
    try {
        $cancelledData = Import-Excel -Path ".\cancelled.xlsx"
        foreach ($row in $cancelledData) {
            $cancelledItems += $row.PSObject.Properties.Value
        }
        $cancelledItems = $cancelledItems | Where-Object { $_ -ne $null -and $_ -ne "" }
    } catch {
        Write-Host "Error loading cancelled.xlsx: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "" # Blank line
    Write-Host "cancelled.xlsx not found in the current directory." -ForegroundColor Yellow
    Write-Host "" # Blank line
}

# Get list of CSV files
$csvFiles = Get-ChildItem -Path . -Filter "all_list_items_*.csv"

if ($csvFiles.Length -eq 0) {
    Write-Host "No CSV files found with the specified pattern." -ForegroundColor Red
    exit
}

Write-Host ""
Write-Host "##########################################################################" -ForegroundColor DarkYellow
Write-Host "Talis Aspire link checking script with cancellation checking (Version 1.3)" -ForegroundColor DarkYellow
Write-Host "##########################################################################" -ForegroundColor DarkYellow
Write-Host ""

$inputFilename = Show-Menu -files $csvFiles

$maxRetries = 1 # Default to full link checking

Write-Host ""
Write-Host "Select link checking mode:"
Write-Host "A. Full link and DOI checking"
Write-Host "B. URL pattern and cancelled item checking only"
Write-Host ""

$modeSelection = Read-Host "Enter your choice (A or B)"

if ($modeSelection -eq 'B') {
    $maxRetries = 0
} elseif ($modeSelection -eq 'A') {
    $maxRetries = 1
} else {
    Write-Host "Invalid selection. Defaulting to Full link checking." -ForegroundColor Red
    $maxRetries = 1
}

Write-Host ""

$outputFilename = "broken-links-$($inputFilename.BaseName).csv"

try {
    if ($inputFilename) {
        Write-Host "`nChecking $($inputFilename.Name)" -ForegroundColor Green
        Write-Host "" #Blank line

        $csv = Import-Csv -Path $inputFilename.FullName
        $output = @()
        $lineCount = 0

        foreach ($row in $csv) {
            $lineCount++
            Write-Host "Processing line $lineCount"

            $columns = @("Online Resource Web Address", "DOI")

            foreach ($column in $columns) {
                $url = $row.$column

                if ($url) {
                    if ($url -match "^10\..*") {
                        $url = "https://doi.org/$url"
                    }

                    Write-Host "Checking URL: $url"
                    $errorCode = Test-Url -url $url -cancelledItems $cancelledItems

                    if ($errorCode) {
                        Write-Host "Errant link detected: $url - Status Code: $errorCode" -ForegroundColor Red

                        $output += [pscustomobject]@{
                            "Title"                       = $row."Title"
                            "Chapter/Article Title"       = $row."Chapter/Article Title"
                            "Item Link"                   = $row."Item Link"
                            "List Appearance"             = $row."List Appearance"
                            "Time Period"                 = $row."Time Period"
                            "List Link"                   = $row."List Link"
                            "Library Note"                = $row."Library Note"
                            "Error message/instructions"  = $errorCode
                            "Broken URL"                  = $url
                        }

                        Write-Host ""
                        break
                    } else {
                        Write-Host "URL OK: $url" -ForegroundColor Green
                    }

                    Write-Host "" # Blank line
                }
            }

            # Start-Sleep -Seconds 1 # Adds a delay between requests to avoid rate limiting. Uncomment to re-enable.
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

# End Talis Aspire link checking script
