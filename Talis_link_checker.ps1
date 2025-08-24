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
    } elseif ($url -match "^https://viewer\.books24x7\.com") {
        return "https://www.library.qut.edu.au/search/status/linking/skillsoft/"
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
    }

# $maxRetries is now set based on user selection below
    $retryCount = 0
    $errorCode = $null

    while ($retryCount -lt $maxRetries -and $errorCode -eq $null) {
        try {
            $response = Invoke-WebRequest -Uri $url -Method Head -TimeoutSec 90 -Headers @{"User-Agent"="Mozilla/5.0"} -MaximumRedirection 5 -ErrorAction Stop
            if ($response.StatusCode -ge 300 -and $response.StatusCode -lt 400) {
                $finalUrl = $response.Headers.Location
                if ($finalUrl -match "^https?://[^/]+/?$") {
                    return "$($response.StatusCode) - Redirected to domain"
                }
            }
            if ($response.StatusCode -eq 400) {
                return $response.StatusCode
            } elseif ($response.StatusCode -eq 404) {
                return $response.StatusCode
            } elseif ($response.StatusCode -eq 418) {
                return "I'm a teapot"
            } elseif ($response.StatusCode -ge 500 -and $response.StatusCode -lt 600) {
                return "Server Error $($response.StatusCode)"
            }
        } catch {
            if ($_.Exception.Response.StatusCode -eq 404) {
                return $_.Exception.Response.StatusCode
            } elseif ($_.Exception -match "The remote name could not be resolved") {
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
if (Test-Path ".\\cancelled.xlsx") {
    try {
        $cancelledData = Import-Excel -Path ".\\cancelled.xlsx"
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

Write-Host "###############################################" -ForegroundColor DarkYellow
Write-Host "Talis Aspire link checking script (Version 1.2)" -ForegroundColor DarkYellow
Write-Host "###############################################" -ForegroundColor DarkYellow
Write-Host ""

$inputFilename = Show-Menu -files $csvFiles
$maxRetries = 1  # Default to full link checking

Write-Host ""
Write-Host "Select link checking mode:"
Write-Host "A. Start of semester full link checking"
Write-Host "B. Roster 3 URL pattern and cancelled item checking"
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
        Write-Host "`nChecking $($inputFilename.Name)" -ForegroundColor Magenta
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
                            "Item Link"       = $row."Item Link"
                            "HTTP Error Code" = $errorCode
                            "Broken URL"      = $url
                        }
                        Write-Host ""
                        break
                    } else {
                        Write-Host "URL OK: $url" -ForegroundColor Green
                    }
                    Write-Host ""  # Blank line
                }
            }
            # Start-Sleep -Seconds 1  # Adds a delay between requests to avoid rate limiting. Uncomment to re-enable.
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
