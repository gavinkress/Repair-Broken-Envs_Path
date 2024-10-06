# Repair-Broken-Envs_Path
---
Repair-Broken-Envs_Path

##
SYNOPSIS
---
Repairs and optimizes the system and user PATH environment variables on Windows 11.

##
DESCRIPTION
---
    This script dynamically rebuilds the PATH environment variables, ensuring they are correctly formatted,
    removing duplicates, and incorporating necessary directories. It handles both system and user PATH variables,
    adjusts them based on predefined directory arrays, and substitutes environment variable shortcuts back into paths
    as efficiently as possible to minimize the total length (using the largest fitting subsets).
    
    The script sets various environment variables and provides utility functions for formatting paths,
    handling slashes, and managing environment variable substitutions.

    It dynamically creates environment variables for MSVC paths and identifies the best substrings in the paths
    to create additional environment variables if the PATH length exceeds the limit.
##
PARAMETER None
---
    No parameters are required.
###
NOTES
---
-   - Name: Repair-Envs_Path
    - Homepage: https://github.com/gavinkress/Repair-Broken-Envs_Path
    - Author: Gavin Kress
    - Email: gavinkress@gmail.com
    - Date: 10/05/2024
    - Version: 1.0.0
    - Programming Language(s): PowerShell
    - License: MIT License
    - Operating System: OS Independent
#
```
function Main {
    Clear-Host

    # Clear PowerShell history and perform garbage collection
    Remove-Item (Get-PSReadlineOption).HistorySavePath -ErrorAction SilentlyContinue
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()

    ##########################################################################################################################
    # Input Arrays, Paths, and Variables
    ##########################################################################################################################

    # User PATH additions
    $Path_User_arr = @(
        "%LAppD%/GitHubDesktop/bin",
        "%LAppD%/Microsoft/WindowsApps",
        "%LAppD%/Microsoft/WinGet/Links",
        "%LAppD%/Programs/Microsoft VS Code Insiders/bin",
        "%LAppD%/Programs/Python/Python312",
        "%LAppD%/Programs/Python/Python312/Scripts",
        "%LAppD%/Programs/Python/Python313",
        "%LAppD%/Programs/Python/Python313/Scripts",
        "%LAppD%/Volta/bin",
        "%OneDrive%/Documents",
        "%OneDrive%/Documents/bw-windows-2024.7.2",
        "%OneDrive%/Documents/PowerShell/Modules",
        "%OneDrive%/Documents/SysinternalsSuite",
        "%ProgH%/.env/.virtualenvs/GeneralPythonVenv",
        "%ProgH%/.env/.virtualenvs/GeneralPythonVenv/Scripts",
        "%ProgH%/.env/.virtualenvs/RCyPyVenv/Scripts",
        "%ProgH%/.env/JDK/Oracle_JDK-22",
        "%ProgH%/PowerShell/Scripts",
        "%USERPROFILE%/.cargo/bin",
        "%USERPROFILE%/.dotnet/tools",
        "%USERPROFILE%/.local/bin",
        "%USERPROFILE%/AppData/Roaming/Thunderbird/Profiles/gwiph5m3.default-esr/chrome"
    )

    # System PATH additions (excluding MSVC paths, which will be added dynamically)
    $Path_System_arr = @(
        "/cmd",
        "/usr",
        "/usr/bin",
        "%CUDA_H%/v12.6/bin",
        "%CUDA_H%/v12.6/libnvvp",
        "%ProgramFiles(x86)%/NVIDIA Corporation/PhysX/Common",
        "%ProgramFiles%/dotnet",
        "%ProgramFiles%/Eclipse Adoptium/jdk-21.0.4.7-hotspot/bin",
        "%ProgramFiles%/Git/cmd",
        "%MATLAB_HOME%/bin",
        "%MATLAB_HOME%/runtime/win64",
        "%ProgramFiles%/nodejs",
        "%NVCORP%/Nsight Compute 2024.3.0",
        "%NVCORP%/NVIDIA app/NvDLISR",
        "%ProgramFiles%/PowerShell/7",
        "%ProgramFiles%/R/R-4.4.1/bin",
        "%ProgramFiles%/R/R-4.4.1/bin/x64",
        "%ProgramFiles%/Wolfram Research/WolframScript",
        "%ProgramFiles%/WSL",
        "%R_HOME%",
        "%R_HOME%/bin",
        "%R_HOME%/bin/x64",
        "%RTOOLS40_HOME%/mingw_64/bin",
        "%RTOOLS40_HOME%/ucrt64/bin",
        "%RTOOLS40_HOME%/usr/bin",
        "%RTOOLS40_HOME%/x86_64-w64-mingw32.static.posix/bin",
        "%SystemRoot%",
        "%SystemRoot%/System32",
        "%SystemRoot%/System32/OpenSSH",
        "%SystemRoot%/System32/Wbem",
        "%SystemRoot%/System32/WindowsPowerShell/v1.0",
        "C:/Rtools/bin",
        "C:/Rtools/mingw_64/bin",
        "C:/Users",
        "C:/WINDOWS"
    )

    # Environment Variables Definitions
    $vardefs = @(
        @{ name = "R_PROFILE_USER"; value = "$env:USERPROFILE/OneDrive/Centralized Programming Heirarchy/.env/R/.Rprofile" },
        @{ name = "ProgH"; value = "$env:USERPROFILE/OneDrive/Centralized Programming Heirarchy" },
        @{ name = "CUDA_H"; value = "C:/Program Files/NVIDIA GPU Computing Toolkit/CUDA" },
        @{ name = "R_HOME"; value = "C:/Program Files/R/R-4.4.1" },
        @{ name = "LAppD"; value = "$env:USERPROFILE/AppData/Local" },
        @{ name = "OneDrive"; value = "$env:OneDrive" },
        @{ name = "USERPROFILE"; value = "$env:USERPROFILE" },
        @{ name = "ProgramFiles"; value = "$env:ProgramFiles" },
        @{ name = "ProgramFiles(x86)"; value = "$env:ProgramFiles(x86)" },
        @{ name = "SystemRoot"; value = "$env:SystemRoot" },
        @{ name = "RTOOLS40_HOME"; value = "C:/rtools40" },
        @{ name = "RTOOLS44_HOME"; value = "C:/rtools44" },
        @{ name = "NVCORP"; value = "%ProgramFiles%/NVIDIA Corporation" },
        @{ name = "MATLAB_HOME"; value = "%ProgramFiles%/MATLAB/R2024b" }
    )

    # Define additional paths to add
    $path_adds = @(
        "%ProgramFiles%/R",
        "%R_HOME%",
        "%R_HOME%/bin",
        "%R_HOME%/bin/x64",
        "%RTOOLS40_HOME%/mingw_64/bin",
        "%RTOOLS40_HOME%/usr/bin",
        "%RTOOLS40_HOME%/ucrt64/bin",
        "%RTOOLS40_HOME%/x86_64-w64-mingw32.static.posix/bin",
        "%ProgH%/.env/.virtualenvs/RCyPyVenv/Scripts",
        "%USERPROFILE%/.local/bin"
    )

    # MSVC root path
    $MSVC_Root = "C:/Program Files (x86)/Microsoft Visual Studio/2022/BuildTools/VC/Tools/MSVC"

    # Terms to identify user-scoped paths
    $userTerms = @("%ProgH%", "%LAppD%", "%OneDrive%", "%USERPROFILE%")

    ##########################################################################################################################
    ##########################################################################################################################

    # Function to write output in a formatted way
    function WriteOutputPretty {
        param (
            [string]$currentmsg
        )
        $n = $currentmsg.length
        $outers = "-" * ($n + 92)
        Write-Output ""
        Write-Output $outers
        Write-Output ("-" * 44 + " $currentmsg " + "-" * 44)
        Write-Output $outers
        Write-Output ""
    }

    WriteOutputPretty -currentmsg "Building path additions"

    # Display Input Arrays and Variables
    Write-Output "User PATH Additions:"
    $Path_User_arr | ForEach-Object { Write-Output "  $_" }
    Write-Output ""
    Write-Output "System PATH Additions:"
    $Path_System_arr | ForEach-Object { Write-Output "  $_" }
    Write-Output ""
    Write-Output "Environment Variable Definitions:"
    foreach ($vardef in $vardefs) {
        Write-Output "  $($vardef.name) = $($vardef.value)"
    }
    Write-Output ""
    Write-Output "MSVC Root Path:"
    Write-Output "  $MSVC_Root"
    Write-Output ""

    # Set environment variables in both User and Machine scope
    foreach ($vardef in $vardefs) {
        [System.Environment]::SetEnvironmentVariable($vardef.name, $vardef.value, "User")
        [System.Environment]::SetEnvironmentVariable($vardef.name, $vardef.value, "Machine")
    }

    # Retrieve environment variables for substitution
    $envVars = Get-ChildItem Env: | Sort-Object Name
    $vardefs += $envVars | Where-Object {
        ($_.Value -match "^[a-zA-Z]:\\") -and (-not [string]::IsNullOrEmpty($_.Value))
    } | ForEach-Object {
        @{ name = $_.Name; value = $_.Value }
    }

    # Function to format paths by replacing backslashes and removing duplicates
    function Format-Locations {
        param (
            [Parameter(Mandatory = $true)]
            [Array]$locations
        )

        # Replace backslashes with forward slashes and normalize slashes
        $formattedLocations = $locations | ForEach-Object {
            $loc = $_.Trim()
            $loc = $loc -replace "[\\]+", "/"
            $loc = $loc -replace "[/]+", "/"
            # Remove any trailing slashes
            $loc = $loc.TrimEnd('/')
            $loc
        }
        # Remove duplicates and sort by length descending
        $formattedLocations = $formattedLocations | Sort-Object -Unique | Sort-Object { -$_.Length }
        return $formattedLocations
    }

    # Function to substitute environment variable shortcuts back into paths
    function Resub {
        param (
            [Parameter(Mandatory = $true)]
            [Array]$locations,
            [Parameter(Mandatory = $true)]
            [Array]$vardefs
        )

        # Sort vardefs by value length descending to match largest fitting subset
        $vardefs = $vardefs | Sort-Object { -($_.value.Length) }

        $substitutedLocations = $locations | ForEach-Object {
            $loc = $_
            foreach ($vardef in $vardefs) {
                $value = $vardef.value
                $name = $vardef.name
                # Escape regex special characters in value
                $escapedValue = [regex]::Escape($value)
                # Replace all occurrences of raw paths with environment variable shortcuts
                $pattern = $escapedValue
                if ($loc -match $pattern) {
                    $loc = $loc -replace $pattern, "%$name%"
                }
            }
            $loc
        }
        return $substitutedLocations
    }

    # Retrieve and split existing PATH variables, expanding any environment variables to raw paths
    $Path_User = [System.Environment]::GetEnvironmentVariable("Path", "User")
    $Path_User_arr_native = $Path_User -split ";" | Where-Object { $_ -ne "" }
    $Path_User_arr_native = $Path_User_arr_native | ForEach-Object { [Environment]::ExpandEnvironmentVariables($_) }
    $Path_System = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
    $Path_System_arr_native = $Path_System -split ";" | Where-Object { $_ -ne "" }
    $Path_System_arr_native = $Path_System_arr_native | ForEach-Object { [Environment]::ExpandEnvironmentVariables($_) }

    # Format existing PATH arrays
    $Path_User_arr_native = Format-Locations -locations $Path_User_arr_native
    $Path_System_arr_native = Format-Locations -locations $Path_System_arr_native

    # Combine user-defined PATH additions with existing ones
    $Path_User_arr = $Path_User_arr_native + $Path_User_arr
    $Path_System_arr = $Path_System_arr_native + $Path_System_arr

    # Format combined PATH arrays again
    $Path_User_arr = Format-Locations -locations $Path_User_arr
    $Path_System_arr = Format-Locations -locations $Path_System_arr

    # Build the new combined PATH array
    $currentpath_arr = $Path_User_arr + $Path_System_arr
    $currentpath_arr = Format-Locations -locations $currentpath_arr

    # Substitute environment variable shortcuts back into paths to save space
    $currentpath_arr = Resub -locations $currentpath_arr -vardefs $vardefs

    # Dynamically generate MSVC paths and environment variables
    $MSVC_versions = Get-ChildItem "$MSVC_Root" -Directory | Select-Object -ExpandProperty Name | Sort-Object -Descending

    foreach ($MSVC_version in $MSVC_versions) {
        # Extract unique identifier (first two components)
        $versionParts = $MSVC_version -split '\.'
        $versionNum = $versionParts[0] + $versionParts[1] # e.g., "14" + "42" = "1442"

        # Host/Target architectures
        $archs = @("x86", "x64")
        foreach ($hostArch in $archs) {
            foreach ($targetArch in $archs) {
                $varName = "MSVC${versionNum}H${hostArch}${targetArch}"
                $varValue = "$MSVC_Root/$MSVC_version/bin/Host$hostArch/$targetArch"
                # Add to environment variables
                $vardefs += @{ name = $varName; value = $varValue }
                [System.Environment]::SetEnvironmentVariable($varName, $varValue, "User")
                [System.Environment]::SetEnvironmentVariable($varName, $varValue, "Machine")
                # Add to paths
                $path_adds += "%$varName%"
            }
        }
        # Add bin path
        $binVarName = "MSVC${versionNum}BIN"
        $binVarValue = "$MSVC_Root/$MSVC_version/bin"
        $vardefs += @{ name = $binVarName; value = $binVarValue }
        [System.Environment]::SetEnvironmentVariable($binVarName, $binVarValue, "User")
        [System.Environment]::SetEnvironmentVariable($binVarName, $binVarValue, "Machine")
        $path_adds += "%$binVarName%"
    }

    # Format the additional paths
    $path_adds = Format-Locations -locations $path_adds

    # Combine current paths with additional paths and remove duplicates
    $newpath = $currentpath_arr + $path_adds
    $newpath = Format-Locations -locations $newpath

    # Substitute environment variable shortcuts back into paths to save space
    $newpath = Resub -locations $newpath -vardefs $vardefs

    # Remove duplicates again after substitution
    $newpath = $newpath | Sort-Object -Unique

    # Split paths into user and system based on predefined terms
    $newpath_unique_user = @()
    $newpath_unique_system = @()
    foreach ($loc in $newpath) {
        $isUserPath = $false
        foreach ($term in $userTerms) {
            $escapedTerm = [regex]::Escape($term)
            if ($loc -match "^.*$escapedTerm.*$") {
                $isUserPath = $true
                break # Break inner loop if user term is found
            }
        }
        if ($isUserPath) {
            $newpath_unique_user += $loc
        } else {
            # Retain paths that were explicitly in system scope
            $newpath_unique_system += $loc
        }
    }

    # Remove duplicates and format paths again
    $newpath_unique_user = $newpath_unique_user | Sort-Object -Unique
    $newpath_unique_system = $newpath_unique_system | Sort-Object -Unique

    # Ensure the total paths add up correctly
    $totalPaths = $newpath.Count
    $totalUserPaths = $newpath_unique_user.Count
    $totalSystemPaths = $newpath_unique_system.Count
    if ($totalPaths -ne ($totalUserPaths + $totalSystemPaths)) {
        Write-Output "Error: Total paths do not add up correctly."
        Write-Output "Total Paths: $totalPaths"
        Write-Output "User Paths: $totalUserPaths"
        Write-Output "System Paths: $totalSystemPaths"
    }

    # Check if system PATH exceeds the character limit
    $newpath_str_system = $newpath_unique_system -join ";"
    $newpath_str_user = $newpath_unique_user -join ";"

    $systemPathLimit = 2048
    $userPathLimit = 2048

    if ($newpath_str_system.Length -ge $systemPathLimit) {
        # Optimize by dynamically finding common substrings
        Write-Output "System PATH exceeds the character limit ($systemPathLimit characters). Optimizing..."

        # Find common prefixes among the paths to create new environment variables
        $commonPrefixes = @{}

        foreach ($path in $newpath_unique_system) {
            $splitPath = $path -split "/"
            for ($i = 1; $i -le $splitPath.Count; $i++) {
                $prefix = ($splitPath[0..($i - 1)] -join "/")
                if ($prefix.Length -gt 10) { # Only consider prefixes longer than 10 characters
                    if ($commonPrefixes.ContainsKey($prefix)) {
                        $commonPrefixes[$prefix] += 1
                    } else {
                        $commonPrefixes[$prefix] = 1
                    }
                }
            }
        }

        # Sort prefixes by frequency and length
        $sortedPrefixes = $commonPrefixes.GetEnumerator() | Sort-Object -Property Value -Descending

        # Create environment variables for the most common prefixes
        $counter = 1
        foreach ($prefix in $sortedPrefixes) {
            $varName = "CPATH$counter"
            $varValue = $prefix.Key
            # Avoid conflicts with existing variables
            if (-not ($vardefs | Where-Object { $_.name -eq $varName })) {
                $vardefs += @{ name = $varName; value = $varValue }
                [System.Environment]::SetEnvironmentVariable($varName, $varValue, "User")
                [System.Environment]::SetEnvironmentVariable($varName, $varValue, "Machine")
                $counter += 1
            }
            # Limit the number of new variables to avoid overcomplicating
            if ($counter -gt 5) { break }
        }

        # Update substitutions
        $newpath_unique_system = Resub -locations $newpath_unique_system -vardefs $vardefs

        # Recalculate system PATH string
        $newpath_str_system = $newpath_unique_system -join ";"

        if ($newpath_str_system.Length -ge $systemPathLimit) {
            Write-Output "System PATH still exceeds the character limit after optimization."
            Write-Output "Current System PATH Length: $($newpath_str_system.Length) characters"
            Write-Output "Consider reviewing the PATH entries."
        } else {
            Write-Output "System PATH optimized successfully using additional environment variables."
        }
    }

    # Display PATH stats
    Write-Output "`n`n`n`n`n---------------------------------------------------------------------------"
    Write-Output "User PATH Entries: ($totalUserPaths)"
    $newpath_unique_user | ForEach-Object { Write-Output $_ }
    Write-Output "---------------------------------------------------------------------------"
    Write-Output "System PATH Entries: ($totalSystemPaths)"
    $newpath_unique_system | ForEach-Object { Write-Output $_ }
    Write-Output "---------------------------------------------------------------------------"
    Write-Output "System PATH Length: $($newpath_str_system.Length) characters"
    Write-Output "User PATH Length: $($newpath_str_user.Length) characters"
    Write-Output "Total PATH Entries: $totalPaths"
    Write-Output "---------------------------------------------------------------------------"

    # Check if PATH lengths are within limits before setting them
    if ($newpath_str_system.Length -lt $systemPathLimit -and $newpath_str_user.Length -lt $userPathLimit) {
        [System.Environment]::SetEnvironmentVariable("Path", $newpath_str_user, "User")
        [System.Environment]::SetEnvironmentVariable("Path", $newpath_str_system, "Machine")
        WriteOutputPretty -currentmsg "PATH variables updated successfully"
    } else {
        WriteOutputPretty -currentmsg "ERROR updating PATH variables"
        Write-Output "One or more PATH variables exceed the character limit."
        Write-Output "System PATH Length: $($newpath_str_system.Length) characters"
        Write-Output "User PATH Length: $($newpath_str_user.Length) characters"
        WriteOutputPretty -currentmsg "Consider further optimizing common prefixes or reviewing the PATH entries."
    }
}

Main
```
