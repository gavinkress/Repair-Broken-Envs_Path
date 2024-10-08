
<#
.SYNOPSIS
    Repairs and optimizes the system and user PATH environment variables on Windows 11.

.DESCRIPTION
    This script dynamically rebuilds the PATH environment variables, ensuring they are correctly formatted,
    removing duplicates, and incorporating necessary directories. It handles both system and user PATH variables,
    adjusts them based on predefined directory arrays, and substitutes environment variable shortcuts back into paths
    as efficiently as possible to minimize the total length.

    The script sets various environment variables and provides utility functions for formatting paths,
    handling slashes, and managing environment variable substitutions.

.PARAMETER None
    No parameters are required.

.NOTES
    - Name: Repair-Envs_Path
    - Homepage: https://github.com/gavinkress/Repair-Broken-Envs_Path
    - Author: Gavin Kress
    - Email: gavinkress@gmail.com
    - Date: 10/06/2024
    - Version: 1.4.0
    - Programming Language(s): PowerShell
    - License: MIT License
    - Operating System: Windows 11
#>

# Clear console and perform garbage collection
Clear-Host
Remove-Item (Get-PSReadlineOption).HistorySavePath -ErrorAction SilentlyContinue
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[System.GC]::Collect()

##########################################################################################################################
# Input Arrays, Paths, and Variables
##########################################################################################################################

# User PATH additions
$Path_User_arr = @(
    "%AppLocal%/GitHubDesktop/bin/",
    "%AppLocal%/Microsoft/WindowsApps/",
    "%AppLocal%/Microsoft/WinGet/Links/",
    "%AppLocal%/Programs/Microsoft VS Code Insiders/bin/",
    "%AppLocal%/Programs/Python/Python312/",
    "%AppLocal%/Programs/Python/Python312/Scripts/",
    "%AppLocal%/Programs/Python/Python313/",
    "%AppLocal%/Programs/Python/Python313/Scripts/",
    "%AppLocal%/Volta/bin/",
    "%OneDrive%/Documents/",
    "%OneDrive%/Documents/bw-windows-2024.7.2/",
    "%OneDrive%/Documents/PowerShell/Modules/",
    "%OneDrive%/Documents/SysinternalsSuite/",
    "%PH%/.env/.virtualenvs/GeneralPythonVenv/",
    "%PH%/.env/.virtualenvs/GeneralPythonVenv/Scripts/",
    "%PH%/.env/.virtualenvs/RCyPyVenv/Scripts/",
    "%PH%/.env/JDK/Oracle_JDK-22/",
    "%PH%/PowerShell/Scripts/",
    "%USERPROFILE%/.cargo/bin/",
    "%USERPROFILE%/.dotnet/tools/",
    "%USERPROFILE%/.local/bin/",
    "%USERPROFILE%/AppData/Roaming/Thunderbird/Profiles/gwiph5m3.default-esr/chrome/"
)

# System PATH additions (excluding MSVC paths, which will be added dynamically)
$Path_System_arr = @(

    "%CUDA_H%/v12.6/bin/",
    "%CUDA_H%/v12.6/libnvvp/",
    "%PF86%/NVIDIA Corporation/PhysX/Common/",
    "%PF86%/Windows Kits/10/bin/10.0.26100.0/x64/",
    "%PF%/dotnet/",
    "%PF%/Eclipse Adoptium/jdk-21.0.4.7-hotspot/bin/",
    "%PF%/Git/cmd/",
    "%MATLAB_HOME%/bin/",
    "%MATLAB_HOME%/runtime/win64/",
    "%PF%/nodejs/",
    "%NVCORP%/Nsight Compute 2024.3.0/",
    "%NVCORP%/NVIDIA app/NvDLISR/",
    "%PF%/PowerShell/7/",
    "%PF%/Wolfram Research/WolframScript/",
    "%PF%/WSL/",
    "%SystemRoot%/",
    "%SystemRoot%/System32/",
    "%SystemRoot%/System32/OpenSSH/",
    "%SystemRoot%/System32/Wbem/",
    "%SystemRoot%/System32/WindowsPowerShell/v1.0/",
    "C:/Users/",
    "C:/WINDOWS/"
)

# Environment Variables Definitions
$vardefs = @(
    @{ name = "R_PROFILE_USER"; value = "$env:USERPROFILE/OneDrive/Centralized Programming Heirarchy/.env/R/.Rprofile/" },
    @{ name = "PH"; value = "$env:USERPROFILE/OneDrive/Centralized Programming Heirarchy/" },
    @{ name = "CUDA_H"; value = "C:/Program Files/NVIDIA GPU Computing Toolkit/CUDA/" },
    @{ name = "AppLocal"; value = "$env:USERPROFILE/AppData/Local/" },
    @{ name = "OneDrive"; value = "$env:OneDrive/" },
    @{ name = "PF"; value = "C:/Program Files/" },
    @{ name = "PF86"; value = "C:/Program Files (x86)/" },
    @{ name = "RTOOLS40_HOME"; value = "C:/rtools40/" },
    @{ name = "RTOOLS44_HOME"; value = "C:/rtools44/" },
    @{ name = "NVCORP"; value = "C:/Program Files/NVIDIA Corporation/" },
    @{ name = "MATLAB_HOME"; value = "C:/Program Files/MATLAB/R2024b/" }
)

# Define additional paths to add
$path_adds = @(
    "%PF%/R",
    "%R_HOME%",
    "%R_HOME%/bin",
    "%R_HOME%/bin/x64",
    "%RTOOLS40_HOME%/mingw_64/bin",
    "%RTOOLS40_HOME%/usr/bin",
    "%RTOOLS40_HOME%/ucrt64/bin",
    "%RTOOLS40_HOME%/x86_64-w64-mingw32.static.posix/bin",
    "%PH%/.env/.virtualenvs/RCyPyVenv/Scripts",
    "%USERPROFILE%/.local/bin"
)



# MSVC root path
$MSVC_Root = "C:/Program Files (x86)/Microsoft Visual Studio/2022/BuildTools/VC/Tools/MSVC/"

# Terms to identify user-scoped paths
$userTerms = @("%PH%", "%AppLocal%", "%OneDrive%", "%USERPROFILE%")

##########################################################################################################################
# Functions
##########################################################################################################################

# Function to write output in a formatted way
function WriteOutputPretty {
    param (
        $currentmsg, # Message to display can be any string, array, or object
        $Color = "Rainbow" # Color of the message (default: Rainbow)
    )
    $RainbowColors = @("DarkRed", "Red", "DarkRed", "DarkYellow", "Yellow", "DarkYellow", "DarkGreen", "Green", "DarkGreen", "DarkCyan", "Cyan", "DarkCyan", "DarkBlue", "DarkBlue", "Blue", "DarkMagenta", "Magenta", "DarkMagenta")
    $width = (Get-Host).UI.RawUI.MaxWindowSize.Width
    if (($currentmsg.GetType().BaseType -in @("System.Array")) -or ($currentmsg.GetType().Name -in @("ArrayList", "List", "Array", "Object[]"))) {
        if ($currentmsg[0].GetType().Name -in @('Hashtable', 'PSCustomObject')) {
            $ukeys = @($currentmsg.Keys | Sort-Object -Unique)
            $sections = @()
            
            foreach ($ikey in $ukeys) {
                $sections+=@{name=$ikey; sectwidth=($currentmsg| ForEach-Object { $_.$ikey.Length } | Measure-Object -Maximum).Maximum+8}
            }
            if ($width -gt ((8*$sections.count)+6)){ 
                $maxsectionwidth = [math]::floor(($width-6)/$sections.count)
                if (($sections.sectwidth | Measure-Object -Sum).Sum -lt ($width-6)){$sw_cond = $true}
                $rows = @()
                $rows+=(@($sections | ForEach-Object {
                    if($sw_cond){$sectionwidth = $_.sectwidth}
                    else{
                        sectionwidth = (@($_.sectwidth, $maxsectionwidth)| Measure-Object -Minimum).Minimum
                    }
                    $maxnamewidth = (@($_.name.Length, ($sectionwidth - 12) )| Measure-Object -Minimum).Minimum
                    $subname = $_.name.substring(0, $maxnamewidth)
                    $inside = "- "+$subname.ToUpper()+" -"
                    $Pad = ($sectionwidth - $inside.Length-8)
                    $PadLeft = [math]::floor($Pad/2)
                    $PadRight = $Pad - $PadLeft
                    $sectiondata = "  ||{0}{1}{2}||  " -f (" "*$PadLeft), $inside, (" "*$PadRight)
                    $sectiondata
                }) -join "")
                $rows+=(@($sections | ForEach-Object {
                    if($sw_cond){$sectionwidth = $_.sectwidth}
                    else{
                        sectionwidth = (@($_.sectwidth, $maxsectionwidth)| Measure-Object -Minimum).Minimum
                    }
                    $maxnamewidth = $sectionwidth - 8
                    $subname = "-"*$maxnamewidth
                    $inside = "--"+$subname+"--"
                    $Pad = ($sectionwidth - $inside.Length-4)
                    $PadLeft = [math]::floor($Pad/2)
                    $PadRight = $Pad - $PadLeft
                    $sectiondata = "  {0}{1}{2}  " -f ("-"*$PadLeft), $inside, ("-"*$PadRight)
                    $sectiondata
                }) -join "")
                foreach ($currentmsgitem in $currentmsg){
                    $rows+=(@($sections | ForEach-Object {
                        if($sw_cond){$sectionwidth = $_.sectwidth}
                        else{
                            sectionwidth = (@($_.sectwidth, $maxsectionwidth)| Measure-Object -Minimum).Minimum
                        }
                        
                        $maxnamewidth = (@($currentmsgitem.($_.name).Length, ($sectionwidth - 12) )| Measure-Object -Minimum).Minimum
                        $subname = $currentmsgitem.($_.name).substring(0, $maxnamewidth)
                        $inside = " "+$subname+" "
                        $Pad = ($sectionwidth - $inside.Length-10)
                        $PadLeft = [math]::floor($Pad/2)
                        $PadRight = $Pad - $PadLeft
                        $sectiondata = "  ||-{0}{1}{2}-||  " -f (" "*$PadLeft), $inside, (" "*$PadRight)
                        $sectiondata
                    }) -join "")
                }
                $currentmsgarr = $rows
            } else {
                
                $currentmsgarr =  @("")

            }
            
        } else {
            $currentmsgarr = $currentmsg
        } 
    } else {
        $currentmsgarr = $currentmsg -split "`n"
    }
     
    if ($Color -eq "Rainbow") {
        $Color_index = (Get-Random -Minimum 0 -Maximum $RainbowColors.Count)
        $Color = $RainbowColors[$Color_index]
    }
    
    
    $outers = "-" * $width
    Write-Host ""
    Write-Host $outers -ForegroundColor $Color
    
    if ($null -ne $Color_index) {
    $Color_index++
    $Color_index = $Color_index % $RainbowColors.Count
    $Color = $RainbowColors[$Color_index]
    }
    
    foreach ($currentmsgit in $currentmsgarr){   
        $n = (@(($width - ($currentmsgit.Length + 6)), 0) | Measure-Object -Maximum).Maximum
        $nh = [math]::Floor($n / 2)
        $sides = "-"*$nh
        if ($n % 2 -eq 0) {$chrext = ""} else {$chrext = "-"}
        Write-Host "| $sides $currentmsgit $sides$chrext |" -ForegroundColor $Color
        
        if ($null -ne $Color_index) {
            $Color_index++
            $Color_index = $Color_index % $RainbowColors.Count
            $Color = $RainbowColors[$Color_index]
        }
    }
    Write-Host $outers -ForegroundColor $Color
    Write-Host ""
}

# Function to format paths by replacing backslashes, ensuring consistent slashes, and removing duplicates
function Format-Locations {
    param (
        [Parameter(Mandatory = $true)]
        [Array]$locations
    )
    if ($locations[0].GetType().Name -eq "String") {
        $formattedLocations = $locations | ForEach-Object {
            $loc = $_.Trim()
            $loc = [Environment]::ExpandEnvironmentVariables($loc)
            $loc = $loc.Trim()
            $loc = $loc -replace "[\\]+", "/"
            $loc = $loc -replace "[/]+", "/"
            # Ensure each path ends with a single '/'
            if (-not $loc.EndsWith("/")) {
                $loc = $loc + "/"
            }
            $loc
        }
        # Remove duplicates
        $formattedLocations = $formattedLocations | Sort-Object -Unique
        return $formattedLocations
    } elseif ($locations[0].GetType().Name -in @("PSCustomObject", "Hashtable")) {
        if ("name" -notin $locations[0].Keys -or "value" -notin $locations[0].Keys) {
            WriteOutputPretty -currentmsg "Error: Invalid format for locations array." -Color Red
            WriteOutputPretty -currentmsg "Object must contian name and value keys" -Color Red
            exit
        }
        $formattedLocations = $locations | ForEach-Object {
            $loc = $_.value
            $loc = $loc.Trim()
            $loc = [Environment]::ExpandEnvironmentVariables($loc)
            $loc = $loc.Trim()
            $loc = $loc -replace "[\\]+", "/"
            $loc = $loc -replace "[/]+", "/"
            # Ensure each path ends with a single '/'
            if (-not $loc.EndsWith("/")) {
                $loc = $loc + "/"
            }
            @{ name = $_.name; value = $value }
        }
        return $formattedLocations
    } else {
        WriteOutputPretty -currentmsg "Error: Invalid format for locations array." -Color Red
        WriteOutputPretty -currentmsg "Array must contain strings, Hashtables, or PSCustomObjects" -Color Red
        exit
    }
}

# Function to substitute environment variable shortcuts back into paths
function Resub {
    param (
        [Parameter(Mandatory = $true)]
        [Array]$locations,
        [Parameter(Mandatory = $true)]
        [Array]$vardefs
    )

    # Normalize vardef values
    $vardefsFormatted = Format-Locations -locations $vardefs

    # Sort vardefs by value length descending
    $vardefsFormatted = $vardefsFormatted | Sort-Object { -($_.value.Length) }

    $substitutedLocations = $locations | ForEach-Object {
        $loc = [Environment]::ExpandEnvironmentVariables($_)
        $loc = $loc -replace "[\\]+", "/"
        $loc = $loc -replace "[/]+", "/"
        $loc = $loc.TrimEnd('/')
        
        foreach ($vardef in $vardefsFormatted) {
            $value = $vardef.value
            $name = $vardef.name
            $escapedValue = [regex]::Escape($value)
            
            if ($loc -match $escapedValue) {
                
                #Robust replacement strategy
                $loc = $loc -replace $escapedValue, "%"+$name+"%"
                $loc = $loc -ireplace $escapedValue, "%"+$name+"%"
                
                $loc = $loc -replace $value, "%"+$name+"%"
                $loc = $loc -ireplace $value, "%"+$name+"%"
                
                $loct = @($loc -split $escapedValue)
                if ($loct.Count -eq 2) {
                    $loc = "{0}/{1}/{2}" -f $loct[0], "%"+$name+"%", $loct[1] 
                }
                $loct = @($loc -split $escapedValue)
                if ($loct.Count -eq 2) {
                $loc = "{0}/{1}/{2}" -f $loct[0], "%"+$name+"%", $loct[1] 
                }
            }
        }
        
        $loc
    }
    return $substitutedLocations
}

# Function to set environment variables efficiently
function Set-EnvironmentVariables {
    param (
        [Parameter(Mandatory = $true)]
        [Array]$vardefs,
        [bool]$overwrite = $false
    )

    $totalEnvVars = $vardefs.Count
    $currentVarIndex = 0

    foreach ($vardef in $vardefs) {
        $currentVarIndex++
        $progressPercent = ($currentVarIndex / $totalEnvVars) * 100
        Write-Progress -Activity "Setting Environment Variables" -Status "Processing $currentVarIndex of $totalEnvVars" -PercentComplete $progressPercent
        
        $cname = $vardef.name
        $cvalue = $vardef.value
        
        $existingValueUser = $null
        $existingValueMachine = $null
        $existingValueUser = [System.Environment]::GetEnvironmentVariable($cname, "User")
        $existingValueMachine = [System.Environment]::GetEnvironmentVariable($cname, "Machine")

        if ((($existingValueUser -notin @($null, $cvalue)) -or ($existingValueMachine -notin @($null, $cvalue))) -and -not $overwrite) {
            # Variable exists with a different value; output warning
            WriteOutputPretty -currentmsg "CRITICAL FAIL: Environment variable '$($envvar.Name)' exists with a different value." -color Red
            WriteOutputPretty -currentmsg "Existing Value: $($envvar.Value)" -color Red
            WriteOutputPretty -currentmsg "Rename it in `\$vardefs." -color Red
            $exitcode = 1
        elseif (($existingValueUser -eq $cvalue) -or ($existingValueMachine -eq $cvalue)) {            }
        } else {
            
            [System.Environment]::SetEnvironmentVariable($cname, $cvalue, "User")
            [System.Environment]::SetEnvironmentVariable($cname, $cvalue, "Machine")
        }
        # If variable exists with the same value, do nothing
    }
    if ($exitcode -ne 0) {exit $exitcode}
}

# Function to get the latest version of software installed in a directory
function Get-LatestVersionPath {
    param (
        [Parameter(Mandatory = $true)]
        [string]$BasePath,
        [Parameter(Mandatory = $true)]
        [string]$Pattern
    )

    $directories = Get-ChildItem -Path $BasePath -Directory -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -match $Pattern
    } | ForEach-Object {
        $versionString = $_.Name -replace '[^\d\.]', ''
        if ([Version]::TryParse($versionString, [ref]$null)) {
            [PSCustomObject]@{
                Name    = $_.Name
                Version = [Version]$versionString
                Path    = $_.FullName
            }
        }
    } | Sort-Object -Property Version -Descending

    return $directories
}



# Main function to encapsulate the script logic
function Main {
    
    
    param([switch]$Elevated)

    function Test-Admin {
        $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
        $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    }

    if ((Test-Admin) -eq $false)  {
        if ($elevated) {
            # tried to elevate, did not work, aborting
        } else {
            Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
        }
        exit
    }

    WriteOutputPretty -currentmsg "Starting PATH configuration script" -Color Cyan
    $exitcode = 0
    
    # Format predefined paths
    $vardefs = Format-Locations -locations $vardefs
    
    # Format existing environmant variable paths
    $envVarsRaw = Get-ChildItem Env: | Sort-Object Name | Where-Object {
        $_.Value -match ".+"
    } | ForEach-Object {
        @{ name = $_.Name; value = $_.Value }
    }
    $envVarsRaw_Paths = $envVarsRaw | Where-Object {
        ($_.Value -match "(?(^.*:.*$)^(?!.*;.*)[a-zA-Z0-9]{1,11}:+(/|\\)+(?!.*:.*).*$|^(?!.*;.*$).*(/|\\)+.*$)") 
    } | ForEach-Object {
        @{ name = $_.Name; value = $_.Value }
    }

    # Save to envs
    $envVarsFormatted = Format-Locations -locations $envVarsRaw_Paths
    Set-EnvironmentVariables -vardefs $envVarsFormatted -overwrite $true
    
    # Redefine with formatted paths
    $envVars = Get-ChildItem Env: | Sort-Object Name | Where-Object {
        $_.Value -match ".+"
    } | ForEach-Object {
        @{ name = $_.Name; value = $_.Value }
    }
    $envVars_Paths = $envVars | Where-Object {
        ($_.Value -match "(?(^.*:.*$)^(?!.*;.*)[a-zA-Z0-9]{1,11}:+(/|\\)+(?!.*:.*).*$|^(?!.*;.*$).*(/|\\)+.*$)") 
    } | ForEach-Object {
        @{ name = $_.Name; value = $_.Value }
    }
    
    # Ensure you arent overwriting things
    foreach ($envvar in $envVars){
        
        if ($envvar.Name -in $vardefs.name){
            $match = $vardefs | Where-Object { $_.name -eq $envvar.Name }
            
            if ($envvar.value -ne $match.value){
                WriteOutputPretty -currentmsg "CRITICAL FAIL: Environment variable '$($envvar.name)' exists with a different value." -color Red
                WriteOutputPretty -currentmsg "Existing Value: $($envvar.value)" -color Red
                WriteOutputPretty -currentmsg "Rename it in `\$vardefs." -color Red
                $ExitCode = 1
            }
        }
    }
    if ($exitcode -ne 0) {exit $exitcode}
    

    # Set environment variables
    Set-EnvironmentVariables -vardefs $vardefs 
    $vardefs += $envvars_paths
    
    
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
    $Path_User_arr = Format-Locations -locations $Path_User_arr
    $Path_System_arr = Format-Locations -locations $Path_System_arr

    # Combine user-defined PATH additions with existing ones
    $Path_User_arr = $Path_User_arr_native + $Path_User_arr
    $Path_System_arr = $Path_System_arr_native + $Path_System_arr

    # Format combined PATH arrays again
    $Path_User_arr = Format-Locations -locations $Path_User_arr
    $Path_System_arr = Format-Locations -locations $Path_System_arr

    # Build the new combined PATH array
    $currentpath_arr = $Path_User_arr + $Path_System_arr
    $currentpath_arr = Format-Locations -locations $currentpath_arr

    # Dynamically generate MSVC paths and environment variables
    WriteOutputPretty -currentmsg "Generating MSVC Paths" -Color Cyan
    
    # Get MSVC versions and sort them numerically in descending order
    $MSVC_versions = Get-ChildItem $MSVC_Root | Where-Object {
        $_.Name -match  "^[a-zA-Z]{0,10}[ _\-\.\d]+[a-zA-Z]{0,10}$"
    } | ForEach-Object {
        [PSCustomObject]@{
            Name    = $_.Name
            Version = [Version]$_.Name
            Path    = $_.FullName
        }
    } | Sort-Object -Property Version -Descending

    if (-not $MSVC_versions) {
        Write-Host "No MSVC versions found in $MSVC_Root" -ForegroundColor Yellow
    } else {
        Write-Host "Found MSVC versions:" -ForegroundColor Green
        $MSVC_versions | ForEach-Object { Write-Host "  $($_.Name)" -ForegroundColor Green }
    }

    foreach ($MSVC_version in $MSVC_versions) {
        # Extract unique identifier (e.g., "1428" from "14.28.29333")
        $versionParts = $MSVC_version.Version
        $versionNum = $versionParts.Major.ToString() + $versionParts.Minor.ToString() # e.g., "14" + "28" = "1428"

        # Host/Target architectures
        $archs = @("x86", "x64")
        foreach ($hostArch in $archs) {
            foreach ($targetArch in $archs) {
                $varName = "MSVC${versionNum}H${hostArch}${targetArch}"
                $varValue = "$MSVC_version.Path/bin/Host$hostArch/$targetArch/"
                $varValue = $varValue -replace "[\\]+", "/"
                $varValue = $varValue -replace "[/]+", "/"
                $varValue = Format-Locations @( $varValue )
                $varValue = $varValue[0]
                $vardefs += @{ name = $varName; value = $varValue }
                $path_adds += "%"+$varName+"%"
            }
        }
        # Add bin path
        $binVarName = "MSVC${versionNum}BIN"
        $binVarValue = "$MSVC_version.Path/bin/"
        $binVarValue = $binVarValue -replace "[\\]+", "/"
        $binVarValue = $binVarValue -replace "[/]+", "/"
        $varValue = Format-Locations @( $varValue )
        $varValue = $varValue[0]
        $vardefs += @{ name = $binVarName; value = $binVarValue }
        $path_adds += "%"+$binVarName+"%"
    }
    
    # Set Environment Variables with Versioned MSVC Paths
    Set-EnvironmentVariables -vardefs $vardefs

    # Combine current paths with additional paths and remove duplicates
    $path_adds = Format-Locations -locations $path_adds
    $newpath = $currentpath_arr + $path_adds
    $newpath = Format-Locations -locations $newpath
    $newpath = $newpath | Sort-Object -Unique

    # Substitute environment variable shortcuts back into paths to save space
    $newpath = Resub -locations $newpath_expanded -vardefs $vardefs

    # Split paths into user and system based on predefined terms
    $newpath_unique_user = @()
    $newpath_unique_system = @()

    Write-Progress -Activity "Classifying Paths" -Status "Processing..." -PercentComplete 0

    $totalPathsToProcess = $newpath.Count
    $currentPathIndex = 0

    foreach ($loc in $newpath) {
        $currentPathIndex++
        $progressPercent = ($currentPathIndex / $totalPathsToProcess) * 100
        Write-Progress -Activity "Classifying Paths" -Status "Processing path $currentPathIndex of $totalPathsToProcess" -PercentComplete $progressPercent

        $expandedLoc = [Environment]::ExpandEnvironmentVariables($loc)
        $isUserPath = $false
        foreach ($term in $userTerms) {
            $expandedTerm = [Environment]::ExpandEnvironmentVariables($term)
            if (($expandedLoc -match ".*$expandedTerm.*") -or ($loc -match ".*$term.*")) {
                $isUserPath = $true
                break
            }
        }
        if ($isUserPath) {
            $newpath_unique_user += $loc
        } else {
            $newpath_unique_system += $loc
        }
    }
    
    # Remove duplicates and format paths again
    $newpath_unique_user = $newpath_unique_user | Sort-Object -Unique
    $newpath_unique_system = $newpath_unique_system | Sort-Object -Unique
    
    # Sub in environment variables for user paths back in
    $newpath_unique_user = Format-Locations -locations $newpath_unique_user
    $newpath_unique_system = Format-Locations -locations $newpath_unique_system
    $newpath_unique_user = Resub -locations $newpath_unique_user -vardefs $vardefs
    $newpath_unique_system = Resub -locations $newpath_unique_system -vardefs $vardefs
    
    # Remove duplicates finally
    $newpath_unique_user = $newpath_unique_user | Sort-Object -Unique
    $newpath_unique_system = $newpath_unique_system | Sort-Object -Unique

    # Ensure the total paths add up correctly
    $totalPaths = $newpath.Count
    $totalUserPaths = $newpath_unique_user.Count
    $totalSystemPaths = $newpath_unique_system.Count
    if ($totalPaths -ne ($totalUserPaths + $totalSystemPaths)) {
        WriteOutputPretty -currentmsg "Error: Total paths do not add up correctly." -Color Red
        WriteOutputPretty -currentmsg "Total Paths: $totalPaths" -Color Red
        WriteOutputPretty -currentmsg "User Paths: $totalUserPaths" -Color Red
        WriteOutputPretty -currentmsg "System Paths: $totalSystemPaths" -Color Red
    
        exit
    }

    # Check if system PATH exceeds the character limit
    $newpath_str_system = $newpath_unique_system -join ";"
    $newpath_str_user = $newpath_unique_user -join ";"

    $systemPathLimit = 2047  # The actual limit is 2047 characters
    $userPathLimit = 2047

    # Display PATH stats
    WriteOutputPretty -currentmsg "PATH Statistics" -Color Cyan
    WriteOutputPretty -currentmsg @(
        "Total PATH Entries: $totalPaths",
        "User PATH Entries: ($totalUserPaths)",
        "User PATH Length: $($newpath_str_user.Length) characters",
        "System PATH Entries: ($totalSystemPaths)",
        "System PATH Length: $($newpath_str_system.Length) characters")
    
    WriteOutputPretty -currentmsg "User Path" -Color Cyan
    WriteOutputPretty -currentmsg $newpath_unique_user 
    
    WriteOutputPretty -currentmsg "System Path" -Color Cyan
    WriteOutputPretty -currentmsg $newpath_unique_system 
    
    WriteOutputPretty -currentmsg "Environment Variables" -Color Cyan
    WriteOutputPretty -currentmsg $vardefs
    
    
    # Check if PATH lengths are within limits before setting them
    if ($newpath_str_system.Length -lt $systemPathLimit -and $newpath_str_user.Length -lt $userPathLimit) {
        [System.Environment]::SetEnvironmentVariable("Path", $newpath_str_user, "User")
        [System.Environment]::SetEnvironmentVariable("Path", $newpath_str_system, "Machine")
        WriteOutputPretty -currentmsg "PATH variables updated successfully" -Color Green
    } else {
        WriteOutputPretty -currentmsg "ERROR updating PATH variables" -Color Red
        Write-Host "One or more PATH variables exceed the character limit." -ForegroundColor Red
        Write-Host "System PATH Length: $($newpath_str_system.Length) characters" -ForegroundColor Red
        Write-Host "User PATH Length: $($newpath_str_user.Length) characters" -ForegroundColor Red
        WriteOutputPretty -currentmsg "Consider reviewing the PATH entries to reduce their length." -Color Yellow
    }

    WriteOutputPretty -currentmsg "Script execution completed" -Color Cyan
}

# Call the Main function
Main
