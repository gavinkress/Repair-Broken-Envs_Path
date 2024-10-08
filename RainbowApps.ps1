function WriteOutputPretty {
    param (
        [string]$currentmsg,
        [ConsoleColor]$Color = "Green"
    )
    $width = (Get-Host).UI.RawUI.MaxWindowSize.Width
    
    $outers = "-" * $width
    $n = $width - ($currentmsg.Length + 6)
    $nh = [math]::Floor($n / 2)
    $sides = "-"*$nh
    if ($n % 2 -eq 0) {$chrext = ""} else {$chrext = "-"}
    Write-Host ""
    Write-Host $outers -ForegroundColor $Color
    Write-Host "| $sides $currentmsg $sides$chrext |" -ForegroundColor $Color
    Write-Host $outers -ForegroundColor $Color
    Write-Host ""
}

$apps = @(Get-AppxPackage | Select-Object -Property Name)
for ($i=0; $i -lt 1000; $i--){
    $part = 5..($apps.Length-5) | Get-Random
    $crntmsg = @($sample[$part] -split [regex]"\s\s")[0]
    $color = [System.Enum]::GetValues([System.ConsoleColor]) | Get-Random -Count 1 
    WriteOutputPretty -currentmsg $crntmsg -Color $color
    Start-Sleep -Seconds 0.15
}
