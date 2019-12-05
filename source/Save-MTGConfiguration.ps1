function Save-MTGConfiguration($Config)
{
    $Path = "$($env:LOCALAPPDATA)\ExchangeTrackingGUI"
    if ( !(Test-Path $Path) ) {
        $null = New-Item $Path -ItemType Directory
    }
    $File = "$Path\Config.json"
    $Content = $Config | ConvertTo-Json
    Set-Content -Path $File -Value $Content -Force
}
