function Get-MTGTrackingLog {
    if ($CBServers.IsChecked) {
        $Fields = ""
        $Config.Fields.F | ForEach-Object {
            $Fields += $_ + ','
        }
        $Fields = $Fields.TrimEnd(',')

        $Path = "$($env:LOCALAPPDATA)\ExchangeTrackingGUI"
        if ( !(Test-Path $Path) ) {
            $null = New-Item $Path -ItemType Directory
        }
        $File = "$Path\GetLog.ps1"

        $Servers.SelectedItems | ForEach-Object {
            $Command = $PSCommand.Text + ' -Server "' + $_.ToString() + '" | Select-Object ' + $Fields
            Set-Content -Path $File -Value $Command -Force
            & $File
        }
    }
}
