function Get-MTGConfiguration
{
    $File = "$($env:LOCALAPPDATA)\ExchangeTrackingGUI\Config.json"
    if (Test-Path $File) {
        $Config = Get-Content -Path $File -Raw | ConvertFrom-Json
    }
    else {
        $Config = New-Object PSObject -Property @{
            PSConnect = "srv01.domain.local"
            Servers = @( @{Name="srv01"}, @{Name="srv02"} )
            EventID = @(
                @{ID="RECEIVE"},
                @{ID="SEND"},
                @{ID="SENDEXTERNAL"},
                @{ID="FAIL"},
                @{ID="DSN"},
                @{ID="DELIVER"},
                @{ID="BADMAIL"},
                @{ID="RESOLVE"},
                @{ID="EXPAND"},
                @{ID="REDIRECT"},
                @{ID="TRANSFER"},
                @{ID="SUBMIT"},
                @{ID="POISONMESSAGE"},
                @{ID="DEFER"}
            )
            Fields = @(
                @{F="Timestamp"},
                @{F="EventId"},
                @{F="Source"},
                @{F="MessageSubject"}
                @{F="Sender"},
                @{F="Recipients"},
                @{F="ClientIP"},
                @{F="ClientHostName"},
                @{F="ServerIP"},
                @{F="ServerHostName"},
                @{F="RecipientStatus"},
                @{F="TotalBytes"}
            )
        }
        Save-MTGConfiguration($Config)
    }

    return $Config
}
