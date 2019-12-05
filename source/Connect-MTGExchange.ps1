function Connect-MTGExchange {
    if ($SessionEx) {
        Remove-PSSession -Session $SessionEx
    }
    $Uri = "http://" + $Config.PSConnect + "/PowerShell/"
    $SessionEx = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $Uri -Authentication Kerberos -ErrorAction SilentlyContinue
    if ($SessionEx) {
        Import-PSSession $SessionEx -DisableNameChecking -AllowClobber
    }
}
