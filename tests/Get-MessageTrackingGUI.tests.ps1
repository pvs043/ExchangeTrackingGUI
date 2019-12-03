Describe "Create Icon Test" {
    it "Create Icon success" {
        Initialize-MTGIcon | Should Not Be $null
    }
 }

 Describe "Get Configuration Test" {
     it "Get Configuration success" {
        $Config = Get-MTGConfiguration
        $Config | Should Not Be $null
        $Config.PSConnect | Should Not Be $null
     }
 }

 Describe "Save Configuration Test" {
     it "Save Configuration success" {
        $Config = Get-MTGConfiguration
        Save-MTGConfiguration($Config)
        "$($env:LOCALAPPDATA)\ExchangeTrackingGUI\Config.json" | Should Be Exist
     }
 }
