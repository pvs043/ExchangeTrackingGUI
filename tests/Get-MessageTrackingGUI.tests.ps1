Describe "PSScriptAnalyzer Test" {
    Context 'PSScriptAnalyzer Standard Rules' {
        $analysis = Invoke-ScriptAnalyzer -Path  '.\Get-MessageTrackingGUI.psm1'
        $scriptAnalyzerRules = Get-ScriptAnalyzerRule
        forEach ($rule in $scriptAnalyzerRules) {
            It "Should pass $rule" {
                If ($analysis.RuleName -contains $rule) {
                    $analysis | Where-Object RuleName -EQ $rule -outvariable failures | Out-Default
                    $failures.Count | Should Be 0
                }
            }
        }
    }
}
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
        Get-Item -Path "$($env:LOCALAPPDATA)\ExchangeTrackingGUI\Config.json" -ErrorAction SilentlyContinue | Should Not Be $null
     }
 }
