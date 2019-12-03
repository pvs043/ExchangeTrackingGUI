Describe "Create Icon Test" {
    it "Create Icon success" {
        Initialize-MTGIcon | Should Not Be $null
    }
 }
