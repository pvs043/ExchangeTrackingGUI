function Update-MainForm {
    [CmdletBinding(SupportsShouldProcess)]
    Param()

    if ($PSCmdlet.ShouldProcess("Update PSCommand Successful")) {
        $Servers.Items.Clear()
        $Config.Servers | ForEach-Object {
            $Servers.Items.Add($_.Name)
        } | Out-Null
        $Servers.SelectAll()

        $EventID.Items.Clear()
        $Config.EventID | ForEach-Object {
            $EventID.Items.Add($_.ID)
        } | Out-Null
        $EventID.Text = $Config.EventID[0].ID

        # Setup Date and Time
        $StartDate.SelectedDate = Get-Date
        $EndDate.SelectedDate = Get-Date

        0..23 | ForEach-Object {
            $Hour = $_.ToString('00')
            $StartHour.Items.Add($Hour)
            $EndHour.Items.Add($Hour)
        } | Out-Null
        $StartHour.SelectedValue = (((Get-Date).AddMinutes(-10)).Hour).ToString('00')
        $EndHour.SelectedValue   = ((Get-Date).Hour).ToString('00')

        0..59 | ForEach-Object {
            $Min = $_.ToString('00')
            $StartMin.Items.Add($Min)
            $EndMin.Items.Add($Min)
        } | Out-Null
        $StartMin.SelectedValue = (((Get-Date).AddMinutes(-10)).Minute).ToString('00')
        $EndMin.SelectedValue   = ((Get-Date).Minute).ToString('00')

        # Construct PSCommand
        Update-MTGCommandText
    }
}
