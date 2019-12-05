function Update-ConfigForm {
    [CmdletBinding(SupportsShouldProcess)]
    Param()

    if ($PSCmdlet.ShouldProcess("Update PSCommand Successful")) {
        $PSConnect.Text = $Config.PSConnect

        $ConfigServers.Items.Clear()
        $Config.Servers | ForEach-Object {
            $ConfigServers.Items.Add($_.Name)
        } | Out-Null
        $ConfigEventID.Items.Clear()
        $Config.EventID | ForEach-Object {
            $ConfigEventID.Items.Add($_.ID)
        } | Out-Null
        $ConfigFields.Items.Clear()
        $Config.Fields | ForEach-Object {
            $ConfigFields.Items.Add($_.F)
        } | Out-Null
    }
}
