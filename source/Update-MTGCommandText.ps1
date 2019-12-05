    # Visual change PSCommandText field
    function Update-MTGCommandText {
        [CmdletBinding(SupportsShouldProcess)]
        Param()

        if ($PSCmdlet.ShouldProcess("Update PSCommand Successful")) {
            $PSCommand.Text = "Get-MessageTrackingLog "
            if ($CBRecipients.IsChecked -and $Recipients.Text) {
                $PSCommand.Text += '-Recipients:' + $Recipients.Text + ' '
            }
            if ($CBSender.IsChecked -and $Sender.Text) {
                $PSCommand.Text += '-Sender "' + $Sender.Text + '" '
            }
            if ($CBEventID.IsChecked) {
                $PSCommand.Text += '-EventId "' + $EventID.Text.ToString() + '" '
            }
            if ($CBSubject.IsChecked -and $Subject.Text) {
                $PSCommand.Text += '-MessageSubject "' + $Subject.Text + '" '
            }
            if ($CBStart.IsChecked) {
                $Start = $StartDate.SelectedDate
                $PSCommand.Text += '-Start "' + ($Start.Month).ToString('00') + '/' + ($Start.Day).ToString('00') + '/' + $Start.Year + ' ' + $StartHour.SelectedValue + ':' + $StartMin.SelectedValue + '" '
            }
            if ($CBEnd.IsChecked) {
                $End = $EndDate.SelectedDate
                $PSCommand.Text += '-End "' + ($End.Month).ToString('00') + '/' + ($End.Day).ToString('00') + '/' + $End.Year + ' ' + $EndHour.SelectedValue + ':' + $EndMin.SelectedValue + '"'
            }
            $PSCommand.Text = $PSCommand.Text.TrimEnd(' ')
        }
    }
