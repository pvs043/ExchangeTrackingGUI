Add-Type -AssemblyName PresentationFramework

function Get-MessageTrackingGUI {
    $Icon = Initialize-MTGIcon

    #region Main form
    [xml]$MainForm = @"
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Exchange tracking Log" Height="568" Width="444" ResizeMode="NoResize" WindowStartupLocation="CenterScreen">
    <Grid>
        <Label Content="Recipients" HorizontalAlignment="Left" Margin="20,14,0,0" VerticalAlignment="Top" Width="65"/>
        <CheckBox Name="CBRecipients" HorizontalAlignment="Left" Margin="90,20,0,0" VerticalAlignment="Top" Width="21"/>
        <TextBox Name="Recipients" HorizontalAlignment="Left" Height="23" Margin="116,16,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="201" IsEnabled="False"/>
        <Label Content="Sender" HorizontalAlignment="Left" Margin="20,46,0,0" VerticalAlignment="Top" Width="53"/>
        <CheckBox Name="CBSender" HorizontalAlignment="Left" Margin="90,51,0,0" VerticalAlignment="Top" Width="21"/>
        <TextBox Name="Sender" HorizontalAlignment="Left" Height="23" Margin="116,47,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="201" IsEnabled="False"/>
        <Label Content="Servers" HorizontalAlignment="Left" Margin="20,74,0,0" VerticalAlignment="Top" Width="54" RenderTransformOrigin="0.481,0.962"/>
        <CheckBox Name="CBServers" HorizontalAlignment="Left" Margin="90,81,0,0" VerticalAlignment="Top" Width="21" IsChecked="True"/>
        <ListBox Name="Servers" HorizontalAlignment="Left" Height="184" Margin="116,77,0,0" VerticalAlignment="Top" Width="201" SelectionMode="Multiple"/>
        <Label Content="EventID" HorizontalAlignment="Left" Margin="20,270,0,0" VerticalAlignment="Top" Width="53"/>
        <CheckBox Name="CBEventID" HorizontalAlignment="Left" Margin="90,277,0,0" VerticalAlignment="Top" Width="21"/>
        <ComboBox Name="EventID" HorizontalAlignment="Left" Margin="116,274,0,0" VerticalAlignment="Top" Width="201" IsEnabled="False"/>
        <Label Content="Subject" HorizontalAlignment="Left" Margin="20,300,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.615,0.154"/>
        <CheckBox Name="CBSubject" HorizontalAlignment="Left" Margin="90,307,0,0" VerticalAlignment="Top" Width="21"/>
        <TextBox Name="Subject" HorizontalAlignment="Left" Height="23" Margin="116,303,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="201" IsEnabled="False"/>
        <Label Content="Start" HorizontalAlignment="Left" Margin="20,335,0,0" VerticalAlignment="Top" RenderTransformOrigin="-0.158,-0.385"/>
        <CheckBox Name="CBStart" HorizontalAlignment="Left" Margin="90,342,0,0" VerticalAlignment="Top" Width="21" IsChecked="True"/>
        <DatePicker Name="StartDate" HorizontalAlignment="Left" Margin="116,337,0,0" VerticalAlignment="Top" Width="102"/>
        <ComboBox Name="StartHour" HorizontalAlignment="Left" Margin="225,338,0,0" VerticalAlignment="Top" Width="41"/>
        <Label Content=":" HorizontalAlignment="Left" Margin="265,337,0,0" VerticalAlignment="Top"/>
        <ComboBox Name="StartMin" HorizontalAlignment="Left" Margin="276,338,0,0" VerticalAlignment="Top" Width="41"/>
        <Label Content="End" HorizontalAlignment="Left" Margin="20,366,0,0" VerticalAlignment="Top" RenderTransformOrigin="-0.158,-0.385"/>
        <CheckBox Name="CBEnd" HorizontalAlignment="Left" Margin="90,372,0,0" VerticalAlignment="Top" Width="21" IsChecked="True"/>
        <DatePicker Name="EndDate" HorizontalAlignment="Left" Margin="116,368,0,0" VerticalAlignment="Top" Width="102"/>
        <ComboBox Name="EndHour" HorizontalAlignment="Left" Margin="225,368,0,0" VerticalAlignment="Top" Width="41"/>
        <Label Content=":" HorizontalAlignment="Left" Margin="265,366,0,0" VerticalAlignment="Top"/>
        <ComboBox Name="EndMin" HorizontalAlignment="Left" Margin="276,368,0,0" VerticalAlignment="Top" Width="41"/>
        <Label Content="Exchange Management Shell Command (one server):" HorizontalAlignment="Left" Margin="20,398,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.053,0.615"/>
        <TextBox Name="PSCommand" HorizontalAlignment="Left" Height="80" Margin="25,424,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="380" IsReadOnly="True"/>
        <Button Name="BTSearch" Content="Search" HorizontalAlignment="Left" Margin="335,16,0,0" VerticalAlignment="Top" Width="70" Height="54"/>
        <Button Name="BTExport" Content="Export" HorizontalAlignment="Left" Margin="335,77,0,0" VerticalAlignment="Top" Width="70" Height="54"/>
        <Button Name="BTConfig" Content="Config" HorizontalAlignment="Left" Margin="335,368,0,0" VerticalAlignment="Top" Width="70" Height="22"/>
    </Grid>
</Window>
"@
    # Create a main form
    $XMLReader = (New-Object System.Xml.XmlNodeReader $MainForm)
    $XMLForm   = [Windows.Markup.XamlReader]::Load($XMLReader)
    $XMLForm.Icon = $Icon
    # Load Controls
    $MainForm.SelectNodes("//*[@Name]") | ForEach-Object {
        Set-Variable -Name ($_.Name) -Value $XMLForm.FindName($_.Name) -Scope Global
    }
    #endregion MainForm

    #region Config form
    $PlusImageSource  = "$PSScriptRoot\plus-24.png"
    $MinusImageSource = "$PSScriptRoot\minus-24.png"
    [xml]$ConfigForm = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Config" Height="450" Width="550" ResizeMode="NoResize" WindowStartupLocation="CenterScreen">
    <Grid>
        <Label Content="Exchange PS connect to:" Margin="18,10,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <TextBox Name="PSConnect" Margin="166,12,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Height="24" Width="140" Text=""/>
        <Button Name="BTSaveConfig" Content="Save" Margin="384,12,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Width="74"/>
        <Label Content="Servers" Margin="18,41,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Width="104"/>
        <Label Content="EventID" Margin="162,41,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <Label Content="Output Fields" Margin="346,41,0,0" HorizontalAlignment="Left" VerticalAlignment="Top"/>
        <TextBox Name="AddServer" Margin="22,67,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Width="100" Height="24"/>
        <Button Name="BTServersPlus" Margin="130,67,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Width="24" Height="24">
            <Image Source="$PlusImageSource" Stretch = "Fill" />
        </Button>
        <ListBox Name="ConfigServers" Margin="22,91,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Width="100" Height="300"/>
        <Button Name="BTServersMinus" Margin="130,226,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Width="24" Height="24">
            <Image Source="$MinusImageSource" Stretch = "Fill" />
        </Button>
        <TextBox Name="AddEventID" Margin="166,67,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Width="140" Height="24"/>
        <Button Name="BTEventIDPlus" HorizontalAlignment="Left" Margin="314,67,0,0" VerticalAlignment="Top" Width="24" Height="24">
            <Image Source="$PlusImageSource" Stretch = "Fill" />
        </Button>
        <ListBox Name="ConfigEventID" Margin="166,91,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Width="140" Height="300"/>
        <Button Name="BTEventIDMinus" Margin="314,226,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Width="24" Height="24">
            <Image Source="$MinusImageSource" Stretch = "Fill" />
        </Button>
        <TextBox Name="AddFields" Margin="350,67,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Width="140" Height="24"/>
        <Button Name="BTFieldsPlus" Margin="498,67,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Width="24" Height="24">
            <Image Source="$PlusImageSource" Stretch = "Fill" />
        </Button>
        <ListBox Name="ConfigFields" Margin="350,91,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Width="140" Height="300"/>
        <Button Name="BTFieldsMinus" Margin="498,226,0,0" HorizontalAlignment="Left" VerticalAlignment="Top" Width="24" Height="24">
            <Image Source="$MinusImageSource" Stretch = "Fill" />
        </Button>
    </Grid>
</Window>
"@
    # Create Config form
    $XMLReader = (New-Object System.Xml.XmlNodeReader $ConfigForm)
    $XMLConfig = [Windows.Markup.XamlReader]::Load($XMLReader)
    $XMLConfig.Icon = $Icon
    $XMLConfig.Add_Closing({
        $_.Cancel = $true
        $this.Visibility = "Hidden"
    })
    # Load Controls
    $ConfigForm.SelectNodes("//*[@Name]") | ForEach-Object {
        Set-Variable -Name ($_.Name) -Value $XMLConfig.FindName($_.Name) -Scope Global
    }
    #endregion ConfigForm

    #region Setup Config form
    # Save Config
    $BTSaveConfig.Add_Click({
        if ($PSConnect.Text) {
            if ($Config.PSConnect -ne $PSConnect.Text) {
                $Config.PSConnect = $PSConnect.Text
                Connect-MTGExchange
            }
        }
        else {
            $Config.PSConnect = "srv1.domain.local"
        }
        $Config.Servers = @()
        foreach ($Srv in $ConfigServers.Items) {
            $Config.Servers += @{Name=$Srv}
        }
        $Config.EventID = @()
        foreach ($ID in $ConfigEventID.Items) {
            $Config.EventID += @{ID=$ID}
        }
        $Config.Fields = @()
        foreach ($F in $ConfigFields.Items) {
            $Config.Fields += @{F=$F}
        }
        Save-MTGConfiguration($Config)
        $XMLConfig.Visibility = "Hidden"
        Update-MainForm
    })

    # Add server to list
    $BTServersPlus.Add_Click({
        if ($AddServer.Text) {
            $ConfigServers.Items.Add($AddServer.Text)
            $AddServer.Clear()
        }
    })
    $AddServer.Add_KeyDown({
        if ($_.Key -eq "Enter") {
            if ($AddServer.Text) {
                $ConfigServers.Items.Add($AddServer.Text)
                $AddServer.Clear()
            }
        }
        if ($_.Key -eq "Escape") {
            $AddServer.Clear()
        }
    })

    # Remove server from list
    $BTServersMinus.Add_Click({
        $ConfigServers.Items.Remove($ConfigServers.SelectedItem)
    })
    $ConfigServers.Add_KeyDown({
        if ($_.Key -eq "Delete") {
            $ConfigServers.Items.Remove($ConfigServers.SelectedItem)
        }
    })

    # Add EventID to list
    $BTEventIDPlus.Add_Click({
        if ($AddEventID.Text) {
            $ConfigEventID.Items.Add($AddEventID.Text)
            $AddEventID.Clear()
        }
    })
    $AddEventID.Add_KeyDown({
        if ($_.Key -eq "Enter") {
            if ($AddEventID.Text) {
                $ConfigEventID.Items.Add($AddEventID.Text)
                $AddEventID.Clear()
            }
        }
        if ($_.Key -eq "Escape") {
            $AddEventID.Clear()
        }
    })

    # Remove EventID from list
    $BTEventIDMinus.Add_Click({
        $ConfigEventID.Items.Remove($ConfigEventID.SelectedItem)
    })
    $ConfigEventID.Add_KeyDown({
        if ($_.Key -eq "Delete") {
            $ConfigEventID.Items.Remove($ConfigEventID.SelectedItem)
        }
    })

    # Add Output field to list
    $BTFieldsPlus.Add_Click({
        if ($AddFields.Text) {
            $ConfigFields.Items.Add($AddFields.Text)
            $AddFields.Clear()
        }
    })
    $AddFields.Add_KeyDown({
        if ($_.Key -eq "Enter") {
            if ($AddFields.Text) {
                $ConfigFields.Items.Add($AddFields.Text)
                $AddFields.Clear()
            }
        }
        if ($_.Key -eq "Escape") {
            $AddFields.Clear()
        }
    })

    # Remove Output field from list
    $BTFieldsMinus.Add_Click({
        $ConfigFields.Items.Remove($ConfigFields.SelectedItem)
    })
    $ConfigFields.Add_KeyDown({
        if ($_.Key -eq "Delete") {
            $ConfigFields.Items.Remove($ConfigFields.SelectedItem)
        }
    })
    #endregion Setup Config form

    #region Setup Main form
    # Recipients
    $CBRecipients.Add_Click({
        if ($CBRecipients.IsChecked) {
            $Recipients.IsEnabled = $true
        }
        else {
            $Recipients.IsEnabled = $false
        }
        Update-MTGCommandText
    })
    $Recipients.Add_KeyUp({
        if ($_.Key -eq "Escape") { $Recipients.Clear() }
        Update-MTGCommandText
    })

    # Sender
    $CBSender.Add_Click({
        if ($CBSender.IsChecked) {
            $Sender.IsEnabled = $true
        }
        else {
            $Sender.IsEnabled = $false
        }
        Update-MTGCommandText
    })
    $Sender.Add_KeyUp({
        if ($_.Key -eq "Escape") { $Sender.Clear() }
        Update-MTGCommandText
    })

    # Servers
    $CBServers.Add_Click({
        if ($CBServers.IsChecked) {
            $Servers.IsEnabled = $true
        }
        else {
            $Servers.IsEnabled = $false
        }
    })

    # EventID
    $CBEventID.Add_Click({
        if ($CBEventID.IsChecked) {
            $EventID.IsEnabled = $true
        }
        else {
            $EventID.IsEnabled = $false
        }
        Update-MTGCommandText
    })
    $EventID.Add_DropDownClosed({Update-MTGCommandText})

    # Subject
    $CBSubject.Add_Click({
        if ($CBSubject.IsChecked) {
            $Subject.IsEnabled = $true
        }
        else {
            $Subject.IsEnabled = $false
        }
        Update-MTGCommandText
    })
    $Subject.Add_KeyUp({
        if ($_.Key -eq "Escape") { $Subject.Clear() }
        Update-MTGCommandText
    })

    $StartDate.Add_CalendarClosed({Update-MTGCommandText})
    $StartHour.Add_SelectionChanged({Update-MTGCommandText})
    $StartMin.Add_SelectionChanged({Update-MTGCommandText})
    $EndDate.Add_CalendarClosed({Update-MTGCommandText})
    $EndHour.Add_SelectionChanged({Update-MTGCommandText})
    $EndMin.Add_SelectionChanged({Update-MTGCommandText})

    # Buttons
    $BTSearch.Add_Click({
        Get-MTGTrackingLog | Out-GridView
    })
    $BTExport.Add_Click({
        $Path = "$($env:LOCALAPPDATA)\ExchangeTrackingGUI"
        if ( !(Test-Path $Path) ) {
            $null = New-Item $Path -ItemType Directory
        }
        $File = "$Path\Export.csv"
        Get-MTGTrackingLog | ConvertTo-Csv -Delimiter ';' -NoTypeInformation > $File
        Invoke-Item $File
    })
    $BTConfig.Add_Click({
        Update-ConfigForm
        $XMLConfig.ShowDialog() | Out-Null
    })
    #endregion Setup Main form

    # Get configuration
    $Config = Get-MTGConfiguration

    # Setup main form
    Update-MainForm

    # Setup config Form
    Update-ConfigForm

    # Connect to Exchange 2016
    Connect-MTGExchange

    # Show Main form
    $XMLForm.ShowDialog() | Out-Null

    # Close connection
    if ($SessionEx) {
        Remove-PSSession -Session $SessionEx
    }
}
