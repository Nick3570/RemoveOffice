# Add required .NET Assemblies for GUI components
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Modern Look Configuration ---
$theme = @{
    BackgroundColor = [System.Drawing.ColorTranslator]::FromHtml("#F0F0F0") # A light grey
    TextColor       = [System.Drawing.ColorTranslator]::FromHtml("#333333")
    ButtonColor     = [System.Drawing.Color]::White
    ButtonHover     = [System.Drawing.ColorTranslator]::FromHtml("#E0E0E0")
    AccentColor     = [System.Drawing.ColorTranslator]::FromHtml("#0078D4") # Windows blue
}

# --- Text and Button Configuration ---
$messageText = "As part of Company's IT policy, Microsoft Office must be updated to a " +
               "new version. This will require uninstalling the current version.`n" +
               "The installation will take approximately 20-30 minutes. Please " +
               "save/close any open files you might have.`n" +
               "Click Accept to begin this process of Wait to be prompted again in " +
               "10 minutes. If you have any issues, please contact IT."

$delayMinutes = 10
$delayButtonText = "Wait ($($delayMinutes) mins)"

# --- Main Loop ---
while ($true) {
    # --- Form Creation ---
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "IT Notification: Microsoft Office Update"
    $form.Size = New-Object System.Drawing.Size(520, 265) # Made slightly wider for icon
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true 
    $form.FormBorderStyle = 'FixedDialog'
    $form.BackColor = $theme.BackgroundColor

    # --- Font Definition ---
    $font = New-Object System.Drawing.Font("Segoe UI", 11)

    # --- Icon ---
    $icon = [System.Drawing.SystemIcons]::Information
    $pictureBox = New-Object System.Windows.Forms.PictureBox
    $pictureBox.Image = $icon.ToBitmap()
    $pictureBox.Location = New-Object System.Drawing.Point(20, 20)
    $pictureBox.Size = New-Object System.Drawing.Size(32, 32)
    $form.Controls.Add($pictureBox)

    # --- Label (Message Text) ---
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $messageText
    $label.Font = $font
    $label.ForeColor = $theme.TextColor
    $label.Location = New-Object System.Drawing.Point(65, 20) # Moved right to make space for icon
    $label.Size = New-Object System.Drawing.Size(420, 145) 
    $form.Controls.Add($label)

    # --- Accept Button (Modernized) ---
    $acceptButton = New-Object System.Windows.Forms.Button
    $acceptButton.Text = "Accept"
    $acceptButton.Font = $font
    $acceptButton.Location = New-Object System.Drawing.Point(140, 180) 
    $acceptButton.Size = New-Object System.Drawing.Size(100, 35)
    $acceptButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $acceptButton.FlatAppearance.BorderSize = 1
    $acceptButton.FlatAppearance.BorderColor = $theme.AccentColor
    $acceptButton.BackColor = $theme.ButtonColor
    $acceptButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($acceptButton)

    # --- Delay/Wait Button (Modernized) ---
    $delayButton = New-Object System.Windows.Forms.Button
    $delayButton.Text = $delayButtonText
    $delayButton.Font = $font
    $delayButton.Location = New-Object System.Drawing.Point(260, 180)
    $delayButton.Size = New-Object System.Drawing.Size(120, 35)
    $delayButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $delayButton.FlatAppearance.BorderSize = 1
    $delayButton.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
    $delayButton.BackColor = $theme.ButtonColor
    $delayButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($delayButton)

    $form.AcceptButton = $acceptButton

    # --- Show the form and wait for a click ---
    $result = $form.ShowDialog()
    
    # ... (Rest of the script is the same) ...
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) { Write-Host "User clicked Accept. Exiting."; break }
    else { Write-Host "User clicked Wait. Snoozing for $delayMinutes minutes."; Start-Sleep -Seconds ($delayMinutes * 60) }
    $form.Dispose()
}
Result of Option 1:

Option 2: The Truly Modern Way (Using WPF)
This approach rebuilds the pop-up using WPF and XAML (a markup language for UIs). The PowerShell code becomes a controller for the interface defined in XAML. This separates the design from the logic and gives you a native Windows 10/11 look.

WPF/XAML Pop-Up Script
PowerShell

# Add required .NET Assemblies for WPF
Add-Type -AssemblyName PresentationFramework

# --- Configuration ---
$delayMinutes = 10

# --- XAML Definition for the Modern Window ---
# We define the entire window layout in this string.
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="IT Notification: Microsoft Office Update"
        Width="520" SizeToContent="Height" WindowStartupLocation="CenterScreen" Topmost="True"
        Background="#F0F0F0" ResizeMode="NoResize">
    <Grid Margin="20">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="*" />
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>

        <Viewbox Grid.Row="0" Grid.Column="0" Width="32" Height="32" Margin="0,0,15,0">
            <TextBlock Text="&#xE946;" FontFamily="Segoe MDL2 Assets" FontSize="32" VerticalAlignment="Top"/>
        </Viewbox>

        <TextBlock x:Name="MessageText" Grid.Row="0" Grid.Column="1" FontSize="15" TextWrapping="Wrap" Foreground="#333333">
            As part of CHD's IT policy, Microsoft Office must be updated to a new version. This will require uninstalling the current version.
            <LineBreak/><LineBreak/>
            The installation will take approximately 20-30 minutes. Please save/close any open files you might have.
            <LineBreak/><LineBreak/>
            Click Accept to begin this process of Wait to be prompted again in 10 minutes. If you have any issues, please contact IT. Thank you.
        </TextBlock>

        <StackPanel Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,25,0,0">
            <Button x:Name="AcceptButton" Content="Accept" Width="100" Height="35" Margin="5" IsDefault="True"/>
            <Button x:Name="WaitButton" Content="Wait (10 mins)" Width="120" Height="35" Margin="5" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@

# --- Main Loop ---
while ($true) {
    # Create the WPF window from XAML
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [System.Windows.Markup.XamlReader]::Load($reader)

    # Get the buttons by name from the XAML
    $acceptButton = $window.FindName("AcceptButton")
    $waitButton = $window.FindName("WaitButton")

    # This is a simple way to handle clicks for this use case.
    # ShowDialog() returns $true if a button with IsDefault="True" is clicked.
    # It returns $false if a button with IsCancel="True" is clicked or the window is closed.
    $result = $window.ShowDialog()

    if ($result) {
        Write-Host "User clicked Accept. Exiting."
        break
    } else {
        Write-Host "User clicked Wait. Snoozing for $delayMinutes minutes."
        Start-Sleep -Seconds ($delayMinutes * 60)
    }
}
