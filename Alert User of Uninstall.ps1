# Add required .NET Assemblies for GUI components
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Text and Button Configuration ---
$messageText = "As part of Company's IT policy, Microsoft Office must be updated to a " +
               "new version. This will require uninstalling the current version.`n" + 
               "The installation will take approximately 20-30 minutes. Please " +
               "save/close any open files you might have.`n`n" + 
               "Click Accept to begin this process or Wait to be prompted again in " +
               "10 minutes. If you have any issues, please contact IT."

$delayMinutes = 10
$delayButtonText = "Wait ($($delayMinutes) mins)"

# --- Main Loop ---
while ($true) {
    # --- Form Creation ---
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "IT Notification: Microsoft Office Update"
    $form.Size = New-Object System.Drawing.Size(480, 265)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true 

    # --- Font Definition ---
    $font = New-Object System.Drawing.Font("Segoe UI", 11)

    # --- Label (Message Text) ---
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $messageText
    $label.Font = $font
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $label.Size = New-Object System.Drawing.Size(425, 145) 
    $form.Controls.Add($label)

    # --- Accept Button ---
    $acceptButton = New-Object System.Windows.Forms.Button
    $acceptButton.Text = "Accept"
    $acceptButton.Font = $font
    $acceptButton.Location = New-Object System.Drawing.Point(120, 180)
    $acceptButton.Size = New-Object System.Drawing.Size(100, 30)
    $acceptButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($acceptButton)

    # --- Delay/Wait Button ---
    $delayButton = New-Object System.Windows.Forms.Button
    $delayButton.Text = $delayButtonText
    $delayButton.Font = $font
    $delayButton.Location = New-Object System.Drawing.Point(240, 180)
    $delayButton.Size = New-Object System.Drawing.Size(120, 30)
    $delayButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($delayButton)

    $form.AcceptButton = $acceptButton

    # --- Show the form and wait for a click ---
    $result = $form.ShowDialog()

    # --- Handle the result ---
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "User clicked Accept. Exiting."
        break
    } else {
        Write-Host "User clicked Wait. Snoozing for $delayMinutes minutes."
        Start-Sleep -Seconds ($delayMinutes * 60)
    }
    
    $form.Dispose()
}
