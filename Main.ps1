Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Diagnostics;

public class WindowActions
{
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int X, int Y);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

    public static void ClickAt(int x, int y)
    {
        SetCursorPos(x, y);
        mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero); // Mouse left button down
        mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero); // Mouse left button up
    }
}
"@

Add-Type -AssemblyName System.Windows.Forms

# Function to show the input form
function Show-InputForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Send Discord Messages"
    $form.Size = New-Object System.Drawing.Size(400, 250)
    $form.StartPosition = "CenterScreen"

    # Label for message count
    $lblCount = New-Object System.Windows.Forms.Label
    $lblCount.Text = "Number of Messages:"
    $lblCount.AutoSize = $true
    $lblCount.Location = New-Object System.Drawing.Point(10, 20)
    $form.Controls.Add($lblCount)

    # Textbox for message count
    $txtCount = New-Object System.Windows.Forms.TextBox
    $txtCount.Location = New-Object System.Drawing.Point(150, 18)
    $form.Controls.Add($txtCount)

    # Label for message content
    $lblMessage = New-Object System.Windows.Forms.Label
    $lblMessage.Text = "Message Content:"
    $lblMessage.AutoSize = $true
    $lblMessage.Location = New-Object System.Drawing.Point(10, 60)
    $form.Controls.Add($lblMessage)

    # Textbox for message content
    $txtMessage = New-Object System.Windows.Forms.TextBox
    $txtMessage.Location = New-Object System.Drawing.Point(150, 58)
    $txtMessage.Size = New-Object System.Drawing.Size(200, 20)
    $form.Controls.Add($txtMessage)

    # Button to submit the form
    $btnSubmit = New-Object System.Windows.Forms.Button
    $btnSubmit.Text = "Send"
    $btnSubmit.Location = New-Object System.Drawing.Point(150, 100)
    $btnSubmit.Add_Click({
        $global:messageCount = [int]$txtCount.Text
        $global:messageContent = $txtMessage.Text
        $form.Close()
    })
    $form.Controls.Add($btnSubmit)

    $form.ShowDialog()
}

# Function to copy the message to the clipboard
function Set-Clipboard {
    param (
        [string]$text
    )
    $clipboard = [System.Windows.Forms.Clipboard]::SetText($text)
}

# Show the input form and get user input
Show-InputForm

# Fix for handling multiple processes
$discordProcess = Get-Process -Name "Discord" -ErrorAction SilentlyContinue | Select-Object -First 1

if ($discordProcess -and $messageCount -gt 0 -and $messageContent) {
    $discordWindowHandle = $discordProcess.MainWindowHandle
    [WindowActions]::SetForegroundWindow($discordWindowHandle)

    for ($i = 0; $i -lt $messageCount; $i++) {
        # Copy message to clipboard
        Set-Clipboard -text $messageContent
        Start-Sleep -Milliseconds 1

        # Click at the specified coordinates
        [WindowActions]::ClickAt(477, 992)

        # Simulate Ctrl + V to paste the message and press Enter
        [System.Windows.Forms.SendKeys]::SendWait("^{v}") # Ctrl + V
        Start-Sleep -Milliseconds 1
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    }
} else {
    Write-Host "Discord process not found, or invalid input."
}
