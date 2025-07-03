' Get TEMP directory path
Set shell = CreateObject("WScript.Shell")
tempPath = shell.ExpandEnvironmentStrings("%TEMP%")

' Define screenshot path
Set fso = CreateObject("Scripting.FileSystemObject")
screenshotPath = fso.BuildPath(tempPath, "screenshot.png")

' Escape backslashes for PowerShell
screenshotPathPS = Replace(screenshotPath, "\", "\\")

' Take a screenshot using PowerShell
powerShellCommand = "Add-Type -AssemblyName System.Windows.Forms; " & _
                    "Add-Type -AssemblyName System.Drawing; " & _
                    "$screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds; " & _
                    "$bitmap = New-Object System.Drawing.Bitmap $screen.Width, $screen.Height; " & _
                    "$graphics = [System.Drawing.Graphics]::FromImage($bitmap); " & _
                    "$graphics.CopyFromScreen(0, 0, 0, 0, $screen.Size); " & _
                    "$bitmap.Save('" & screenshotPathPS & "', 'Png'); " & _
                    "$graphics.Dispose(); $bitmap.Dispose();"

shell.Run "powershell -Command " & Chr(34) & powerShellCommand & Chr(34), 0, True
