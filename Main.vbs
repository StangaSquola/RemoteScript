Option Explicit

Dim shell, fso, urlImmagine, urlMessaggio, percorsoLocale, percorsoScript, scriptPS, comandoPopup
Dim messaggioDalWeb, exec, comandoScaricaTesto, file

' --- CONFIGURAZIONE URL ---
urlImmagine = "https://raw.githubusercontent.com/StangaSquola/RemoteScript/immagine.jpg"
urlMessaggio = "https://raw.githubusercontent.com/StangaSquola/RemoteScript/messaggio.txt"

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' --- 1. RECUPERA IL MESSAGGIO DA GITHUB ---
comandoScaricaTesto = "powershell -WindowStyle Hidden -NoProfile -Command ""[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object System.Net.WebClient).DownloadString('" & urlMessaggio & "')"""

Set exec = shell.Exec(comandoScaricaTesto)
messaggioDalWeb = exec.StdOut.ReadAll
messaggioDalWeb = Trim(messaggioDalWeb)

' Se il file è vuoto o non raggiungibile
If messaggioDalWeb = "" Then messaggioDalWeb = "Messaggio non disponibile"

' --- 2. MOSTRA IL POPUP (Solo con il messaggio da GitHub) ---
' Ho rimosso ogni riferimento ai giorni di Natale
comandoPopup = "powershell -WindowStyle Hidden -NoProfile -Command ""Add-Type -AssemblyName System.Windows.Forms; " & _
    "$form = New-Object System.Windows.Forms.Form; " & _
    "$form.Text = 'GEA'; $form.Size = New-Object System.Drawing.Size(400,180); " & _
    "$form.StartPosition = 'CenterScreen'; $form.FormBorderStyle = 'FixedDialog'; " & _
    "$form.MaximizeBox = $false; $form.MinimizeBox = $false; $form.ControlBox = $false; " & _
    "$form.TopMost = $true; " & _
    "$label = New-Object System.Windows.Forms.Label; " & _
    "$label.Text = '" & messaggioDalWeb & "'; " & _
    "$label.TextAlign = 'MiddleCenter'; $label.Dock = 'Fill'; " & _
    "$label.Font = New-Object System.Drawing.Font('Segoe UI', 11); " & _
    "$form.Controls.Add($label); " & _
    "$panel = New-Object System.Windows.Forms.Panel; $panel.Dock = 'Bottom'; $panel.Height = 50; " & _
    "$btn = New-Object System.Windows.Forms.Button; $btn.Text = 'OK'; " & _
    "$btn.Left = 160; $btn.Top = 10; $btn.Add_Click({$form.Close()}); " & _
    "$panel.Controls.Add($btn); $form.Controls.Add($panel); " & _
    "$form.ShowDialog() | Out-Null"""

' Esecuzione popup
shell.Run comandoPopup, 0, True

' --- 3. CAMBIO SFONDO ---
percorsoLocale = shell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Pictures\sfondo_natale.jpg"
percorsoScript = shell.ExpandEnvironmentStrings("%TEMP%") & "\cambia_sfondo.ps1"

scriptPS = "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12" & vbCrLf & _
    "try {" & vbCrLf & _
    "    $wc = New-Object System.Net.WebClient" & vbCrLf & _
    "    $wc.DownloadFile('" & urlImmagine & "', '" & percorsoLocale & "')" & vbCrLf & _
    "    if (Test-Path '" & percorsoLocale & "') {" & vbCrLf & _
    "        $path = (Get-Item '" & percorsoLocale & "').FullName" & vbCrLf & _
    "        Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name Wallpaper -Value $path" & vbCrLf & _
    "        Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name WallpaperStyle -Value '10'" & vbCrLf & _
    "        $code = @'" & vbCrLf & _
    "using System; using System.Runtime.InteropServices;" & vbCrLf & _
    "public class Wallpaper { [DllImport(""user32.dll"")] public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni); }" & vbCrLf & _
    "'@" & vbCrLf & _
    "        Add-Type -TypeDefinition $code" & vbCrLf & _
    "        [Wallpaper]::SystemParametersInfo(0x0014, 0, $path, 0x0003)" & vbCrLf & _
    "    }" & vbCrLf & _
    "} catch {}"

Set file = fso.CreateTextFile(percorsoScript, True)
file.Write scriptPS
file.Close

shell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & percorsoScript & """", 0, True

' --- 4. PULIZIA ---
WScript.Sleep 2000
On Error Resume Next
fso.DeleteFile percorsoScript
On Error GoTo 0

Set fso = Nothing

Set shell = Nothing
