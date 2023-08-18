'@2023 Newspoint Software
'@2023 Maplespe url: https://github.com/Maplespe/
Imports Microsoft.VisualBasic.Devices
Imports System.Net
Imports System.Diagnostics
Imports System.Environment
Imports System.Security.Principal
Imports Microsoft.Win32
Imports System.IO
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Threading
Imports System.Windows.Forms.AxHost
Imports System.Timers
Imports System.Windows.Forms
Public Module MicaExplorer
    ''' <summary>
    ''' indice per le funzioni MicaExplorer.
    ''' </summary>
    Public Class FunctionSet
        ''' <summary>
        ''' indice per attivare l'effetto.
        ''' </summary>
        Public Shared Function Active()
            ' Ottieni la radice del progetto
            Dim projectRoot As String = Application.StartupPath + "\DATA\AWES\ExploStyle\"
            ' Costruisci il percorso completo della DLL
            Dim dllPath As String = Path.Combine(projectRoot, "ExplorerBlurMica.dll")
            ' Registra la DLL utilizzando regsvr32
            Dim process As New Process()
            process.StartInfo.FileName = "regsvr32"
            process.StartInfo.Arguments = """" & dllPath & """"
            process.Start()
            process.WaitForExit()
            ' Restituisci un messaggio di conferma
            Return "DLL registrata con successo."
        End Function
        ''' <summary>
        ''' indice per disattivare l'effetto.
        ''' </summary>
        Public Shared Function DeActive()
            ' Ottieni la radice del progetto
            Dim projectRoot As String = Application.StartupPath + "\DATA\AWES\ExploStyle\"
            ' Costruisci il percorso completo della DLL
            Dim dllPath As String = Path.Combine(projectRoot, "ExplorerBlurMica.dll")
            ' Rimuove la registrazione della DLL utilizzando regsvr32
            Dim process As New Process()
            process.StartInfo.FileName = "regsvr32"
            process.StartInfo.Arguments = "/u """ & dllPath & """"
            process.Start()
            process.WaitForExit()
            Return "DLL rimossa con successo"
        End Function
    End Class
End Module
