'@2023 Newspoint Software
Imports System.IO
Imports System.Net
Imports System.Reflection
Imports System.Windows.Forms
Public Module Update
    Public Class LoadUpd
        Public Shared Function UpdCheck(dllPath As String) As String
            'UpdCheck("percorso\alla\libreria.dll")
            Dim assembly As Assembly = Assembly.LoadFrom(dllPath)
            Dim versione As Version = assembly.GetName().Version

            Return ("Versione della libreria: " & versione.ToString())
        End Function
        Public Shared Function UpdStart(fileUrl As String, downloadLabel As Label)
            'UpdStart("URL_del_file_da_scaricare", Label1)
            Try
                Using client As New WebClient()
                    AddHandler client.DownloadProgressChanged, Sub(sender, e)
                                                                   Dim percentage As Integer = CInt((e.BytesReceived / e.TotalBytesToReceive) * 100)
                                                                   Dim remainingSize As Long = e.TotalBytesToReceive - e.BytesReceived
                                                                   downloadLabel.Text = $"Download: {percentage}% - Rimasti: {remainingSize} byte"
                                                               End Sub
                    AddHandler client.DownloadFileCompleted, Sub(sender, e)
                                                                 downloadLabel.Text = "Download completato"
                                                             End Sub
                    Dim localFilePath As String = Path.GetFileName(fileUrl)
                    client.DownloadFileAsync(New Uri(fileUrl), localFilePath)
                End Using
            Catch ex As Exception
                If downloadLabel IsNot Nothing Then
                    downloadLabel.Text = "Errore durante il download: " & ex.Message
                Else
                    Throw New Exception("Errore durante il download: " & ex.Message)
                End If
            End Try
            Return ""
        End Function
    End Class
End Module
Public Module UpdateCore
    Public Function UpdCheck(dllPath As String) As String
        'UpdCheck("percorso\alla\libreria.dll")
        Dim assembly As Assembly = Assembly.LoadFrom(dllPath)
        Dim versione As Version = assembly.GetName().Version

        Return ("Versione della libreria: " & versione.ToString())
    End Function
    Public Function UpdStart(fileUrl As String, downloadLabel As Label)
        'UpdStart("URL_del_file_da_scaricare", Label1)
        Try
            Using client As New WebClient()
                AddHandler client.DownloadProgressChanged, Sub(sender, e)
                                                               Dim percentage As Integer = CInt((e.BytesReceived / e.TotalBytesToReceive) * 100)
                                                               Dim remainingSize As Long = e.TotalBytesToReceive - e.BytesReceived
                                                               downloadLabel.Text = $"Download: {percentage}% - Rimasti: {remainingSize} byte"
                                                           End Sub
                AddHandler client.DownloadFileCompleted, Sub(sender, e)
                                                             downloadLabel.Text = "Download completato"
                                                         End Sub
                Dim localFilePath As String = Path.GetFileName(fileUrl)
                client.DownloadFileAsync(New Uri(fileUrl), localFilePath)
            End Using
        Catch ex As Exception
            If downloadLabel IsNot Nothing Then
                downloadLabel.Text = "Errore durante il download: " & ex.Message
            Else
                Throw New Exception("Errore durante il download: " & ex.Message)
            End If
        End Try
        Return ""
    End Function
    Public Function UpdCheckAndStartDownload(libreriaPath As String, fileUrl As String, downloadLabel As Label) As String
        Dim versioneLibreria As String = UpdCheck(libreriaPath)
        If Not String.IsNullOrEmpty(versioneLibreria) Then
            ' La verifica della versione ha avuto successo, quindi procedi con il download
            UpdStart(fileUrl, downloadLabel)
            Return versioneLibreria
        Else
            ' La verifica della versione non ha avuto successo, restituisci un messaggio di errore
            Return "Verifica della versione non riuscita"
        End If
    End Function
End Module