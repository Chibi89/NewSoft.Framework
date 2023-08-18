'@2023 Newspoint Software
Imports System.Windows.Forms
''' <summary>
''' indice che ragruppa funzioni di networking
''' </summary>
Public Module Network
    ''' <summary>
    ''' indice che ragruppa funzioni di networking
    ''' </summary>
    Public Class Webrequest
        ''' <summary>
        ''' A seconda del parametro passato ricorre un'azione.
        ''' exsample:
        ''' Webrequest.GoNet("next", WebBrowser1)
        ''' Webrequest.GoNet("back", WebBrowser1)
        ''' Webrequest.GoNet("home", WebBrowser1)
        ''' </summary>
        Public Shared Function GoNet(action As String, webBrowser As WebBrowser)
            Select Case action
                Case "next"
                    If webBrowser.CanGoForward Then
                        webBrowser.GoForward()
                    End If
                Case "back"
                    If webBrowser.CanGoBack Then
                        webBrowser.GoBack()
                    End If
                Case "home"
                    webBrowser.Navigate("http://www.example.com")
            End Select
            Return ""
        End Function
        Public Shared Sub Navigate(url As String, webBrowser As WebBrowser)
            If Not String.IsNullOrEmpty(url) AndAlso webBrowser IsNot Nothing Then
                webBrowser.Navigate(url)
            End If
        End Sub
        ''' <summary>
        ''' exsample: Webrequest.GoNet(textbox, WebBrowser1)
        ''' </summary>
        ''' <param name="textbox">definisce il controllo textbox da associare.</param>
        ''' <param name="webbrowser"> definisce il controllo web browser da associare.</param>
        ''' <returns>Una stringa che rappresenta l'architettura della CPU.</returns>
        Public Shared Function GoNet(textBox As TextBox, webBrowser As WebBrowser)
            If webBrowser IsNot Nothing AndAlso textBox IsNot Nothing Then
                Navigate(textBox.Text, webBrowser)
            End If
            Return ""
        End Function
        ''' <summary>
        ''' Esegue il ping da url passato tramite parametro
        ''' exsample: Webrequest.GoNetPing(textbox, Label)
        ''' </summary>
        Public Shared Function GoNetPing(urlTextBox As TextBox, resultLabel As Label)
            If urlTextBox IsNot Nothing AndAlso resultLabel IsNot Nothing Then
                Dim url As String = urlTextBox.Text
                Dim pingResult As String = PerformPing(url)
                resultLabel.Text = pingResult
            Else
                Return ""
            End If
            Return ""
        End Function
        Private Shared Function PerformPing(url As String) As String
            Dim ping As New System.Net.NetworkInformation.Ping()
            Dim reply As System.Net.NetworkInformation.PingReply = ping.Send(url)
            If reply.Status = System.Net.NetworkInformation.IPStatus.Success Then
                Return $"Ping riuscito a {url} - Tempo di risposta: {reply.RoundtripTime} ms"
            Else
                Return $"Ping non riuscito a {url} - Stato: {reply.Status}"
            End If
        End Function
        ''' <summary>
        ''' Esegue la verifica della connessione internet
        ''' </summary>
        Public Shared Function NetAvailable() As Boolean
            Dim ping As New System.Net.NetworkInformation.Ping()
            Try
                Dim reply As System.Net.NetworkInformation.PingReply = ping.Send("www.google.com")
                If reply.Status = System.Net.NetworkInformation.IPStatus.Success Then
                    Return True ' Connessione Internet disponibile
                    Return "Connesso a internet"
                Else
                    Return False ' Connessione Internet non disponibile
                    Return "Nessuna connessione a internet"
                End If
            Catch ex As Exception
                Return False ' Connessione Internet non disponibile
                Return "Nessuna connessione a internet"
            End Try
        End Function
    End Class
End Module
