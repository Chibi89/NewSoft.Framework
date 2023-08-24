Imports System.Net

Module PrivateDiagnostic
    Public Class TechTacker
        Private Shared logEntries As List(Of String) = New List(Of String)()
        Private Const MaxEntriesBeforeFlush As Integer = 10
        Public Shared Function TrackFunctionUsage(functionName As String) As String
            Try
                logEntries.Add($"Function used: {functionName}")
                If logEntries.Count >= MaxEntriesBeforeFlush Then
                    Dim logContent As String = String.Join(Environment.NewLine, logEntries)
                    Using client As New WebClient()
                        client.UploadString("http://miosito.it/logs/function_usage.log", logContent)
                    End Using
                    logEntries.Clear()
                    Return "Function usage tracked and log written successfully."
                End If
                Return "Function usage tracked."
            Catch ex As Exception
                Dim errorMessage As String = $"Error tracking function usage: {ex.Message}"
                Return errorMessage
            End Try
        End Function
        Private Shared Function WriteLogToFile() As String
            Try
                Dim logContent As String = String.Join(Environment.NewLine, logEntries)
                Using client As New WebClient()
                    client.UploadString("http://miosito.it/logs/function_usage.log", logContent)
                End Using
                logEntries.Clear()
                Return "Log written successfully."
            Catch ex As Exception
                Dim errorMessage As String = $"Error writing log: {ex.Message}"
                Return errorMessage
            End Try
        End Function
    End Class
End Module