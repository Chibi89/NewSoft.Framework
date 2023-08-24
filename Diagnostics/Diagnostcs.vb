Imports System.IO
Imports System.Net
Imports System.Reflection
Imports NewSoft.Framework.PrivateDiagnostic
Public Module Diagnostcs
    Public Class ForApplication
        Public Function GetNumExecuteLib(url As String) As Integer
            Try
                Dim currentNumber As Integer = ReadCurrentNumber(url)
                Dim newNumber As Integer = currentNumber + 1
                WriteNewNumber(url, newNumber)
                TechTacker.TrackFunctionUsage("Diagnostic :: For Application :: GetNumExecuteLib")
                Return newNumber
            Catch ex As Exception
                Return -1 ' Indica che si è verificato un errore
            End Try
        End Function
        Private Function ReadCurrentNumber(url As String) As Integer
            Dim filePath As String = Path.Combine(url)
            If File.Exists(filePath) Then
                Dim content As String = File.ReadAllText(filePath)
                Dim currentNumber As Integer
                If Integer.TryParse(content, currentNumber) Then
                    Return currentNumber
                End If
            End If
            Return 0 ' Default value if the file doesn't exist or content isn't valid.
        End Function
        Private Sub WriteNewNumber(url As String, newNumber As Integer)
            Dim filePath As String = Path.Combine(url)
            File.WriteAllText(filePath, newNumber.ToString())
        End Sub
    End Class
    Public Class ForServer
        Private Const MaxEntriesBeforeFlush As Integer = 10
        Public Function LogFunctionCall(logFileUrl As String) As Boolean
            Try
                Dim methodName As String = GetCallingMethodName()
                Dim logEntries As List(Of String) = New List(Of String)()
                logEntries.Add($"Functions 01 = '{methodName}'")
                Dim logContent As String = String.Join(Environment.NewLine, logEntries)
                Using client As New WebClient()
                    client.UploadString(logFileUrl, logContent)
                End Using
                Console.WriteLine($"Logged function call: '{methodName}'")
                TechTacker.TrackFunctionUsage("Diagnostic :: For Server :: LogFunctionCall")
                Return True
            Catch ex As Exception
                Console.WriteLine($"Error writing log: {ex.Message}")
                Return False
            End Try
        End Function
        Private Function GetCallingMethodName() As String
            Dim stackTrace As New StackTrace()
            Dim callingMethod As MethodBase = stackTrace.GetFrame(2).GetMethod()
            Return callingMethod.Name
        End Function
    End Class
End Module