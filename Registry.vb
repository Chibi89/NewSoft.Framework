'@2023 Newspoint Software
Imports System.IO
Imports Microsoft.Win32
''' <summary>
''' Ottiene l'indice per le funzioni collegate al registro di sistema
''' </summary>
Public Module Registry
    ''' <summary>
    ''' Ottiene la funzione per la registrazione di una librearia dll.
    ''' </summary>
    Public Class RegistryFile
        'SetValue(rootLibrary,NameLibrary.dll)
        Public Shared Function SetValue(projectRoot As String, dllPathFile As String) As String
            ' Costruisci il percorso completo della DLL
            Dim dllPath = Path.Combine(projectRoot, dllPathFile)
            ' Registra la DLL utilizzando regsvr32
            Dim process As New Process()
            process.StartInfo.FileName = "regsvr32"
            process.StartInfo.Arguments = """" & dllPath & """"
            process.Start()
            process.WaitForExit()
            ' Restituisci un messaggio di conferma con il percorso della DLL registrata
            Dim resultMessage As String = $"DLL registrata con successo. Percorso: {dllPath}"
            Return resultMessage
        End Function
    End Class
    ''' <summary>
    ''' Ottiene la funzione per la rimozione di una libreria dll.
    ''' </summary>
    Public Class DeRegistryFile
        'SetValue(rootLibrary,NameLibrary.dll)
        Public Shared Function SetValue(projectRoot As String, dllPathFile As String) As String
            ' Costruisci il percorso completo della DLL
            Dim dllPath = Path.Combine(projectRoot, dllPathFile)
            ' Registra la DLL utilizzando regsvr32
            Dim process As New Process()
            process.StartInfo.FileName = "regsvr32"
            process.StartInfo.Arguments = "/u """ & dllPath & """"
            process.Start()
            process.WaitForExit()
            ' Restituisci un messaggio di conferma con il percorso della DLL registrata
            Dim resultMessage As String = $"DLL rimossa con successo. Percorso: {dllPath}"
            Return resultMessage
        End Function
    End Class
    ''' <summary>
    ''' Ottiene la funzione per verificare l'esistenza di una chiave nel registro di sistema.
    ''' </summary>
    Public Class VerifyFile
        Public Shared Function SetValue(RegistryPath As String) As Boolean
            Dim exists As Boolean = False
            Try
                If My.Computer.Registry.CurrentUser.OpenSubKey(RegistryPath) IsNot Nothing Then
                    exists = True
                End If
            Finally
                My.Computer.Registry.CurrentUser.Close()
            End Try
            Return ""
        End Function
    End Class
    ''' <summary>
    ''' Ottiene la funzione per la lettura di una chiave/valore dal registro di sistema.
    ''' </summary>
    Public Class ReadFile
        Public Shared Function SetValue(RegFilePath As String, ValueName As String) As String
            Dim keyValue As Object
            keyValue = My.Computer.Registry.GetValue(RegFilePath, ValueName, "Default Value")
            Return ""
        End Function
    End Class
    ''' <summary>
    ''' Ottiene la funzione per modificare una chiave/valore dal registro di sistema.
    ''' </summary>
    Public Class RenameValue
        Public Shared Function SetValue(RegFilePath As String, KeyName As String, Value As String)
            My.Computer.Registry.SetValue(RegFilePath, KeyName, Value)
            Return ""
        End Function
    End Class
    ''' <summary>
    ''' Ottiene la funzione per l'eliminazione di una chiave/valore dal registro di sistema.
    ''' </summary>
    Public Class DeleteValue
        Public Shared Function SetValue(filePath As String)
            Using key As RegistryKey = My.Computer.Registry.LocalMachine.OpenSubKey("Software")
                key.DeleteSubKey(filePath)
            End Using
            Return ""
        End Function
    End Class
    ''' <summary>
    ''' Ottiene la funzione per la creazione di una chiave/valore nel registro di sistema.
    ''' </summary>
    Public Class CreateValue
        Public Shared Function SetValue(hivetype As RegistryHive, keyPath As String, valueName As String, value As String) As Boolean
            Dim success As Boolean = False
            Dim hive As RegistryHive = RegistryHive.LocalMachine
            Select Case hivetype
                Case RegistryHive.CurrentUser
                    hive = RegistryHive.CurrentUser
                Case RegistryHive.CurrentConfig
                    hive = RegistryHive.CurrentConfig
                Case RegistryHive.LocalMachine
                    hive = RegistryHive.LocalMachine
                Case RegistryHive.ClassesRoot
                    hive = RegistryHive.ClassesRoot
                Case RegistryHive.Users
                    hive = RegistryHive.Users
            End Select
            Using regKey As RegistryKey = RegistryKey.OpenBaseKey(hive, RegistryView.Registry64).CreateSubKey(keyPath)
                If regKey IsNot Nothing Then
                    regKey.SetValue(valueName, value)
                    success = True
                End If
            End Using

            Return success
        End Function
    End Class
    Public Class RegistryHelper
        Public Shared Function GetRegistryValue(keyPath As String, valueName As String) As String
            Dim registryValue As String = ""

            Using regKey As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64).OpenSubKey(keyPath)
                If regKey IsNot Nothing Then
                    registryValue = regKey.GetValue(valueName, "").ToString()
                End If
            End Using
            Return registryValue
        End Function
    End Class
End Module
