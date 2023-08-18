'@2023 Newspoint Software
Imports Microsoft.VisualBasic.Devices
Imports System.Security.Principal
Imports Microsoft.Win32
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
''' <summary>
''' indice per poter ottenere varie informazioni dal sistema in uso.
''' </summary>
Public Module SW
    Private Function ReadRegistryValue(rootKey As RegistryHive, subKeyPath As String, valueName As String) As String
        'accedi al valori dei registri
        Using root As RegistryKey = RegistryKey.OpenBaseKey(rootKey, RegistryView.Default)
            Using subKey As RegistryKey = root.OpenSubKey(subKeyPath)
                If subKey IsNot Nothing Then
                    Dim value As Object = subKey.GetValue(valueName)
                    If value IsNot Nothing Then
                        Return value.ToString()
                    End If
                End If
            End Using
        End Using
        Return ""
    End Function
    Dim configFile As String = Application.StartupPath & "\DATA\AWES\ExploStyle\config.ini"
    Dim processorName As String = infoSF.GetProcessorName()
    Dim graphicsCardName As String = infoHW.GetGraphicsCardName()
    Dim totalRAM As Double = infoHW.GetTotalRAM()
    Dim usedRAM As Double = infoHW.GetUsedRAM()
    Dim rootKey As RegistryHive = RegistryHive.LocalMachine
    Dim subKeyPath As String = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    Dim subKeyHWPach As String = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
    Dim subkeyBiosPach As String = "HARDWARE\DESCRIPTION\System\BIOS"
    Dim displayVer As String = ReadRegistryValue(rootKey, subKeyPath, "DisplayVersion")
    Dim editionID As String = ReadRegistryValue(rootKey, subKeyPath, "EditionID")
    Dim BuildBranch As String = ReadRegistryValue(rootKey, subKeyPath, "BuildBranch")
    Dim currentBuild As String = ReadRegistryValue(rootKey, subKeyPath, "CurrentBuild")
    Dim registeredOwner As String = ReadRegistryValue(rootKey, subKeyPath, "RegisteredOwner")
    'Dim subBuild As String = ReadRegistryValue(subKeyPath, "UBR")
    'Dim connectedAccountName As String = GetConnectedWindowsAccountName()
    Dim osArchitecture As Architecture = RuntimeInformation.OSArchitecture
    Dim cpuArchitecture As String = ""
    '----------------------------------------------------------------------------------
    ''' <summary>
    ''' indice per poter ottenere varie informazioni sul Bios.
    ''' </summary>
    Public Class BiosInfo
        ''' <summary>
        ''' Ottiene il nome del prodotto della scheda madre dal registro di sistema.
        ''' </summary>
        Private Shared subkeyBiosPath As String = "HARDWARE\DESCRIPTION\System\BIOS"
        Public Shared Function BoardManufacturer()
            Using regkey As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(subkeyBiosPath)
                If regkey IsNot Nothing Then
                    Dim baseBoardManufacturer As String = regkey.GetValue("BaseBoardManufacturer", "").ToString()
                    Return baseBoardManufacturer
                Else
                    Return "n/a"
                End If
            End Using
        End Function
        ''' <summary>
        ''' Ottiene il nome del prodotto della scheda madre dal registro di sistema.
        ''' </summary>
        ''' <returns>Restituisce il nome del prodotto della scheda madre o "n/a" se non disponibile.</returns>
        Public Shared Function BoardProduct()
            Using regkey As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(subkeyBiosPath)
                If regkey IsNot Nothing Then
                    Dim baseBoardProduct As String = regkey.GetValue("BaseBoardProduct", "").ToString()
                    Return baseBoardProduct
                Else
                    Return "n/a"
                End If
            End Using
        End Function
        ''' <summary>
        ''' Ottiene il nome della versione del Bios in uso.
        ''' </summary>
        ''' <returns>Restituisce il nome della versione del bios o "n/a" se non disponibile.</returns>
        Public Shared Function BiosVersion()
            Using regkey As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(subkeyBiosPath)
                If regkey IsNot Nothing Then
                    Dim biosVer As String = regkey.GetValue("BIOSVersion", "").ToString()
                    Return biosVer
                Else
                    Return "n/a"
                End If
            End Using
        End Function
        ''' <summary>
        ''' Ottiene il nome del produttore del bios.
        ''' </summary>
        ''' <returns>Restituisce il nome del produttore del bios o "n/a" se non disponibile.</returns>
        Public Shared Function BiosVendor()
            Using regkey As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(subkeyBiosPath)
                If regkey IsNot Nothing Then
                    Dim baseBoardProduct As String = regkey.GetValue("BaseBoardProduct", "").ToString()
                    Return baseBoardProduct
                Else
                    Return "n/a"
                End If
            End Using
        End Function
    End Class

    '-------------------------------------------------------------------------------
    ''' <summary>
    ''' indice per poter ottenere varie informazioni sull'Hardware del sistema in uso.
    ''' </summary>
    Public Class infoHW
        ''' <summary>
        ''' Restituisce una stringa che rappresenta l'architettura della CPU.
        ''' </summary>
        ''' <returns>Una stringa che rappresenta l'architettura della CPU.</returns>
        Public Shared Function GetCpuArchitecture() As String
            Dim osArchitectures As Architecture = RuntimeInformation.OSArchitecture
            Dim cpuArchitecture As String
            Select Case osArchitectures
                Case Architecture.X86
                    cpuArchitecture = "32 bit"
                Case Architecture.X64
                    cpuArchitecture = "64 bit"
                Case Architecture.Arm
                    cpuArchitecture = "ARM"
                Case Architecture.Arm64
                    cpuArchitecture = "ARM64"
                Case Else
                    cpuArchitecture = "Sconosciuta"
            End Select

            Return cpuArchitecture
        End Function
        ''' <summary>
        ''' Ottiene informazioni sulla CPU dal registro di sistema.
        ''' Sample: GetCpuInfo(nomelabel)
        ''' </summary>
        ''' <param name="Cpu">Il controllo Label in cui verranno visualizzate le informazioni sulla CPU.</param>
        Public Shared Sub GetCpuInfo(Cpu As Label)
            Dim subKeyHwPath As String = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
            Using regKey As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(subKeyHwPath)
                If regKey IsNot Nothing Then
                    Dim baseBoardProcessor As String = regKey.GetValue("ProcessorNameString", "").ToString()
                    Dim maxClockSpeed As Integer = CInt(regKey.GetValue("~MHz", 0))
                    Dim maxClockSpeedGHz As Double = maxClockSpeed / 1000.0
                    Cpu.Text = $"CPU: {baseBoardProcessor} ({maxClockSpeedGHz:F2} GHz)"
                Else
                    Cpu.Text = "CPU Sconosciuta"
                End If
            End Using
        End Sub
        ''' <summary>
        ''' Ottiene il nome della scheda grafica.
        ''' </summary>
        ''' <returns>Restituisce il nome della scheda grafica o "n/a" se non disponibile.</returns>
        Public Shared Function GetGraphicsCardName() As String
            Dim subkeyPath As String = "SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}"

            Using regkey As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(subkeyPath)
                If regkey IsNot Nothing Then
                    For Each subKeyName As String In regkey.GetSubKeyNames()
                        Dim subKey As RegistryKey = regkey.OpenSubKey(subKeyName)
                        If subKey IsNot Nothing AndAlso subKey.GetValue("DriverDesc") IsNot Nothing Then
                            Dim driverDesc As String = subKey.GetValue("DriverDesc").ToString()
                            subKey.Close()
                            Return driverDesc
                        End If
                        subKey.Close()
                    Next
                End If
            End Using

            Return ""
        End Function
        ''' <summary>
        ''' Ottiene il nome del produttore del bios.
        ''' </summary>
        ''' <returns>Sample:
        ''' Dim drive As DriveInfo = New DriveInfo("C")
        ''' Dim driveLetter As String = "C"
        ''' Dim labelControl As Label = C_Label ' Assumi che C_Label sia il nome del tuo controllo Label
        ''' Dim spaceControl As Label = C_Space ' Assumi che C_Space sia il nome del tuo controllo Label
        ''' Dim hardDiskInfo As String = GetHardDiskInfo(drive, driveLetter, labelControl, spaceControl)
        ''' </returns>
        Public Shared Function GetHardDiskInfo(drive As DriveInfo, driveLetter As String, labelControl As Label, spaceControl As Label) As String
            Dim fileSystemType As String = infoSF.GetFileSystemType(driveLetter)
            Dim model As String = GetHardDiskModel()
            Dim totalSize As Double = 0
            Dim freeSpace As Double = 0
            Dim driveLabel As String = ""
            If drive.IsReady Then
                totalSize = drive.TotalSize / (1024 * 1024 * 1024) ' Convert bytes to gigabytes
                freeSpace = drive.AvailableFreeSpace / (1024 * 1024 * 1024) ' Convert bytes to gigabytes
                driveLabel = drive.VolumeLabel
            End If
            If String.IsNullOrEmpty(driveLabel) Then
                labelControl.Text = $"Disco Locale - {driveLetter}:"
            Else
                labelControl.Text = $"{driveLabel} - {driveLetter}:"
            End If
            spaceControl.Text = $"Spazio totale: {totalSize:F2} GB - Spazio libero: {freeSpace:F2} GB"
            Return $"Model Drive {driveLetter}:\ : {model} - FileSystem: {fileSystemType}{vbCrLf} - Spazio totale: {totalSize:F2} GB - Spazio libero: {freeSpace:F2} GB"
        End Function
        ''' <summary>
        ''' Ottiene il nome del modello dell'hdd/ssd.
        ''' </summary>
        ''' <returns>Sample:
        ''' UpdateProgressBar("C"c, ProgressBar_C, LabelCState)
        ''' </returns>
        Public Shared Function GetHardDiskModel() As String
            'ottieni il modello del hhd/ssd
            Dim processStartInfo As New ProcessStartInfo()
            processStartInfo.FileName = "wmic"
            processStartInfo.Arguments = "diskdrive get model /format:list"
            processStartInfo.RedirectStandardOutput = True
            processStartInfo.UseShellExecute = False
            processStartInfo.CreateNoWindow = True
            Dim process As Process = Process.Start(processStartInfo)
            Dim output As String = process.StandardOutput.ReadToEnd()
            process.WaitForExit()
            Dim lines As String() = output.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
            Dim model As String = ""
            For Each line As String In lines
                If line.StartsWith("Model=") Then
                    model = line.Substring(6).Trim()
                    Exit For
                End If
            Next
            Return model
        End Function
        ''' <summary>
        ''' Ottiene il totale della ram espresso in GB.
        ''' </summary>
        Public Shared Function GetTotalRAM() As Double
            Dim computerInfo As New ComputerInfo()
            Dim totalRamInBytes As Long = computerInfo.TotalPhysicalMemory
            Dim totalRamInGB As Double = totalRamInBytes / (1024 * 1024 * 1024)
            Return totalRamInGB
        End Function
        ''' <summary>
        ''' Ottiene la ram in uso espresso in GB.
        ''' </summary>
        Public Shared Function GetUsedRAM() As Double
            Dim computerInfo As New ComputerInfo()
            Dim usedRamInBytes As Long = computerInfo.TotalPhysicalMemory - computerInfo.AvailablePhysicalMemory
            Dim usedRamInGB As Double = usedRamInBytes / (1024 * 1024 * 1024)
            Return usedRamInGB
        End Function
    End Class

    '---------------------------------------------------------------------------------------

    'Public Class AmbientInfo

    'End Class

    '---------------------------------------------------------------------------------------
    ''' <summary>
    ''' indice per poter ottenere varie informazioni dal sistema in uso.
    ''' </summary>
    Public Class infoSF
        Public Shared Function IsRunAsAdmin() As Boolean
            Dim identity As WindowsIdentity = WindowsIdentity.GetCurrent()
            Dim principal As New WindowsPrincipal(identity)
            Return principal.IsInRole(WindowsBuiltInRole.Administrator)
        End Function
        ''' <summary>
        ''' Verifica se l'applicazione è eseguita come Amministratore.
        ''' In caso errato chiude e lo riavvia come tale.
        ''' </summary>
        Public Shared Function RunAsAdmin()
            If IsRunAsAdmin() Then
                ' Inserisci qui il codice dell'applicazione se viene avviata come amministratore
            Else
                MessageBox.Show("L'applicazione deve essere eseguita come amministratore." + vbCrLf +
                    "Verrà riavviata con i privilegi amministrativi.")
                ' Ottieni il percorso dell'eseguibile dell'applicazione corrente
                Dim exePath As String = Application.ExecutablePath
                ' Avvia un nuovo processo dell'applicazione come amministratore
                Dim startInfo As New ProcessStartInfo()
                startInfo.FileName = exePath
                startInfo.Verb = "runas" ' Imposta il verbo su "runas" per eseguire come amministratore
                Try
                    Process.Start(startInfo)
                Catch ex As Exception
                    ' Gestisci eventuali errori durante il riavvio come amministratore
                    MessageBox.Show("Errore durante il riavvio come amministratore: " & ex.Message)
                End Try
            End If
            Return ""
        End Function
        ''' <summary>
        ''' indice per poter ottenere varie informazioni dall'edizione del sistema in uso.
        ''' </summary>
        Public Shared Function GetWindowsEdition() As String
            Dim productName As String = ""

            Using regKey As RegistryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion")
                If regKey IsNot Nothing Then
                    productName = regKey.GetValue("ProductName", "").ToString()
                End If
            End Using

            Return productName
        End Function
        ''' <summary>
        ''' indice per poter ottenere varie informazioni dalla build di sistema in uso.
        ''' </summary>
        Public Shared Function GetWindowsBuildNumber() As String
            Return Environment.GetEnvironmentVariable("OSBuildNumber")
        End Function
        ''' <summary>
        ''' indice per poter ottenere varie informazioni dal processore in uso.
        ''' </summary>
        Public Shared Function GetProcessorName() As String
            Return Environment.GetEnvironmentVariable("PROCESSOR_IDENTIFIER")
        End Function
        ''' <summary>
        ''' indice per poter ottenere sull'username in uso.
        ''' </summary>
        Public Shared Function GetConnectedWindowsAccountName() As String
            Return Environment.GetEnvironmentVariable("USERNAME")
        End Function
        ''' <summary>
        ''' indice per poter ottenere informazioni sul tipo di account /Admin or Normal.
        ''' </summary>
        Public Shared Function IsAccountAdmin() As Boolean
            Dim principal As WindowsPrincipal = New WindowsPrincipal(WindowsIdentity.GetCurrent())
            Return principal.IsInRole(WindowsBuiltInRole.Administrator)
        End Function
        ''' <summary>
        ''' indice per poter ottenere il FileSystem del sistema in uso.
        ''' </summary>
        Public Shared Function GetFileSystemType(driveLetter As String) As String
            'ottieni il filesystem
            Dim drive As DriveInfo = New DriveInfo(driveLetter)
            If drive.IsReady Then
                Return drive.DriveFormat
            Else
                Return ""
            End If
        End Function
        ''' <summary>
        ''' Ottiene i valori che puntano verso un HD/SSD come spazio e nome etichetta.
        ''' </summary>
        ''' <returns>Sample:
        ''' UpdateProgressBar("C"c, ProgressBar_C, LabelCState)
        ''' </returns>
        Public Shared Function UpdateProgressBar(driveLetter As Char, progressBar As ProgressBar, labelState As Label)
            Dim drive As DriveInfo = New DriveInfo(driveLetter.ToString().ToUpper())
            Dim totalSize As Double = 0
            Dim freeSpace As Double = 0
            If drive.IsReady Then
                totalSize = drive.TotalSize / (1024 * 1024 * 1024) ' Convert bytes to gigabytes
                freeSpace = drive.AvailableFreeSpace / (1024 * 1024 * 1024) ' Convert bytes to gigabytes
                Dim spaceUsed As Double = totalSize - freeSpace
                Dim spacePercentage As Integer = CInt((spaceUsed / totalSize) * 100)
                ' Imposta il valore della ProgressBar e il testo del Label
                progressBar.Value = spacePercentage
                labelState.Text = $"{spacePercentage}%"
            End If
            Return ""
        End Function
    End Class

    '-------------------------------------------------------------------------------

    ' Public Class ShowInfo

    ' End Class
End Module