'@2023 Newspoint Software
Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports NewSoft.Framework.PrivateDiagnostic
Public Module MicaLibrary
    Public Class ParameterTypes
        <Flags>
        Public Enum DWMWINDOWATTRIBUTE
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            DWMWA_SYSTEMBACKDROP_TYPE = 38
        End Enum
        <StructLayout(LayoutKind.Sequential)>
        Public Structure MARGINS
            Public cxLeftWidth As Integer
            Public cxRightWidth As Integer
            Public cyTopHeight As Integer
            Public cyBottomHeight As Integer
        End Structure
    End Class
    Public Class MicaEffectMethods
        <DllImport("DwmApi.dll")>
        Public Shared Function DwmExtendFrameIntoClientArea(ByVal hwnd As IntPtr, ByRef pMarInset As ParameterTypes.MARGINS) As Integer
        End Function

        <DllImport("dwmapi.dll")>
        Public Shared Function DwmSetWindowAttribute(ByVal hwnd As IntPtr, ByVal dwAttribute As ParameterTypes.DWMWINDOWATTRIBUTE, ByRef pvAttribute As Integer, ByVal cbAttribute As Integer) As Integer
        End Function

        Public Shared Function ExtendFrame(ByVal hwnd As IntPtr, ByVal margins As ParameterTypes.MARGINS) As Integer
            Return DwmExtendFrameIntoClientArea(hwnd, margins)
        End Function

        Public Shared Function SetWindowAttribute(ByVal hwnd As IntPtr, ByVal attribute As ParameterTypes.DWMWINDOWATTRIBUTE, ByVal parameter As Integer) As Integer
            Return DwmSetWindowAttribute(hwnd, attribute, parameter, Marshal.SizeOf(Of Integer)())
        End Function
    End Class
    ''' <summary>
    ''' Applica l'estensione del frame per l'effetto mica. 
    ''' Suggerimento: ApplyEffectAllForm(me or your form name,color.yourcolor)
    ''' <example>
    ''' <code>
    ''' ApplyEffectAllForm(me or your form name,color.yourcolor,0/1,0/4)
    ''' </code>
    ''' </example>
    ''' </summary>
    Public Class LoadMicaSet
        Public Shared Sub ApplyEffectAllForm(ByRef form As Form, bgColor As Color, immersiveDarkModeValue As Integer, systemBackdropTypeValue As Integer)
            Dim bounds As New ParameterTypes.MARGINS
            Dim hwnd As IntPtr = form.Handle
            With bounds
                .cxLeftWidth = 0 : .cxRightWidth = 0
                .cyTopHeight = Screen.PrimaryScreen.Bounds.Height - 60 : .cyBottomHeight = 0
            End With
            MicaEffectMethods.ExtendFrame(hwnd, bounds)
            form.BackColor = bgColor ' Imposta il colore di sfondo del form
            Dim Panel As New Panel
            With Panel
                .Size = New Size(form.Width, form.Height - 60)
                .Location = New Point(0, 0)
                .Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Top Or AnchorStyles.Bottom
                .BackColor = Color.FromKnownColor(KnownColor.Control)
            End With
            MicaEffectMethods.SetWindowAttribute(form.Handle, ParameterTypes.DWMWINDOWATTRIBUTE.DWMWA_USE_IMMERSIVE_DARK_MODE, immersiveDarkModeValue)
            MicaEffectMethods.SetWindowAttribute(form.Handle, ParameterTypes.DWMWINDOWATTRIBUTE.DWMWA_SYSTEMBACKDROP_TYPE, systemBackdropTypeValue)
            TechTacker.TrackFunctionUsage($"MicaEffect All Form {form} :: Active")
        End Sub
    End Class
    Public Sub ApplyEffectToControl(ByRef control As Control, bgColor As Color, immersiveDarkModeValue As Integer, systemBackdropTypeValue As Integer)
        ' Calcola le dimensioni del controllo rispetto al form
        Dim bounds As New ParameterTypes.MARGINS
        Dim hwnd As IntPtr = control.FindForm().Handle ' Ottieni l'handle del form genitore del controllo
        With bounds
            .cxLeftWidth = 0 : .cxRightWidth = 0
            .cyTopHeight = control.Top : .cyBottomHeight = 0
        End With
        MicaEffectMethods.ExtendFrame(hwnd, bounds)

        ' Imposta il colore di sfondo del controllo
        control.BackColor = bgColor

        ' Imposta gli attributi di finestra del form genitore
        MicaEffectMethods.SetWindowAttribute(hwnd, ParameterTypes.DWMWINDOWATTRIBUTE.DWMWA_USE_IMMERSIVE_DARK_MODE, immersiveDarkModeValue)
        MicaEffectMethods.SetWindowAttribute(hwnd, ParameterTypes.DWMWINDOWATTRIBUTE.DWMWA_SYSTEMBACKDROP_TYPE, systemBackdropTypeValue)
        TechTacker.TrackFunctionUsage($"MicaEffect for {control} :: Active")
    End Sub
End Module