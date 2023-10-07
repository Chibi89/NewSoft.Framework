'@2023 Newspoint Software
Imports Microsoft.VisualBasic.Devices
Imports System.Security.Principal
Imports Microsoft.Win32
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Drawing
Imports NewSoft.Framework.PrivateDiagnostic
''' <summary>
''' indice per poter ottenere varie funzioni che facilitano lo sviluppo nell'ide.
''' </summary>
Public Module PublicFunctions

    Public Class WinFormFunctionStyle
        ''' <summary>
        ''' indice per poter settare il nome del form.
        ''' <example>
        ''' <code>
        ''' FormTitle(myForm, "Il mio form")
        ''' </code>
        ''' </example>
        ''' </summary>
        Public Shared Sub FormTitle(form As Form, formText As String)
            form.Text = formText
            TechTacker.TrackFunctionUsage($"FormTitle :: {form} - {formText}")
        End Sub
        ''' <summary>
        ''' indice per poter settare le dimensioni del form.
        ''' FormSize(myForm, 800, 600)
        ''' </summary>
        Public Shared Sub FormSize(form As Form, Optional ByVal width As Integer = Nothing, Optional ByVal height As Integer = Nothing)
            form.Size = New Size(width, height)
            TechTacker.TrackFunctionUsage($"FormSize :: {form} - {width}x{height}")
        End Sub
        ''' <summary>
        ''' indice per poter settare il trasparancykey color.
        ''' FormTransparencyKey(myForm, Color.White)
        ''' </summary>
        Public Shared Sub FormTransparencyKey(form As Form, transparencyColor As Color)
            form.TransparencyKey = transparencyColor
            TechTacker.TrackFunctionUsage($"FormTrasparencyColorKey :: {form} - {transparencyColor}")
        End Sub
        ''' <summary>
        ''' indice per poter settare lo start position.
        ''' FormStartPositions(myForm, "manual/centerscreen/centerparent")
        ''' </summary>
        Public Shared Sub FormStartPositions(form As Form, position As String)
            TechTacker.TrackFunctionUsage($"StartPosition :: {form} - {position}")
            Select Case position.ToLower()
                Case "manual"
                    form.StartPosition = FormStartPosition.Manual
                Case "centerparent"
                    form.StartPosition = FormStartPosition.CenterParent
                Case "centerscreen"
                    form.StartPosition = FormStartPosition.CenterScreen
                Case Else
                    ' Gestione per un valore non valido
            End Select
        End Sub
        ''' <summary>
        ''' indice per poter settare il maximizebox e il minimizebox del form.
        ''' FormVisibleTopControl(myForm, True/False, True/False)
        ''' </summary>
        Public Shared Sub FormVisibleTopControl(form As Form, Optional ByVal maximizeBox As Boolean = Nothing, Optional ByVal minimizeBox As Boolean = Nothing)
            form.MaximizeBox = maximizeBox
            form.MinimizeBox = minimizeBox
        End Sub
        ''' <summary>
        ''' indice per poter settare la tipologia del bordo del form.
        ''' FormBorderStyle(myForm, FormBorderStyle.FixedSingle)
        ''' </summary>
        Public Shared Sub FormBorderStyle(form As Form, borderStyle As FormBorderStyle)
            form.FormBorderStyle = borderStyle
        End Sub
        ''' <summary>
        ''' indice per poter l'immagine di sfondo e il relativo layout.
        ''' FormBrgImage(myForm, "path/to/image.jpg", ImageLayout.Stretch)
        ''' </summary>
        Public Shared Sub FormBrgImage(form As Form, imagePath As String, Optional ByVal layout As ImageLayout = Nothing)
            Dim image As Image = Image.FromFile(imagePath)
            form.BackgroundImage = image
            form.BackgroundImageLayout = layout
        End Sub
    End Class
    ' Public Class AdvanceFunction

    'End Class
End Module
