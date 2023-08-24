'@2023 Newspoint Software
Imports Microsoft.VisualBasic.Devices
Imports System.Security.Principal
Imports Microsoft.Win32
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Drawing
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
        End Sub
        ''' <summary>
        ''' indice per poter settare le dimensioni del form.
        ''' FormSize(myForm, 800, 600)
        ''' </summary>
        Public Shared Sub formSize(form As Form, width As Integer, height As Integer)
            form.Size = New Size(width, height)
        End Sub
        ''' <summary>
        ''' indice per poter settare il trasparancykey color.
        ''' FormTransparencyKey(myForm, Color.White)
        ''' </summary>
        Public Shared Sub formTransparencyKey(form As Form, transparencyColor As Color)
            form.TransparencyKey = transparencyColor
        End Sub
        ''' <summary>
        ''' indice per poter settare lo start position.
        ''' FormStartPositions(myForm, "centerscreen")
        ''' </summary>
        Public Shared Sub formStartPositions(form As Form, position As String)
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
        ''' FormVisibleTopControl(myForm, True, True)
        ''' </summary>
        Public Shared Sub formVisibleTopControl(form As Form, maximizeBox As Boolean, minimizeBox As Boolean)
            form.MaximizeBox = maximizeBox
            form.MinimizeBox = minimizeBox
        End Sub
        ''' <summary>
        ''' indice per poter settare la tipologia del bordo del form.
        ''' FormBorderStyle(myForm, FormBorderStyle.FixedSingle)
        ''' </summary>
        Public Shared Sub formBorderStyle(form As Form, borderStyle As FormBorderStyle)
            form.FormBorderStyle = borderStyle
        End Sub
        ''' <summary>
        ''' indice per poter l'immagine di sfondo e il relativo layout.
        ''' FormBrgImage(myForm, "path/to/image.jpg", ImageLayout.Stretch)
        ''' </summary>
        Public Shared Sub formBrgImage(form As Form, imagePath As String, layout As ImageLayout)
            Dim image As Image = Image.FromFile(imagePath)
            form.BackgroundImage = image
            form.BackgroundImageLayout = layout
        End Sub
    End Class
    ' Public Class AdvanceFunction

    'End Class
End Module