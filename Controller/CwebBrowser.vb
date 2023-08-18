'@2023 Newspoint Software
Imports System.Windows.Forms

Public Class CwebBrowser
    ''' <summary>
    ''' Modulo WebBrowser Personalizzato con TabSystem
    ''' </summary>
    Private _tabControl As TabControl

    Public Sub New(parentControl As Control)
        _tabControl = New TabControl()
        _tabControl.Dock = DockStyle.Fill
        parentControl.Controls.Add(_tabControl)
        Dim addTabButton As New Button()
        addTabButton.Text = "Aggiungi Tab"
        addTabButton.Dock = DockStyle.Top
        parentControl.Controls.Add(addTabButton)
        Dim removeTabButton As New Button()
        removeTabButton.Text = "Rimuovi Tab"
        removeTabButton.Dock = DockStyle.Top
        parentControl.Controls.Add(removeTabButton)
        AddHandler addTabButton.Click, AddressOf AddTabButton_Click
        AddHandler removeTabButton.Click, AddressOf RemoveTabButton_Click
    End Sub
    Private Sub AddTabButton_Click(sender As Object, e As EventArgs)
        Dim tabPage As New TabPage("Nuova Scheda")
        Dim webBrowser As New WebBrowser()
        webBrowser.Dock = DockStyle.Fill ' Imposta il controllo WebBrowser per occupare tutto lo spazio disponibile nella scheda
        tabPage.Controls.Add(webBrowser)
        _tabControl.TabPages.Add(tabPage)
    End Sub
    Private Sub RemoveTabButton_Click(sender As Object, e As EventArgs)
        If _tabControl.SelectedTab IsNot Nothing Then
            _tabControl.TabPages.Remove(_tabControl.SelectedTab)
        End If
    End Sub
    ''' <summary>
    ''' Esegue la navigazione all'url specificato
    ''' Exsample: 
    ''' cwebBroser.NavUrl("https://www.google.com")
    ''' </summary>
    Public Sub NavUrl(url As String)
        If _tabControl.SelectedTab IsNot Nothing AndAlso _tabControl.SelectedTab.Controls.Count > 0 AndAlso TypeOf _tabControl.SelectedTab.Controls(0) Is WebBrowser Then
            Dim webBrowser As WebBrowser = DirectCast(_tabControl.SelectedTab.Controls(0), WebBrowser)
            webBrowser.Navigate(url)
        End If
    End Sub
End Class
