'@2023 Newspoint Software
'@2023 Newtonsoft.Json.Linq
Imports System.Drawing
Imports System.Net
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Module interop
    Public Class GitHub
        Public Shared Function GetKey()

        End Function
        Public Shared Function GetRepository()

        End Function
        Public Shared Function GetDisplayRepo(repositoryData As String) As Panel
            Dim repositories As JArray = JsonConvert.DeserializeObject(Of JArray)(repositoryData)

            Dim repoPanel As New Panel()
            repoPanel.BackColor = Color.White
            repoPanel.BorderStyle = BorderStyle.FixedSingle
            repoPanel.Padding = New Padding(10)
            repoPanel.Margin = New Padding(5)

            For Each repo As JObject In repositories
                Dim repoName As String = repo("name").ToString()
                Dim repoDescription As String = repo("description").ToString()

                Dim repoNameLabel As New Label()
                repoNameLabel.Text = repoName
                repoPanel.Controls.Add(repoNameLabel)

                Dim repoDescriptionLabel As New Label()
                repoDescriptionLabel.Text = repoDescription
                repoPanel.Controls.Add(repoDescriptionLabel)
            Next

            Return repoPanel
        End Function
        Public Shared Function Config()

        End Function
        Public Shared Function GetInfo()

        End Function
        Public Shared Function Update()

        End Function
        Public Shared Function Test()

        End Function
    End Class
End Module
