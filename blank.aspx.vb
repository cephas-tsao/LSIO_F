
Partial Class blank
    Inherits PageBase

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) Then
            BrowserName.Text = Request.Browser.Browser
            BrowserVersion.Text = Request.Browser.Version
        End If
    End Sub
End Class
