Imports System.Web.Helpers

Partial Class SSOlogin
    Inherits PageBase
    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
        'checks database connection string and handles error if there is one
        'anticfrs.InnerHtml = AntiForgery.GetHtml().ToString()
    End Sub
End Class
