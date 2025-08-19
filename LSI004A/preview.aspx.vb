
Partial Class basic_LSI004A_preview
    Inherits PageBase

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Response.CacheControl = "no-cache"
        Response.AddHeader("Pragma", "no-cache")
        Response.Expires = -1
        Response.Buffer = True

        Dim ws As New WebReference.LSIO_WebService
        Dim sPaperCode As String = RepSql(Request.QueryString("code"))
        l_paper_html.Text = ws.Get_PaperHtml(sPaperCode)
    End Sub
End Class
