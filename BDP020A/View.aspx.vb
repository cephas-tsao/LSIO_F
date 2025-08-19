
Partial Class management_BDP019A_View
    Inherits PageBase

    Protected Function Chk_data() As Boolean
        'Server端資料檢查 自訂函數
        Chk_data = True

        ' ''檢查重覆資料
        ''If txt_pgm_type.Text = "ADD" Then
        ''    If Chk_RelData("select * from BDP190 where bdp190='" & txt_dep_code.Text & "'") = False Then
        ''        Chk_data = False
        ''        Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('有重覆編號資料，請確認');</script>"))
        ''    End If
        ''End If

    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '接收default頁面變數
        If Not IsPostBack Then
            Dim sNo As String = Request("no")
            Dim oDtDb As Object
            Dim sSQL As String = ""
            '標頭
            'Label1.Text = CType(PreviousPage.FindControl("Label1"), Label).Text
            Label1.Text = Get_PrgName("BDP019A")
            '編輯模式
            'txt_pgm_type.Text = CType(PreviousPage.FindControl("txtPgmType"), TextBox).Text
            txt_pgm_type.Text = Request("txtPgmType")
            '首頁條件
            'txtSubWhere.Text = CType(PreviousPage.FindControl("txtsubwhere"), TextBox).Text
            txtSubWhere.Text = Request("txtsubwhere")

            ImageButton1.Visible = False
            Label2.Visible = False

            sSQL = " SELECT * FROM BDP190 " & _
                   " LEFT JOIN bdp080 ON bdp080.usr_code = BDP190.usr_code " & _
                   " WHERE bdp190 = '" & sNo & "' "
            oDtDb = Me.Get_DataTable(sSQL)

            If oDtDb.rows.count > 0 Then
                Label3.Text = Trim(oDtDb.rows(0)("theme").ToString)
                Label4.Text = Trim(oDtDb.rows(0)("usr_name").ToString)
                Label5.Text = Trim(oDtDb.rows(0)("bull_con").ToString)
                Label6.Text = "有效期限：" & Trim(oDtDb.rows(0)("bull_date").ToString) & " ~ " & Trim(oDtDb.rows(0)("exp_date").ToString)
                If Trim(oDtDb.rows(0)("file1")) <> "" Then
                    ImageButton1.Visible = True
                    TextBox1.Text = Trim(oDtDb.rows(0)("file1").ToString)
                Else
                    Label2.Visible = True
                End If
            End If
            oDtDb = Nothing
        End If
    End Sub

    Protected Sub ImageButton1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton1.Click

        If IO.File.Exists(TextBox1.Text) Then
        Else
            TextBox1.Text = "0"
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('無檔案，請洽公告人');</script>"))
            Exit Sub
        End If


        Dim oFile As System.IO.FileInfo = New System.IO.FileInfo(TextBox1.Text)
        Response.Clear() ' // clear the current output content from the buffer
        Response.AppendHeader("Content-Disposition", "attachment; filename=" & Server.UrlEncode(oFile.Name))
        Response.AppendHeader("Content-Length", oFile.Length.ToString())
        Response.ContentType = "application/octet-stream"
        Response.WriteFile(oFile.FullName)
    End Sub
End Class