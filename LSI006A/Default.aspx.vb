
Partial Class basic_LSI006A_Default
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI006A"
    Dim s_prgname As String = ws.Get_PrgName(s_prgcode)

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Label1.Text = s_prgcode & "(" & s_prgname & ")"
            Label1.ForeColor = Drawing.Color.White
            Label1.Font.Bold = True

            txt_sender.Text = ws.Get_SysPara("sender", "")
            txt_sender_mail.Text = ws.Get_SysPara("sender_mail", "")
            txt_resender_mail.Text = ws.Get_SysPara("resender_mail", "")
        End If
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '設定記錄資料
        Dim sSql As String = Get_SqlStr(s_prgcode, "Detail")
        Dim sFieldName As String = ws.Get_FieldName(s_prgcode)
        Dim sOldData As String = ws.Get_OldData(sSql, s_prgcode)
        Dim sNewData As String = Get_NewData(s_prgcode)

        If Chk_data() = False Then Exit Sub

        If SaveEditData(s_prgcode, "MDY", txt_sender.Text, txt_sender_mail.Text, txt_resender_mail.Text) Then
            ws.InsUsrRec(s_prgname, "MDY", s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('儲存完成');</script>"))
        Else
            '存檔出錯就秀出錯誤訊息
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('儲存失敗!');</script>"))
        End If
    End Sub

    Protected Function Chk_data() As Boolean
        'Server端資料檢查 自訂函數
        Chk_data = True

        '檢查空白資料
        'If DropDownList1.SelectedValue = "" Then
        '    Chk_data = False
        '    Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('欄位不可有空白，請確認');</script>"))
        'End If
    End Function

End Class