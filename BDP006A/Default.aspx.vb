Imports System.Data

Partial Class management_BDP006A_Default
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "BDP006A"
    Dim s_prgname As String = ws.Get_PrgName(s_prgcode)

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Label1.Text = "修改個人基本資料"
            Label1.ForeColor = Drawing.Color.White
            Label1.Font.Bold = True
            txt_UsrNewPW.Attributes.Add("value", txt_UsrNewPW.Text)
            txt_UsrNewPW2.Attributes.Add("value", txt_UsrNewPW2.Text)
            txt_UsrOldPW.Attributes.Add("value", txt_UsrOldPW.Text)
            txt_UsrOldPW.TextMode = TextBoxMode.Password
            'txt_UsrPW.Attributes.Add("value", txt_UsrPW.Text)
            'txt_UsrPW.TextMode = TextBoxMode.Password
            txt_UsrNewPW.TextMode = TextBoxMode.Password
            txt_UsrNewPW2.TextMode = TextBoxMode.Password
            txt_UsrCode.Text = Session("usr_code")

            Dim oDtDb As DataTable = Get_DataSet(s_prgcode, "", Session("usr_code")).Tables(0)

            If oDtDb.Rows.Count > 0 Then
                txt_UsrPW.Text = ws.EncodeString(oDtDb.Rows(0)("usr_pass").ToString, "D")
            End If

            txt_UsrPW.Enabled = False
        End If
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sPrgMemo As String = "修改個人基本資料"
        
        txt_UsrNewPW.Attributes.Add("value", txt_UsrNewPW.Text)
        txt_UsrNewPW2.Attributes.Add("value", txt_UsrNewPW2.Text)

        If txt_UsrPW.Text <> (Trim(txt_UsrOldPW.Text) & "") Then
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('舊密碼輸入錯誤');</script>"))
            Exit Sub
        ElseIf txt_UsrNewPW.Text <> txt_UsrNewPW2.Text Or txt_UsrNewPW.Text = "" Or txt_UsrNewPW2.Text = "" Then
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('密碼輸入錯誤');</script>"))
            Exit Sub
        End If

        If SaveEditData(s_prgcode, "MDY", txt_UsrCode.Text, ws.EncodeString(txt_UsrNewPW2.Text, "E")) Then
            ws.InsUsrRec(s_prgname, "MDY", s_prgcode, Session("usr_code"), Session("usr_ip"), "", "", "")
            Response.Write("<script>alert('儲存完成')</script>")
            '存檔正確導回default頁面
            'Response.Redirect("~/blank.aspx")
            Response.Write("<script>setTimeout(location.href=""../../blank.aspx"",1000)</script>")
        Else
            '存檔出錯就秀出錯誤訊息
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('儲存失敗!');</script>"))
        End If
    End Sub

End Class