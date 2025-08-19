Imports System.Data

Partial Class basic_OTB020A_edit
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "BDP190A"

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '檢查資料正確性
        If Chk_data() = False Then Exit Sub

        '設定Key欄位來源
        Dim sKeyPrimary As String = txt_bull_code.Text '*修改處*

        '設定記錄資料
        Dim sSql As String = Get_SqlStr(s_prgcode, "Detail", sKeyPrimary)
        Dim sFieldName As String = ws.Get_FieldName(s_prgcode)
        Dim sOldData As String = ws.Get_OldData(sSql, s_prgcode)
        Dim sNewData As String = Get_NewData(s_prgcode)

        '資料存檔
        Dim Pgm_Type As String = txtPgmType.Text
        Dim PrgMemo As String = RepSql(Label1.Text & ":" & sKeyPrimary)

        '自動編號
        'If txtPgmType.Text <> "MDY" Then txt_bull_code.Text = Format(Now, "yyMMdd") & Right("00" & Get_AutoIntMax("OTB040", "tn_code", " WHERE tn_code LIKE '" & Format(Now, "yyMMdd") & "%'") + 1, 2)

        If SaveEditData(s_prgcode, Pgm_Type, txt_bull_code.Text, txt_usr_name.Text, txt_usr_mail.Text, txt_theme.Text, _
                        txt_bull_con.Text, txt_ins_date.Text, txt_ins_time.Text, txt_usr_ip.Text) Then
            Show_Message("存檔成功")
            '儲存使用紀錄
            ws.InsUsrRec(PrgMemo, Pgm_Type, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)

            If Pgm_Type <> "MDY" Then
                txtPgmType.Text = "MDY"
                detail_edit.setKeyName(txt_bull_code.Text)
                txt_bull_code.Enabled = False
                Panel1.Visible = True

                '取得最新條件
                If txtSubWhere.Text <> "" Then
                    txtSubWhere.Text &= " OR bull_code='" & RepSql(sKeyPrimary) & "' and bull_type='0'"
                Else
                    txtSubWhere.Text = " WHERE bull_code='" & RepSql(sKeyPrimary) & "' and bull_type='0'"
                End If
            End If
        Else
            '存檔出錯就秀出錯誤訊息
            Show_Message("儲存失敗!")
        End If
    End Sub

    Protected Function Chk_data() As Boolean
        'Server端資料檢查 自訂函數
        Chk_data = True

        ''檢查重覆資料
        'If txtPgmType.Text = "ADD" Then
        '    If Chk_RelData(s_prgcode, "", txt_tn_code.Text) = False Then
        '        Chk_data = False
        '        Show_Message("有重覆資料，請確認")
        '    End If
        'End If

        '檢查必填欄位是否空白
        'If txt_usr_name.Text = "" Or txt_theme.Text = "" Then
        '    Chk_data = False
        '    Show_Message("必填欄位不得為空白，請確認")
        'End If

        '檢查日期格式是否正確
        If Chk_DateForm(txt_ins_date.Text) = False Then
            Chk_data = False
            Show_Message("日期格式錯誤，請確認")
        End If
        If Chk_TimeForm(txt_ins_time.Text, ":") = False Then
            Chk_data = False
            Show_Message("時間格式錯誤，請確認")
        End If
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '接收default頁面變數
        If Not IsPostBack Then
            Call SetClientCheck()
            Set_DDL_Option("IsUse", "", ddl_is_hid, "", "", "")

            '標頭
            Label1.Text = CType(PreviousPage.FindControl("Label1"), Label).Text
            '編輯模式
            txtPgmType.Text = CType(PreviousPage.FindControl("txtPgmType"), TextBox).Text
            '首頁條件
            txtSubWhere.Text = CType(PreviousPage.FindControl("txtsubwhere"), TextBox).Text
            txtOrder.Text = CType(PreviousPage.FindControl("txtOrder"), TextBox).Text

            '非新增模式則預設取得資料
            Dim Pgm_type As String = txtPgmType.Text
            txtPrimary.Text = CType(PreviousPage.FindControl("txtPrimary"), TextBox).Text
            txt_bull_code.Attributes.Add("ReadOnly", "ReadOnly")

            Select Case Pgm_type
                Case "ADD"
                    txt_theme.Text = Format(Now, "yyyy/MM/dd")
                Case "MDY", "COPY"
                    Dim Dt_tmp As DataTable = Get_DataSet(s_prgcode, "Detail", txtPrimary.Text).Tables(0)
                    If Dt_tmp.Rows.Count > 0 Then
                        If Pgm_type = "COPY" Then
                            txt_bull_code.Text = ""
                            txt_bull_code.Enabled = True
                        Else
                            txt_bull_code.Text = txtPrimary.Text
                            txt_bull_code.Enabled = False
                            detail_edit.setKeyName(txtPrimary.Text)
                            Panel1.Visible = True
                        End If
                        txt_usr_name.Text = Trim(Dt_tmp.Rows(0).Item("usr_name") & "")
                        txt_theme.Text = Trim(Dt_tmp.Rows(0).Item("theme") & "")
                        txt_ins_date.Text = Trim(Dt_tmp.Rows(0).Item("ins_date") & "")
                        txt_bull_con.Text = Trim(Dt_tmp.Rows(0).Item("bull_con") & "")
                        txt_usr_ip.Text = Trim(Dt_tmp.Rows(0).Item("usr_ip") & "")
                        txt_usr_mail.Text = Trim(Dt_tmp.Rows(0).Item("usr_mail") & "")
                        txt_ins_time.Text = Trim(Dt_tmp.Rows(0).Item("ins_time") & "")
                        ddl_is_hid.SelectedValue = Trim(Dt_tmp.Rows(0).Item("is_hid") & "")
                    End If
            End Select
        End If
        Button1.Visible = False
        l_ErrMsg.Text = ""
    End Sub

    Protected Sub SetClientCheck()
        '網頁伺服器端檢查點建置
        'txtPrimary.Attributes("onblur") = "javascript:IsEmpty(this);"
    End Sub

    Protected Sub Bt_BackUp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Bt_BackUp.Click
        Response.Redirect("Default.aspx?subwhere=" & Replace(txtSubWhere.Text, "%", "$") & "&order=" & txtOrder.Text)
    End Sub

    Protected Sub Show_Message(ByVal sMsg As String)
        Dim ErrMsg_Type As String = ws.Get_SysPara("ErrMsg_Type", "0")

        If ErrMsg_Type = "0" Or ErrMsg_Type = "2" Then
            '彈出式對話方塊
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('" & sMsg & "');</script>"))
        End If

        If ErrMsg_Type = "1" Or ErrMsg_Type = "2" Then
            '下方紅字訊息
            l_ErrMsg.Text &= sMsg & "<br />"
        End If
    End Sub

    Protected Sub ddl_is_hid_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddl_is_hid.SelectedIndexChanged
        Dim Pgm_Type As String = "MDY"
        '設定Key欄位來源
        Dim sKeyPrimary As String = txt_bull_code.Text '*修改處*

        '設定記錄資料
        Dim sSql As String = Get_SqlStr(s_prgcode, "Detail", sKeyPrimary)

        '設定記錄資料
        Dim sFieldName As String = ws.Get_FieldName(s_prgcode)
        Dim sOldData As String = ws.Get_OldData(sSql, s_prgcode)
        Dim sNewData As String = Get_NewData(s_prgcode)
        Dim PrgMemo As String = RepSql(Label1.Text & ":" & sKeyPrimary)

        If SaveEditData(s_prgcode, Pgm_Type, txt_bull_code.Text, ddl_is_hid.SelectedValue) Then
            Show_Message("存檔成功")
            '儲存使用紀錄
            ws.InsUsrRec(PrgMemo, Pgm_Type, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)
        Else
            '存檔出錯就秀出錯誤訊息
            Show_Message("存檔失敗")
        End If
    End Sub
End Class

