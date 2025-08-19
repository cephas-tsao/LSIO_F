Imports System.Data

Partial Class management_BDP080A_edit
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "BDP080A"

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '檢查資料正確性
        If Chk_data() = False Then Exit Sub

        '設定Key欄位來源
        Dim sKeyPrimary As String = txt_usr_code.Text '*修改處*

        '設定記錄資料
        Dim sSql As String = Get_SqlStr(s_prgcode, "Detail", sKeyPrimary)
        Dim sFieldName As String = ws.Get_FieldName(s_prgcode)
        Dim sOldData As String = ws.Get_OldData(sSql, s_prgcode)
        Dim sNewData As String = Get_NewData(s_prgcode)

        '資料存檔
        Dim Pgm_Type As String = txtPgmType.Text
        Dim PrgMemo As String = RepSql(Label1.Text & ":" & sKeyPrimary)

        If SaveEditData(s_prgcode, Pgm_Type, sKeyPrimary, txt_usr_name.Text, ws.EncodeString(txt_usr_pass.Text, "E"), ddl_grp_code.SelectedValue, _
                        ddl_Limit_type.SelectedValue, txt_usr_mail.Text, ddl_is_use.SelectedValue, txt_usr_memo.Text, ddl_dep_code.SelectedValue, _
                        "", "E") Then
            '儲存使用紀錄
            ws.InsUsrRec(PrgMemo, Pgm_Type, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)

            If Pgm_Type <> "MDY" Then
                '取得最新條件
                If txtSubWhere.Text <> "" Then
                    txtSubWhere.Text &= " OR usr_code='" & RepSql(sKeyPrimary) & "' "
                Else
                    txtSubWhere.Text = " WHERE usr_code='" & RepSql(sKeyPrimary) & "' "
                End If
            End If
            '存檔正確導回default頁面
            Response.Redirect("Default.aspx?subwhere=" & Replace(txtSubWhere.Text, "%", "$") & "&order=" & txtOrder.Text)
        Else
            '存檔出錯就秀出錯誤訊息
            Show_Message("儲存失敗!")
        End If
    End Sub

    Protected Function Chk_data() As Boolean
        'Server端資料檢查 自訂函數
        Chk_data = True

        '檢查重覆資料
        If txtPgmType.Text = "ADD" Then
            If Chk_RelData(s_prgcode, "", txt_usr_code.Text) = False Then
                Chk_data = False
                Show_Message("有重覆編號資料，請確認")
            End If

            Dim Dt_db As DataTable = Get_DataSet("Login", "IDno").Tables(0)
            'For i As Integer = 0 To Dt_db.Rows.Count - 1
            '    If ws.EncodeString(Dt_db.Rows(i).Item("usr_idno") & "", "D") = txt_usr_idno.Text Then
            '        Chk_data = False
            '        Show_Message("有重覆身份證字號資料，請確認")
            '        Exit For
            '    End If
            'Next
            'If Chk_RelData(s_prgcode, "IDno", txt_usr_idno.Text) = False Then
            '    Chk_data = False
            '    Show_Message("有重覆身份證字號資料，請確認")
            'End If
            '檢查身份證字號
            'If Chk_ID(txt_usr_idno.Text) = False Then
            '    Chk_data = False
            '    Show_Message("身份證字號有誤，請確認")
            'End If
        End If

        '檢查必填欄位是否空白
        If txt_usr_code.Text = "" Or txt_usr_name.Text = "" Or txt_usr_pass.Text = "" Then
            Chk_data = False
            Show_Message("必填欄位不得為空白，請確認")
        End If
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '接收default頁面變數
        If Not IsPostBack Then
            SetClientCheck()
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

            '輔助下拉式窗
            Dim sSql As String = Get_SqlStr(s_prgcode, "GrpCode")
            Set_DDL_Option("", "", ddl_grp_code, sSql, "", "")
            Set_DDL_Option("IsUse", "", ddl_is_use, "", "", "")

            Dim depStr As String = Get_SqlStr(s_prgcode, "DepCode")
            Set_DDL_Option("", "", ddl_dep_code, depStr, "", "")

            '暫不使用 by ivan
            'Set_DDL_Option("LimitLevel", "", ddl_limit_level, "", "", "")
            Set_DDL_Option("LimitType", "", ddl_Limit_type, "", "", "")

            Select Case Pgm_type
                Case "MDY", "COPY"
                    Dim Dt_tmp As DataTable = Get_DataSet(s_prgcode, "Detail", txtPrimary.Text).Tables(0)
                    If Dt_tmp.Rows.Count > 0 Then
                        txt_usr_code.Text = Trim(Dt_tmp.Rows(0).Item("usr_code") & "")
                        txt_usr_name.Text = Trim(Dt_tmp.Rows(0).Item("usr_name") & "")
                        'txt_usr_idno.Text = Rep_ID(ws.EncodeString(Trim(Dt_tmp.Rows(0).Item("usr_idno") & ""), "D"))
                        txt_usr_pass.Text = ws.EncodeString(Trim(Dt_tmp.Rows(0).Item("usr_pass") & ""), "D")
                        txt_usr_pass.Attributes.Add("value", txt_usr_pass.Text)
                        ddl_grp_code.SelectedValue = Trim(Dt_tmp.Rows(0).Item("grp_code") & "")
                        txt_usr_mail.Text = Trim(Dt_tmp.Rows(0).Item("usr_mail") & "")
                        txt_usr_memo.Text = Trim(Dt_tmp.Rows(0).Item("usr_memo") & "")
                        ddl_is_use.SelectedValue = Trim(Dt_tmp.Rows(0).Item("is_use") & "")
                        ddl_dep_code.SelectedValue = Trim(Dt_tmp.Rows(0).Item("dep_code") & "")
                        'ddl_limit_level.SelectedValue = Trim(Dt_tmp.Rows(0).Item("limit_level") & "")
                        ddl_Limit_type.SelectedValue = Trim(Dt_tmp.Rows(0).Item("Limit_type") & "")
                        If Pgm_type = "COPY" Then
                            txt_usr_code.Text = ""
                            txt_usr_code.Enabled = True
                            'txt_usr_idno.Enabled = True
                        Else
                            'txt_usr_code.Text = txtPrimary.Text
                            txt_usr_code.Enabled = False
                            'txt_usr_idno.Enabled = False
                        End If
                    End If
            End Select
        End If
        l_ErrMsg.Text = ""

        '強制大寫欄位
        txt_usr_code.Attributes("onblur") = "javascript:txt_usr_code.value = txt_usr_code.value.toUpperCase(); "
    End Sub

    Protected Sub SetClientCheck()
        '網頁伺服器端檢查點建置
        txt_usr_pass.Attributes("onblur") = "javascript:IsEmpty(this);"
        txt_usr_name.Attributes("onblur") = "javascript:IsEmpty(this);"
        txt_usr_code.Attributes("onblur") = "javascript:IsEmpty(this);"
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
End Class
