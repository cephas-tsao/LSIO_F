Imports System.Data

Partial Class basic_LSI003A_edit
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI003A"

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '自動編號
        If txtPgmType.Text <> "MDY" Then txt_ord_code.Text = "OR" & Format(Now, "yyMMdd") & Right("0000" & Get_AutoIntMax("e_paper_order", "RIGHT(ord_code, 10)", " WHERE ord_code LIKE 'OR" & Format(Now, "yyMMdd") & "%'") + 1, 4)

        '檢查資料正確性
        If Chk_data() = False Then Exit Sub

        '設定Key欄位來源
        Dim sKeyPrimary As String = txtPrimary.Text '*修改處*

        '設定記錄資料
        Dim sSql As String = Get_SqlStr(s_prgcode, "Detail", sKeyPrimary)
        Dim sFieldName As String = ws.Get_FieldName(s_prgcode)
        Dim sOldData As String = ws.Get_OldData(sSql, s_prgcode)
        Dim sNewData As String = Get_NewData(s_prgcode)

        '資料存檔
        Dim Pgm_Type As String = txtPgmType.Text
        Dim PrgMemo As String = RepSql(Label1.Text & ":" & sKeyPrimary)

        If SaveEditData(s_prgcode, Pgm_Type, txt_ord_code.Text, txt_com_name.Text, Get_Pet_List(), rbl_com_type2.SelectedValue, _
                        txt_com_type2_memo.Text, rbl_per_type.SelectedValue, rbl_job_kind.SelectedValue, txt_job_kind_memo.Text, _
                        txt_per_tel1.Text, txt_per_mail.Text, txt_pet_date.Text, ddl_Is_order.SelectedValue, Session("usr_code"), _
                        Format(Now, "yyyy/MM/dd")) Then

            '儲存使用紀錄
            ws.InsUsrRec(PrgMemo, Pgm_Type, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)

            If Pgm_Type <> "MDY" Then
                '取得最新條件
                If txtSubWhere.Text <> "" Then
                    txtSubWhere.Text &= " OR ord_code='" & RepSql(sKeyPrimary) & "' "
                Else
                    txtSubWhere.Text = " WHERE ord_code='" & RepSql(sKeyPrimary) & "' "
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

        ''檢查重覆資料
        If txtPgmType.Text = "ADD" Or txtPgmType.Text = "COPY" Then
            If Chk_RelData(s_prgcode, "", txt_ord_code.Text) = False Then
                Chk_data = False
                Show_Message("有重覆參數名稱，請確認")
            End If
        End If

        '檢查資料正確性 
        If Chk_Tel(txt_per_tel1.Text) = False Then
            Chk_data = False
            Show_Message("行動電話格式不正確，請確認")
        End If
        If Chk_Email(txt_per_mail.Text) = False Then
            Chk_data = False
            Show_Message("Mail格式不正確，請確認")
        End If
        If Chk_DateForm(txt_pet_date.Text) = False Then
            Chk_data = False
            Show_Message("日期格式不正確，請確認")
        End If

        '檢查必填欄位是否空白
        If txt_com_name.Text = "" Or rbl_per_type.SelectedValue = "" Or txt_per_tel1.Text = "" Or Get_Pet_List() = "" _
                                  Or rbl_job_kind.SelectedValue = "" Then
            Chk_data = False
            Show_Message("必填欄位不得為空白，請確認")
        End If
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '接收default頁面變數
        If Not IsPostBack Then
            SetClientCheck()
            Set_RBL_Option("job_kind", "", rbl_job_kind, "", "", "")
            Set_RBL_Option("com_type2", "", rbl_com_type2, "", "", "")
            Set_RBL_Option("per_type", "", rbl_per_type, "", "", "")
            Set_DDL_Option("YorN", "", ddl_Is_order, "", "", "")
            Set_DDL_Option("com_type22", "", ddl_com_type22, "", "----", "")
            Set_DDL_Option("com_type23", "", ddl_com_type23, "", "----", "")
            Set_DDL_Option("com_type24", "", ddl_com_type24, "", "----", "")
            Set_DDL_Option("com_type25", "", ddl_com_type25, "", "----", "")

            chk_com_type1_1.Text = "勞工安全衛生人員"
            chk_com_type1_2.Text = "營造作業主管"
            chk_com_type1_3.Text = "有害作業主管"
            chk_com_type1_4.Text = "危險性機械(設備)操作人員"
            chk_com_type1_5.Text = "其他相關人員"

            'txt_pet_date.Attributes.Add("ReadOnly", "ReadOnly")

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

            Select Case Pgm_type
                Case "ADD"
                    txt_pet_date.Text = Format(Now, "yyyy/MM/dd")
                Case "MDY", "COPY"
                    Dim Dt_tmp As DataTable = Get_DataSet(s_prgcode, "Detail", txtPrimary.Text).Tables(0)
                    If Dt_tmp.Rows.Count > 0 Then
                        txt_ord_code.Text = Trim(Dt_tmp.Rows(0).Item("ord_code") & "")
                        txt_pet_date.Text = Trim(Dt_tmp.Rows(0).Item("pet_date") & "")
                        txt_per_tel1.Text = Trim(Dt_tmp.Rows(0).Item("per_tel1") & "")
                        rbl_per_type.SelectedValue = Trim(Dt_tmp.Rows(0).Item("per_type") & "")
                        rbl_job_kind.SelectedValue = Trim(Dt_tmp.Rows(0).Item("job_kind") & "")
                        txt_com_name.Text = Trim(Dt_tmp.Rows(0).Item("com_name") & "")
                        txt_per_mail.Text = Trim(Dt_tmp.Rows(0).Item("per_mail") & "")
                        txt_job_kind_memo.Text = Trim(Dt_tmp.Rows(0).Item("job_kind_memo") & "")
                        txt_com_type2_memo.Text = Trim(Dt_tmp.Rows(0).Item("com_type2_memo") & "")
                        ddl_com_type22.SelectedValue = Trim(Dt_tmp.Rows(0).Item("com_type1") & "")
                        ddl_com_type23.SelectedValue = Trim(Dt_tmp.Rows(0).Item("com_type1") & "")
                        ddl_com_type24.SelectedValue = Trim(Dt_tmp.Rows(0).Item("com_type1") & "")
                        ddl_com_type25.SelectedValue = Trim(Dt_tmp.Rows(0).Item("com_type1") & "")
                        ddl_Is_order.SelectedValue = Trim(Dt_tmp.Rows(0).Item("Is_order") & "")
                        rbl_com_type2.SelectedValue = Trim(Dt_tmp.Rows(0).Item("com_type2") & "")
                        Set_com_type(Trim(Dt_tmp.Rows(0).Item("com_type1") & ""))
                    End If

                    If Pgm_type = "COPY" Then
                        txt_ord_code.Text = ""
                        txt_ord_code.Enabled = False
                    Else
                        txt_ord_code.Enabled = False
                    End If
            End Select
        End If

        img_date1.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_pet_date")
        l_ErrMsg.Text = ""
    End Sub

    Protected Sub SetClientCheck()
        '網頁伺服器端檢查點建置
        'txt_par_value.Attributes("onblur") = "javascript:IsEmpty(this);"
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

    Protected Function Get_Pet_List() As String
        '儲存勞動安全衛生相關人員
        Dim sStr As String = ""
        Dim bFirst As Boolean = True

        If chk_com_type1_1.Checked = True Then
            bFirst = False
            sStr &= "1A"
        End If
        If chk_com_type1_2.Checked = True Then
            If bFirst <> True Then sStr &= "|"
            bFirst = False
            sStr &= ddl_com_type22.SelectedValue
        End If
        If chk_com_type1_3.Checked = True Then
            If bFirst <> True Then sStr &= "|"
            bFirst = False
            sStr &= ddl_com_type23.SelectedValue
        End If
        If chk_com_type1_4.Checked = True Then
            If bFirst <> True Then sStr &= "|"
            bFirst = False
            sStr &= ddl_com_type24.SelectedValue
        End If
        If chk_com_type1_5.Checked = True Then
            If bFirst <> True Then sStr &= "|"
            bFirst = False
            sStr &= ddl_com_type25.SelectedValue
        End If

        Return sStr
    End Function
    ' 讀取勞動安全衛生相關人員
    Protected Sub Set_com_type(ByVal p_str As String)
        If Trim(p_str & "") <> "" Then

            If InStr(p_str, "1") > 0 Then
                chk_com_type1_1.Checked = True
            End If

            If InStr(p_str, "2") > 0 Then
                chk_com_type1_2.Checked = True
                ddl_com_type22.SelectedValue = Mid(p_str, InStr(p_str, "2"), 2)
            End If

            If InStr(p_str, "3") > 0 Then
                chk_com_type1_3.Checked = True
                ddl_com_type23.SelectedValue = Mid(p_str, InStr(p_str, "3"), 2)
            End If

            If InStr(p_str, "4") > 0 Then
                chk_com_type1_4.Checked = True
                ddl_com_type24.SelectedValue = Mid(p_str, InStr(p_str, "4"), 2)
            End If

            If InStr(p_str, "5") > 0 Then
                chk_com_type1_5.Checked = True
                ddl_com_type25.SelectedValue = Mid(p_str, InStr(p_str, "5"), 2)
            End If

        End If
    End Sub
End Class
