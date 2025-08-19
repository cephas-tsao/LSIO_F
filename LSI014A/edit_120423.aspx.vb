Imports System.Data

Partial Class basic_LSI014A_edit
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI014A"

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '自動編號
        If txtPgmType.Text <> "MDY" Then txt_dan_code.Text = "DA" & Format(Now, "yyMMdd") & Right("0000" & Get_AutoIntMax("Danger_work", "RIGHT(dan_code, 10)", " WHERE dan_code LIKE 'DA" & Format(Now, "yyMMdd") & "%'") + 1, 4)

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

        If SaveEditData(s_prgcode, Pgm_Type, txt_dan_code.Text, txt_com_name.Text, txt_pet_name.Text, txt_pet_date.Text, txt_pet_time.Text _
                        , txt_pet_tel1.Text, txt_att_name1.Text, txt_att_tel1.Text, txt_att_name2.Text, txt_att_tel2.Text, txt_att_add1.Text _
                        , ddl_att_area.SelectedValue, ddl_att_road.SelectedValue, txt_bulid_name.Text, txt_bulid_code.Text, txt_work_date_s.Text _
                        , txt_work_date_e.Text, ddl_is_night.SelectedValue, txt_work_floor_s.Text, txt_work_floor_e.Text, ddl_work_type1.SelectedValue _
                        , ddl_work_type2.SelectedValue, ddl_work_type3.SelectedValue, ddl_work_type4.SelectedValue, ddl_work_type5.SelectedValue _
                        , ddl_work_type6.SelectedValue, ddl_work_type7.SelectedValue, ddl_work_type8.SelectedValue, ddl_work_type9.SelectedValue _
                        , txt_work_type_memo.Text, ddl_use_tool1.SelectedValue, ddl_use_tool2.SelectedValue, ddl_use_tool3.SelectedValue, ddl_use_tool4.SelectedValue _
                        , ddl_use_tool5.SelectedValue, txt_use_tool3_memo.Text, txt_use_tool5_memo.Text, Get_CKL_Click(ckl_save_tool), txt_save_tool_memo.Text _
                        , Get_CKL_Click(ckl_dan_memo), txt_dan_memo_memo.Text, txt_dan_mail.Text) Then


            '儲存使用紀錄
            ws.InsUsrRec(PrgMemo, Pgm_Type, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)

            If Pgm_Type <> "MDY" Then
                '取得最新條件
                If txtSubWhere.Text <> "" Then
                    txtSubWhere.Text &= " OR dan_code='" & RepSql(sKeyPrimary) & "' "
                Else
                    txtSubWhere.Text = " WHERE dan_code='" & RepSql(sKeyPrimary) & "' "
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
            If Chk_RelData(s_prgcode, "", txt_dan_code.Text) = False Then
                Chk_data = False
                Show_Message("有重覆參數名稱，請確認")
            End If

        End If
        '檢查資料正確性       
        If Chk_DateForm(txt_pet_date.Text) = False Or Chk_DateForm(txt_work_date_s.Text) = False Or Chk_DateForm(txt_work_date_e.Text) = False Then
            Chk_data = False
            Show_Message("日期格式不正確，請確認")
        End If

        If Chk_Email(txt_dan_mail.Text) = False Then
            Chk_data = False
            Show_Message("日期格式不正確，請確認")
        End If

    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '接收default頁面變數
        If Not IsPostBack Then
            SetClientCheck()
            Set_DDL_Option("ZipCode", "", ddl_att_area, "", "---請選擇---", "")
            Set_DDL_Option("ZipCode", "", ddl_att_road, "", "---請選擇---", "")
            Set_CKL_Option("save_tool", "", ckl_save_tool, "", "", "")
            Set_CKL_Option("dan_memo", "", ckl_dan_memo, "", "", "")
            Set_DDL_Option("YorN", "", ddl_is_night, "", "---請選擇---", "")
            Set_DDL_Option("work_type1", "", ddl_work_type1, "", "---請選擇---", "")
            Set_DDL_Option("work_type2", "", ddl_work_type2, "", "---請選擇---", "")
            Set_DDL_Option("work_type3", "", ddl_work_type3, "", "---請選擇---", "")
            Set_DDL_Option("work_type4", "", ddl_work_type4, "", "---請選擇---", "")
            Set_DDL_Option("work_type5", "", ddl_work_type5, "", "---請選擇---", "")
            Set_DDL_Option("work_type6", "", ddl_work_type6, "", "---請選擇---", "")
            Set_DDL_Option("work_type7", "", ddl_work_type7, "", "---請選擇---", "")
            Set_DDL_Option("work_type8", "", ddl_work_type8, "", "---請選擇---", "")
            Set_DDL_Option("work_type9", "", ddl_work_type9, "", "---請選擇---", "")
            Set_DDL_Option("use_tool1", "", ddl_use_tool1, "", "---請選擇---", "")
            Set_DDL_Option("use_tool2", "", ddl_use_tool2, "", "---請選擇---", "")
            Set_DDL_Option("use_tool3", "", ddl_use_tool3, "", "---請選擇---", "")
            Set_DDL_Option("use_tool4", "", ddl_use_tool4, "", "---請選擇---", "")
            Set_DDL_Option("use_tool5", "", ddl_use_tool5, "", "---請選擇---", "")

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
                Case "MDY", "COPY"
                    Dim Dt_tmp As DataTable = Get_DataSet(s_prgcode, "Detail", txtPrimary.Text).Tables(0)
                    If Dt_tmp.Rows.Count > 0 Then
                        txt_dan_code.Text = Trim(Dt_tmp.Rows(0).Item("dan_code") & "")
                        txt_com_name.Text = Trim(Dt_tmp.Rows(0).Item("com_name") & "")
                        txt_pet_name.Text = Trim(Dt_tmp.Rows(0).Item("pet_name") & "")
                        txt_pet_tel1.Text = Trim(Dt_tmp.Rows(0).Item("pet_tel1") & "")
                        txt_pet_date.Text = Trim(Dt_tmp.Rows(0).Item("pet_date") & "")
                        txt_pet_time.Text = Trim(Dt_tmp.Rows(0).Item("pet_time") & "")
                        txt_att_add1.Text = Trim(Dt_tmp.Rows(0).Item("att_add1") & "")
                        ddl_att_area.SelectedValue = Trim(Dt_tmp.Rows(0).Item("att_area") & "")
                        ddl_att_road.SelectedValue = Trim(Dt_tmp.Rows(0).Item("att_road") & "")
                        txt_bulid_code.Text = Trim(Dt_tmp.Rows(0).Item("bulid_code") & "")
                        txt_bulid_name.Text = Trim(Dt_tmp.Rows(0).Item("bulid_name") & "")
                        txt_work_floor_s.Text = Trim(Dt_tmp.Rows(0).Item("work_floor_s") & "")
                        txt_work_floor_e.Text = Trim(Dt_tmp.Rows(0).Item("work_floor_e") & "")
                        txt_work_date_s.Text = Trim(Dt_tmp.Rows(0).Item("work_date_s") & "")
                        txt_work_date_e.Text = Trim(Dt_tmp.Rows(0).Item("work_date_e") & "")
                        txt_att_name1.Text = Trim(Dt_tmp.Rows(0).Item("att_name1") & "")
                        txt_att_tel1.Text = Trim(Dt_tmp.Rows(0).Item("att_tel1") & "")
                        txt_att_name2.Text = Trim(Dt_tmp.Rows(0).Item("att_name2") & "")
                        txt_att_tel2.Text = Trim(Dt_tmp.Rows(0).Item("att_tel2") & "")
                        ddl_is_night.SelectedValue = Trim(Dt_tmp.Rows(0).Item("is_night") & "")
                        Set_CKL_Click(Trim(Dt_tmp.Rows(0).Item("save_tool") & ""), ckl_save_tool, "save_tool")
                        ddl_work_type1.SelectedValue = Trim(Dt_tmp.Rows(0).Item("work_type1") & "")
                        ddl_work_type2.SelectedValue = Trim(Dt_tmp.Rows(0).Item("work_type2") & "")
                        ddl_work_type3.SelectedValue = Trim(Dt_tmp.Rows(0).Item("work_type3") & "")
                        ddl_work_type4.SelectedValue = Trim(Dt_tmp.Rows(0).Item("work_type4") & "")
                        ddl_work_type5.SelectedValue = Trim(Dt_tmp.Rows(0).Item("work_type5") & "")
                        ddl_work_type6.SelectedValue = Trim(Dt_tmp.Rows(0).Item("work_type6") & "")
                        ddl_work_type7.SelectedValue = Trim(Dt_tmp.Rows(0).Item("work_type7") & "")
                        ddl_work_type8.SelectedValue = Trim(Dt_tmp.Rows(0).Item("work_type8") & "")
                        ddl_work_type9.SelectedValue = Trim(Dt_tmp.Rows(0).Item("work_type9") & "")
                        txt_work_type_memo.Text = Trim(Dt_tmp.Rows(0).Item("work_type_memo") & "")
                        ddl_use_tool1.SelectedValue = Trim(Dt_tmp.Rows(0).Item("use_tool1") & "")
                        ddl_use_tool2.SelectedValue = Trim(Dt_tmp.Rows(0).Item("use_tool2") & "")
                        ddl_use_tool3.SelectedValue = Trim(Dt_tmp.Rows(0).Item("use_tool3") & "")
                        ddl_use_tool4.SelectedValue = Trim(Dt_tmp.Rows(0).Item("use_tool4") & "")
                        ddl_use_tool5.SelectedValue = Trim(Dt_tmp.Rows(0).Item("use_tool5") & "")
                        txt_use_tool3_memo.Text = Trim(Dt_tmp.Rows(0).Item("use_tool3_memo") & "")
                        txt_use_tool5_memo.Text = Trim(Dt_tmp.Rows(0).Item("use_tool5_memo") & "")
                        Set_CKL_Click(Trim(Dt_tmp.Rows(0).Item("dan_memo") & ""), ckl_dan_memo, "dan_memo")
                        txt_save_tool_memo.Text = Trim(Dt_tmp.Rows(0).Item("save_tool_memo") & "")
                        txt_dan_memo_memo.Text = Trim(Dt_tmp.Rows(0).Item("dan_memo_memo") & "")
                        txt_dan_mail.Text = Trim(Dt_tmp.Rows(0).Item("dan_mail") & "")
                    End If

                    If Pgm_type = "COPY" Then
                        txt_dan_code.Text = ""
                        txt_dan_code.Enabled = False
                    Else
                        txt_dan_code.Enabled = False
                    End If
            End Select
        End If
        img_date1.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_pet_date")
        img_date2.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_work_date_s")
        img_date3.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_work_date_e")
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
End Class
