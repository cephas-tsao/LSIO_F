Imports System.Data

Partial Class basic_LSI002A_edit
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI002A"

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '自動編號
        If txtPgmType.Text <> "MDY" Then txt_cmp_code.Text = "AP" & Format(Now, "yyMMdd") & Right("0000" & Get_AutoIntMax("Mail_Complain", "RIGHT(cmp_code, 10)", " WHERE cmp_code LIKE 'AP" & Format(Now, "yyMMdd") & "%'") + 1, 4)
        If txtPgmType.Text <> "MDY" Then
            ins_date.Text = Format(Now, "yyyy/MM/dd")
            cmp_date.Text = DateTime.Now.ToString("HH:mm:ss")
            cmp_pass.Text = Session("ComplainA_VerCode")
        End If
        '檢查資料正確性
        If Chk_data() = False Then Exit Sub

        '設定Key欄位來源
        Dim sKeyPrimary As String = txtPrimary.Text '*修改處*

        '設定記錄資料
        Dim sSql As String = Get_SqlStr(s_prgcode, "Detail", sKeyPrimary)
        Dim sFieldName As String = ws.Get_FieldName(s_prgcode)
        Dim sOldData As String = ws.Get_OldData(sSql, s_prgcode)
        Dim sNewData As String = Get_NewData(s_prgcode)

        '是否為代他人申訴
        If rb_entrustY.Checked = True Then txt_cmp_entrust.Text = ("1" & "#" & txt_cmp_depart.Text & "#" & txt_cmp_job.Text)
        If rb_entrustN.Checked = True Then txt_cmp_entrust.Text = ("0" & "#" & txt_cmp_depart.Text & "#" & txt_cmp_job.Text)

        Dim showtype As String = ""
        If rblist.Checked = True Then
            cmp_idclass.Text = "1"
            showtype = "事業單位所屬勞工(目前尚在職者)</br>" & "到職日期：" & cmp_sdate.Text & "</br>擔任職務及工作內容：" & jobcontent.Text
        End If
        If rblist2.Checked = True Then
            cmp_idclass.Text = "2"
            showtype = "之前曾經是事業單位勞工，離職日期：" & cmp_edate.Text
        End If
        If rblist3.Checked = True Then
            cmp_idclass.Text = "3"
            showtype = "民眾(非屬前二項者)"
        End If

        If com_Contract_chk.Checked = True Then
            com_Contract_chk.Text = "Y"
        Else
            com_Contract_chk.Text = "N"
        End If

        If com_relax_chk.Checked = True Then
            com_relax_chk.Text = "Y"
        Else
            com_relax_chk.Text = "N"
        End If

        If com_Welfare_chk.Checked = True Then
            com_Welfare_chk.Text = "Y"
        Else
            com_Welfare_chk.Text = "N"
        End If

        Dim dtype As String = ""
        If DN.Checked = True Then
            com_disputetype.Text = "N"
            dtype = "否"
        End If
        If DY.Checked = True Then
            com_disputetype.Text = "Y"
            dtype = "是，請註明日期及機關" & com_disputecon.Text
        End If

        Dim mtype As String = ""
        If MN.Checked = True Then
            com_Mediationtype.Text = "N"
            mtype = "否"
        End If
        If MYY.Checked = True Then
            com_Mediationtype.Text = "Y"
            mtype = "是，請註明日期及機關" & com_Mediationcon.Text
        End If

        Dim stype As String = ""
        If SN.Checked = True Then
            com_sectype.Text = "N"
            stype = "目前仍在職，請將我的身分保密"
        End If
        If SY.Checked = True Then
            com_sectype.Text = "Y"
            stype = "我願意將姓名提供給事業單位，以利檢查時協助追償受損權益"
        End If

        If msex.Checked = True Then
            com_sex.Text = "m"
        End If

        If fsex.Checked = True Then
            com_sex.Text = "f"
        End If

        '資料存檔
        Dim Pgm_Type As String = txtPgmType.Text
        Dim PrgMemo As String = RepSql(Label1.Text & ":" & sKeyPrimary)

        If SaveEditData(s_prgcode, Pgm_Type, txt_cmp_code.Text, txt_cmp_name.Text, txt_cmp_tel1.Text, txt_cmp_email.Text, "", _
                        txt_com_name.Text, txt_com_add.Text, txt_com_idno.Text, txt_cmp_memo.Text, ddl_cmp_type.SelectedValue, "", "", "", "", "", _
                        Session("usr_code"), ins_date.Text, ddl_att_area.SelectedValue, ddl_att_road.SelectedValue, "Y", cmp_date.Text, _
                        cmp_idclass.Text, cmp_sdate.Text, jobcontent.Text, cmp_edate.Text, com_type.Text, ddl_att_area1.SelectedValue, ddl_att_road1.SelectedValue, _
                        txt_com_add1.Text, Get_CKL_Click(com_Contract), com_Contract_chk.Text, com_Contract_other.Text, Get_CKL_Click(com_paymon), Get_CKL_Click(com_relax), _
                        com_relax_chk.Text, com_relax_other.Text, Get_CKL_Click(com_child), Get_CKL_Click(com_Welfar), com_Welfare_chk.Text, com_Welfare_other.Text, com_disputetype.Text, _
                        com_disputecon.Text, com_Mediationtype.Text, com_Mediationcon.Text, cmp_pass.Text, com_schedule.SelectedValue, com_sectype.Text, com_sex.Text, txt_cmp_age.Text, txt_cmp_entrust.Text) Then
            '儲存使用紀錄
            ws.InsUsrRec(PrgMemo, Pgm_Type, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)

            If Pgm_Type <> "MDY" Then
                '取得最新條件
                If txtSubWhere.Text <> "" Then
                    txtSubWhere.Text &= " OR cmp_code='" & RepSql(sKeyPrimary) & "' "
                Else
                    txtSubWhere.Text = " WHERE cmp_code='" & RepSql(sKeyPrimary) & "' "
                End If
            End If
            '存檔正確導回default頁面
			 REM 2024.02.27 增加輸入編碼過濾機制 
            Response.Redirect("Default.aspx?subwhere=" & HttpUtility.HtmlEncode(Replace(txtSubWhere.Text, "%", "$")) & "&order=" &  HttpUtility.HtmlEncode(txtOrder.Text))
        Else
            '存檔出錯就秀出錯誤訊息
            Show_Message("儲存失敗!")
        End If
    End Sub

    Protected Function Chk_data() As Boolean
        'Server端資料檢查 自訂函數
        Chk_data = True

        ''檢查重覆資料
        If txtPgmType.Text = "ADD" Then
            If Chk_RelData(s_prgcode, "", txt_cmp_code.Text) = False Then
                Chk_data = False
                Show_Message("有重覆參數名稱，請確認")
            End If
        End If
        '檢查資料正確性
        If com_schedule.selectedvalue = "" Then
            Show_Message("請挑選受理進度")
        End If

        If ddl_cmp_type.SelectedValue = "A" Then
            If txt_cmp_name.Text = "" Then
                Chk_data = False
                Show_Message("申訴人姓名不得為空白，請確認")
            End If
            If txt_cmp_depart.Text = "" Then
                Chk_data = False
                Show_Message("申訴人部門不得為空白，請確認")
            End If
            If txt_cmp_job.Text = "" Then
                Chk_data = False
                Show_Message("申訴人職稱不得為空白，請確認")
            End If

            If msex.Checked = False And fsex.Checked = False Then
                Chk_data = False
                Show_Message("請挑選性別")
            End If
            If rb_entrustY.Checked = False And rb_entrustN.Checked = False Then
                Chk_data = False
                Show_Message("請挑選是否為代他人申訴，如是必須上傳委託函")
            End If

            If txt_cmp_age.Text = "" Or txt_cmp_age.Text = "請輸入年齡" Then
                Chk_data = False
                Show_Message("年齡不得為空白，請確認")
            Else
                If IsNumeric(txt_cmp_age.Text) = False Then
                    Chk_data = False
                    Show_Message("年齡請輸入數字")
                End If
            End If

            If Chk_Email(txt_cmp_email.Text) = False Then
                Chk_data = False
                Show_Message("電子郵件格式不正確，請確認")
            End If
            If Chk_Tel(txt_cmp_tel1.Text) = False Then
                Chk_data = False
                Show_Message("聯絡電話格式不正確，請確認")
            End If
            If rblist.checked = False And rblist2.checked = False And rblist3.checked = False Then
                Chk_data = False
                Show_Message("請挑選身份")
            ElseIf rblist.checked = True Then
                If Chk_DateForm(cmp_sdate.Text) = False Then
                    Chk_data = False
                    Show_Message("到職日期格式不正確，請確認")
                End If
                If jobcontent.Text = "" Then
                    Chk_data = False
                    Show_Message("請輸入擔任職務及工作內容")
                End If
            ElseIf rblist2.checked = True Then
                If Chk_DateForm(cmp_edate.Text) = False Then
                    Chk_data = False
                    Show_Message("離職日期格式不正確，請確認")
                End If

            End If

            If txt_com_idno.Text <> "" And txt_com_idno.Text <> "未填寫" Then
                If IsNumeric(txt_com_idno.Text) = False Then
                    Chk_data = False
                    Show_Message("統一編號請輸入數字")
                End If
            End If

            If txt_com_add.Text = "" Then
                Chk_data = False
                Show_Message("業單位地址不得為空白，請確認")
            End If

            If com_type.Text = "" Then
                Chk_data = False
                Show_Message("請輸入事業單位行業別")
            End If


            If txt_com_add1.Text = "" Then
                Chk_data = False
                Show_Message("勞工勞務給付地不得為空白，請確認")
            End If
            'If txt_com_idno.Text = "" Or txt_com_idno.Text = "請輸入被申訴事業單位統編" Then
            '    Chk_data = False
            '    sErrMsg &= "被申訴事業單位統編不得為空白，請確認\n"
            'End If
            If com_Contract.SelectedValue = "" And com_Contract_chk.Checked = False And com_paymon.SelectedValue = "" And com_relax.SelectedValue = "" And com_relax_chk.Checked = False And com_child.SelectedValue = "" And com_Welfar.SelectedValue = "" And com_Welfare_chk.Checked = False Then
                Chk_data = False
                Show_Message("請挑選申訴議題")
            End If

            If DN.checked = False And DY.checked = False Then
                Chk_data = False
                Show_Message("請勾選本爭議事項是否有向其他勞工主管機關申訴")
            ElseIf DY.checked = True Then
                If com_disputecon.Text = "" Then
                    Chk_data = False
                    Show_Message("請註明日期及機關")
                End If
            End If

            If MN.checked = False And MYY.checked = False Then
                Chk_data = False
                Show_Message("請勾選本爭議事項是否曾申請勞資爭議調解")
            ElseIf MYY.checked = True Then
                If com_Mediationcon.Text = "" Then
                    Chk_data = False
                    Show_Message("請註明日期及機關")
                End If
            End If

            If txt_cmp_memo.Text = "" Then
                Chk_data = False
                Show_Message("申訴具體內容不得為空白，請確認")
            End If
            'If Chk_ID(txt_cmp_idno.Text) = False Then
            '    Chk_data = False
            '    Show_Message("身份證字號格式不正確，請確認")
            'End If
            'If ddl_is_agree.SelectedValue = "" Then
            '    Chk_data = False
            '    Show_Message("個資限制未選擇，請確認")
            'End If

        End If
        'If Chk_CompanyNo(txt_com_idno.Text) = False Then
        '    Chk_data = False
        '    Show_Message("統編格式不正確，請確認")
        'End If

        '檢查必填欄位是否空白
        If txt_com_name.Text = "" Or txt_com_add.Text = "" Or txt_cmp_memo.Text = "" _
                                  Or ddl_att_area.SelectedValue = "" Or ddl_att_road.SelectedValue = "" Then
            Chk_data = False
            Show_Message("必填欄位不得為空白，請確認")
        End If
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '接收default頁面變數
        If Not IsPostBack Then

            SetClientCheck()
            '設定下拉式選單
            'Set_DDL_Option("ZipCode", "110", ddl_att_area, "", "---請選擇---", "")
            Set_DDL_Option("areacode", "are01", ddl_att_area, "", "---請選擇---", "")
            Set_DDL_Option(ddl_att_area.SelectedValue, "", ddl_att_road, "", "---請選擇---", "")

            'Set_DDL_Option("ZipCode", "110", ddl_att_area1, "", "---請選擇---", "")
            Set_DDL_Option("areacode", "are01", ddl_att_area1, "", "---請選擇---", "")
            Set_DDL_Option(ddl_att_area1.SelectedValue, "", ddl_att_road1, "", "---請選擇---", "")

            Set_DDL_Option("CmpType", "", ddl_cmp_type, "", "", "")

            Set_CKL_Option("com_Contra", "", com_Contract, "", "", "")

            Set_CKL_Option("com_paymon", "", com_paymon, "", "", "")

            Set_CKL_Option("com_relax", "", com_relax, "", "", "")

            Set_CKL_Option("com_child", "", com_child, "", "", "")

            Set_CKL_Option("com_Welfar", "", com_Welfar, "", "", "")

            If txtPgmType.Text <> "MDY" Then
                Set_Ver_Code()
                ins_date.Text = Format(Now, "yyyy/MM/dd")
                cmp_date.Text = DateTime.Now.ToString("HH:mm:ss")
                cmp_pass.Text = Session("ComplainA_VerCode")
            End If
            REM Set_DDL_Option("YorN", "", ddl_is_agree, "", "---請選擇---", "")
            REM 標頭 
            Label1.Text = CType(PreviousPage.FindControl("Label1"), Label).Text
            REM 編輯模式
            txtPgmType.Text = CType(PreviousPage.FindControl("txtPgmType"), TextBox).Text
            ddl_is_agree.SelectedValue = "Y"
            REM 首頁條件
            txtSubWhere.Text = CType(PreviousPage.FindControl("txtsubwhere"), TextBox).Text
            txtOrder.Text = CType(PreviousPage.FindControl("txtOrder"), TextBox).Text
            REM  非新增模式則預設取得資料 2024.0227 加上 HtmlEncode
            Dim Pgm_type As String =HttpUtility.HtmlEncode( txtPgmType.Text)
			
            txtPrimary.Text = HttpUtility.HtmlEncode(CType(PreviousPage.FindControl("txtPrimary"), TextBox).Text)

            Select Case Pgm_type
                Case "MDY", "COPY"
                    Dim Dt_tmp As DataTable = Get_DataSet(s_prgcode, "Detail", txtPrimary.Text).Tables(0)
                    If Dt_tmp.Rows.Count > 0 Then

                        ins_date.Text = Trim(Dt_tmp.Rows(0).Item("ins_date") & "")
                        cmp_date.Text = Trim(Dt_tmp.Rows(0).Item("cmp_date") & "")
                        txt_cmp_code.Text = Trim(Dt_tmp.Rows(0).Item("cmp_code") & "")
                        If Trim(Dt_tmp.Rows(0).Item("cmp_entrust") & "") = "1" Then
                            rb_entrustY.Checked = True
                        End If
                        If Trim(Dt_tmp.Rows(0).Item("cmp_entrust") & "") = "0" Then
                            rb_entrustN.Checked = True
                        End If
                        txt_cmp_name.Text = Trim(Dt_tmp.Rows(0).Item("cmp_name") & "")
                        txt_cmp_tel1.Text = Trim(Dt_tmp.Rows(0).Item("cmp_tel1") & "")
                        txt_cmp_depart.Text = Trim(Dt_tmp.Rows(0).Item("cmp_depart") & "")
                        txt_cmp_job.Text = Trim(Dt_tmp.Rows(0).Item("cmp_job") & "")
                        txt_cmp_email.Text = Trim(Dt_tmp.Rows(0).Item("cmp_email") & "")
                        If Trim(Dt_tmp.Rows(0).Item("cmp_idclass") & "") = "1" Then
                            rblist.Checked = True
                            cmp_sdate.Text = Trim(Dt_tmp.Rows(0).Item("cmp_sdate") & "")
                            jobcontent.Text = Trim(Dt_tmp.Rows(0).Item("cmp_jobcontent") & "")
                        End If
                        If Trim(Dt_tmp.Rows(0).Item("cmp_idclass") & "") = "2" Then
                            rblist2.Checked = True
                            cmp_edate.Text = Trim(Dt_tmp.Rows(0).Item("cmp_edate") & "")
                        End If
                        If Trim(Dt_tmp.Rows(0).Item("cmp_idclass") & "") = "3" Then
                            rblist3.Checked = True
                        End If

                        If Trim(Dt_tmp.Rows(0).Item("com_Contract_chk") & "") = "Y" Then
                            com_Contract_chk.Checked = True
                            com_Contract_other.Text = Trim(Dt_tmp.Rows(0).Item("com_Contract_other") & "")
                        End If
                        com_type.Text = Trim(Dt_tmp.Rows(0).Item("com_type") & "")
                        'txt_cmp_idno.Text = Trim(Dt_tmp.Rows(0).Item("cmp_idno") & "")
                        com_schedule.SelectedValue = Trim(Dt_tmp.Rows(0).Item("com_schedule") & "")
                        txt_com_name.Text = Trim(Dt_tmp.Rows(0).Item("com_name") & "")
                        txt_com_add.Text = Trim(Dt_tmp.Rows(0).Item("com_add") & "")
                        txt_com_add1.Text = Trim(Dt_tmp.Rows(0).Item("com_payadd") & "")
                        txt_com_idno.Text = Trim(Dt_tmp.Rows(0).Item("com_idno") & "")
                        ddl_cmp_type.SelectedValue = Trim(Dt_tmp.Rows(0).Item("cmp_type") & "")
                        ddl_att_area.SelectedValue = Trim(Dt_tmp.Rows(0).Item("att_area") & "")
                        ddl_att_area_SelectedIndexChanged(sender, e)
                        ddl_att_road.SelectedValue = Trim(Dt_tmp.Rows(0).Item("att_road") & "")
                        ddl_att_area1.SelectedValue = Trim(Dt_tmp.Rows(0).Item("com_payarea") & "")
                        ddl_att_area1_SelectedIndexChanged(sender, e)
                        ddl_att_road1.SelectedValue = Trim(Dt_tmp.Rows(0).Item("com_payroad") & "")
                        Set_CKL_Click(Trim(Dt_tmp.Rows(0).Item("com_Contract") & ""), com_Contract, "com_Contra")
                        Set_CKL_Click(Trim(Dt_tmp.Rows(0).Item("com_paymon") & ""), com_paymon, "com_paymon")
                        Set_CKL_Click(Trim(Dt_tmp.Rows(0).Item("com_relax") & ""), com_relax, "com_relax")
                        If Trim(Dt_tmp.Rows(0).Item("com_relax_chk") & "") = "Y" Then
                            com_relax_chk.Checked = True
                            com_relax_other.Text = Trim(Dt_tmp.Rows(0).Item("com_relax_other") & "")
                        End If
                        Set_CKL_Click(Trim(Dt_tmp.Rows(0).Item("com_child") & ""), com_child, "com_child")
                        Set_CKL_Click(Trim(Dt_tmp.Rows(0).Item("com_Welfar") & ""), com_Welfar, "com_Welfar")
                        If Trim(Dt_tmp.Rows(0).Item("com_Welfare_chk") & "") = "Y" Then
                            com_Welfare_chk.Checked = True
                            com_Welfare_other.Text = Trim(Dt_tmp.Rows(0).Item("com_Welfare_other") & "")
                        End If

                        If Trim(Dt_tmp.Rows(0).Item("com_disputetype") & "") = "Y" Then
                            DY.Checked = True
                            com_disputecon.Text = Trim(Dt_tmp.Rows(0).Item("com_disputecon") & "")
                        Else
                            DN.Checked = True
                        End If


                        If Trim(Dt_tmp.Rows(0).Item("com_Mediationtype") & "") = "Y" Then
                            MYY.Checked = True
                            com_Mediationcon.Text = Trim(Dt_tmp.Rows(0).Item("com_Mediationcon") & "")
                        Else
                            MN.Checked = True
                        End If

                        If Trim(Dt_tmp.Rows(0).Item("com_sectype") & "") = "Y" Then
                            SY.Checked = True
                        Else
                            SN.Checked = True
                        End If

                        If Trim(Dt_tmp.Rows(0).Item("com_sex") & "") = "m" Then
                            msex.Checked = True
                        Else
                            fsex.Checked = True
                        End If

                        txt_cmp_age.Text = Trim(Dt_tmp.Rows(0).Item("cmp_age") & "")

                        cmp_pass.Text = Trim(Dt_tmp.Rows(0).Item("cmp_pass") & "")
                        ddl_is_agree.SelectedValue = Trim(Dt_tmp.Rows(0).Item("is_agree") & "")
                        txt_cmp_memo.Text = Trim(Dt_tmp.Rows(0).Item("cmp_memo") & "")
                        lab_cmp_file1.Text = Set_Files("File_LSI002A", Trim(Dt_tmp.Rows(0).Item("cmp_file1") & ""))
                        lab_cmp_file2.Text = Set_Files("File_LSI002A", Trim(Dt_tmp.Rows(0).Item("cmp_file2") & ""))
                        lab_cmp_file3.Text = Set_Files("File_LSI002A", Trim(Dt_tmp.Rows(0).Item("cmp_file3") & ""))
                        lab_cmp_file4.Text = Set_Files("File_LSI002A", Trim(Dt_tmp.Rows(0).Item("cmp_file4") & ""))
                        lab_cmp_file5.Text = Set_Files("File_LSI002A", Trim(Dt_tmp.Rows(0).Item("cmp_file5") & ""))
                    End If

                    If Pgm_type = "COPY" Then
                        txt_cmp_code.Text = ""
                        txt_cmp_code.Enabled = False
                    Else
                        txt_cmp_code.Enabled = False
                        '具名申訴則顯示補發按鈕
                        If ddl_cmp_type.SelectedValue = "A" And Chk_Email(txt_cmp_email.Text) Then
                            txtFileName1.Text = Trim(Dt_tmp.Rows(0).Item("cmp_file1") & "")
                            txtFileName2.Text = Trim(Dt_tmp.Rows(0).Item("cmp_file2") & "")
                            txtFileName3.Text = Trim(Dt_tmp.Rows(0).Item("cmp_file3") & "")
                            txtFileName4.Text = Trim(Dt_tmp.Rows(0).Item("cmp_file4") & "")
                            txtFileName5.Text = Trim(Dt_tmp.Rows(0).Item("cmp_file5") & "")
                            Bt_send.Visible = True
                        End If
                    End If
            End Select
        End If
        img_date2.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "cmp_sdate")
        img_date3.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "cmp_edate")
        l_ErrMsg.Text = ""
    End Sub

    Protected Sub SetClientCheck()
        '網頁伺服器端檢查點建置
        'txt_par_value.Attributes("onblur") = "javascript:IsEmpty(this);"
    End Sub
    Protected Sub Bt_BackUp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Bt_BackUp.Click
        
		Response.Redirect("Default.aspx?subwhere=" & HttpUtility.HtmlEncode(Replace(txtSubWhere.Text, "%", "$")) & "&order=" &  HttpUtility.HtmlEncode(txtOrder.Text))
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

    Protected Sub ddl_att_area_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddl_att_area.SelectedIndexChanged
        ddl_att_road.Items.Clear()
        Set_DDL_Option(ddl_att_area.SelectedValue, "", ddl_att_road, "", "---請選擇---", "")
    End Sub
    Protected Sub ddl_att_area1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddl_att_area1.SelectedIndexChanged
        ddl_att_road1.Items.Clear()
        Set_DDL_Option(ddl_att_area1.SelectedValue, "", ddl_att_road1, "", "---請選擇---", "")
    End Sub

    Protected Sub Bt_send_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Bt_send.Click
        '發送Mail通知
        Dim dtMail As New DataTable
        Dim column1 As New DataColumn
        Dim column2 As New DataColumn
        column1.DataType = System.Type.GetType("System.String")
        column2.DataType = System.Type.GetType("System.String")
        dtMail.Columns.Add(column1)
        dtMail.Columns.Add(column2)
        '加入申訴人Mail
        Dim rowMail As DataRow = dtMail.NewRow
        rowMail(0) = txt_cmp_email.Text
        rowMail(1) = txt_cmp_name.Text
        dtMail.Rows.Add(rowMail)
        '加入承辦人信箱
        Dim MailList As String() = Split(Get_SysPara("sug_receive"), "|")
        For i As Integer = 0 To UBound(MailList)
            If String.IsNullOrEmpty(MailList(i)) = False Then
                Dim rowTmp As DataRow = dtMail.NewRow
                rowTmp(0) = MailList(i)
                rowTmp(1) = "申訴承辦人"
                dtMail.Rows.Add(rowTmp)
            End If
        Next
        Dim sIs_agree As String = "不同意本處繼續使用資料"
        'If ddl_is_agree.SelectedValue = "Y" Then sIs_agree = "同意本處繼續使用資料"
        '郵件內容

        Dim showtype1 As String = ""
        If rblist.checked = True Then
            showtype1 = "事業單位所屬勞工(目前尚在職者)</br>" & "到職日期：" & cmp_sdate.Text & "</br>擔任職務及工作內容：" & jobcontent.Text
        End If
        If rblist2.checked = True Then
            showtype1 = "之前曾經是事業單位勞工，離職日期：" & cmp_edate.Text
        End If
        If rblist3.checked = True Then
            showtype1 = "民眾(非屬前二項者)"
        End If

        Dim bFirst As Boolean = True
        bFirst = True

        Dim ContractList As String = ""
        For i As Integer = 0 To com_Contract.Items.Count - 1
            If com_Contract.Items(i).Selected = True Then
                If bFirst = False Then ContractList &= "、"
                bFirst = False
                ContractList &= com_Contract.Items(i).Text
            End If
        Next
        Dim show1 As String = ""
        If com_Contract_chk.checked = True Then
            If ContractList <> "" Then
                show1 = "、其他：" & com_Contract_other.text
            Else
                show1 = "其他：" & com_Contract_other.text
            End If

        End If

        bFirst = True
        Dim paymoneyList As String = ""
        For i As Integer = 0 To com_paymon.Items.Count - 1
            If com_paymon.Items(i).Selected = True Then
                If bFirst = False Then paymoneyList &= "、"
                bFirst = False
                paymoneyList &= com_paymon.Items(i).Text
            End If
        Next

        bFirst = True
        Dim relaxList As String = ""
        For i As Integer = 0 To com_relax.Items.Count - 1
            If com_relax.Items(i).Selected = True Then
                If bFirst = False Then relaxList &= "、"
                bFirst = False
                relaxList &= com_relax.Items(i).Text
            End If
        Next
        Dim show2 As String = ""
        If com_relax_chk.checked = True Then
            If relaxList <> "" Then
                show2 = "、其他：" & com_relax_other.text
            Else
                show2 = "其他：" & com_relax_other.text
            End If

        End If

        bFirst = True
        Dim childList As String = ""
        For i As Integer = 0 To com_child.Items.Count - 1
            If com_child.Items(i).Selected = True Then
                If bFirst = False Then childList &= "、"
                bFirst = False
                childList &= com_child.Items(i).Text
            End If
        Next

        bFirst = True
        Dim WelfareList As String = ""
        For i As Integer = 0 To com_Welfar.Items.Count - 1
            If com_Welfar.Items(i).Selected = True Then
                If bFirst = False Then WelfareList &= "、"
                bFirst = False
                WelfareList &= com_Welfar.Items(i).Text
            End If
        Next
        Dim show3 As String = ""
        If com_Welfare_chk.checked = True Then
            If WelfareList <> "" Then
                show3 = "、其他：" & com_Welfare_other.text
            Else
                show3 = "其他：" & com_Welfare_other.text
            End If

        End If
        Dim dtype1 As String = ""
        If DN.checked = True Then
            dtype1 = "否"
        End If
        If DY.checked = True Then
            dtype1 = "是，請註明日期及機關" & com_disputecon.Text
        End If

        Dim mtype1 As String = ""
        If MN.checked = True Then
            mtype1 = "否"
        End If
        If MYY.checked = True Then
            mtype1 = "是，請註明日期及機關" & com_Mediationcon.Text
        End If

        Dim ssex As String = ""
        If msex.Checked = True Then
            com_sex.Text = "m"
            ssex = "男"
        End If

        If fsex.Checked = True Then
            com_sex.Text = "f"
            ssex = "女"
        End If

        Dim stype1 As String = ""
        If SN.checked = True Then
            stype1 = "目前仍在職，請將我的身分保密"
        End If
        If SY.checked = True Then
            stype1 = "我願意將姓名提供給事業單位，以利檢查時協助追償受損權益"
        End If

        Dim sBody As String = "" & _
            "<html xmlns='http://www.w3.org/1999/xhtml'>" & _
            "<head><meta http-equiv='Content-Type' content='text/html; charset=utf-8' /></head><body>" & _
            "<table cellpadding=5 cellspacing=0 bordercolor=#7EBAD1 border=1 bgcolor=#ffffff style='border-collapse:collapse;'" & _
            "        width='720' summary='排版表格' align=center class='c12'>" & _
            "<tr bgcolor='#7EBAD1' align='center'><td colspan='4'><p><strong>具名申訴</strong></p></td></tr>" & _
                "<tr>" & _
                "  <td width='25%'>受理日期：</td>" & _
                "  <td colspan='3'>" & ins_date.Text & "</td>" & _
                "</tr>" & _
                "<tr>" & _
                "  <td>受理時間：</td>" & _
                "  <td colspan='3'>" & cmp_date.Text & "</td>" & _
                "</tr>" & _
                "<tr>" & _
                "  <td>案件編號：</td>" & _
                "  <td colspan='3'>" & txt_cmp_code.Text & "</td>" & _
                "</tr>" & _
                "<tr bgcolor='#7EBAD1' align='left'><td colspan='4'><p><strong>當事人基本資料</strong></p></td></tr>" & _
                "<tr>" & _
                "  <td>姓　　　　　名：</td>" & _
                "  <td colspan='3'>" & txt_cmp_name.Text & "</td>" & _
                "</tr>" & _
                "</tr>" & _
                "<tr>" & _
                "  <td>性　　　　　別：</td>" & _
                "  <td colspan='3'>" & ssex & "</td>" & _
                "</tr>" & _
                "<tr>" & _
                "  <td>年　　　　　齡：</td>" & _
                "  <td colspan='3'>" & txt_cmp_age.Text & "</td>" & _
                "</tr>" & _
                "<tr>" & _
                "<td>電子信箱E-mail：</td>" & _
                "<td colspan='3'>" & txt_cmp_email.Text & "</td>" & _
                "</tr>" & _
                "<tr>" & _
                "  <td>聯　絡　電　話：</td>" & _
                "  <td colspan='3'>" & txt_cmp_tel1.Text & "</td>" & _
                "</tr>" & _
                "<tr>" & _
                "  <td>身份：</td>" & _
                "  <td colspan='3'>" & showtype1 & "</td>" & _
                "</tr>" & _
                "<tr bgcolor='#7EBAD1' align='left'><td colspan='4'><p><strong>事業單位基本資料</strong></p></td></tr>" & _
                "<tr>" & _
                "<td>全名：</td>" & _
                "<td colspan='3'>" & txt_com_name.Text & "</td>" & _
                "</tr>" & _
                "<tr>" & _
                "<td>統一編號：</td>" & _
                "<td colspan='3'>" & txt_com_idno.Text & "</td>" & _
                "</tr>" & _
                "<tr>" & _
                "<td>地址：</td>" & _
                "<td colspan='3'>" & ddl_att_area.SelectedItem.Text & ddl_att_road.SelectedItem.Text & txt_com_add.Text & "</td>" & _
                "</tr>" & _
                "<tr>" & _
                "<td>事業單位行業別：</td>" & _
                "<td colspan='3'>" & com_type.Text & "</td>" & _
                "</tr>" & _
                "<tr>" & _
                "<td>勞工勞務給付地：</td>" & _
                "<td colspan='3'>" & ddl_att_area1.SelectedItem.Text & ddl_att_road1.SelectedItem.Text & txt_com_add1.Text & "</td>" & _
                "</tr>" & _
                "<tr align='left'><td colspan='4'><p><strong>申訴議題</strong></p></td></tr>" & _
                "<tr bgcolor='#7EBAD1' align='left'><td colspan='4'><p><strong>勞動契約</strong></p></td></tr>" & _
                "<tr>" & _
                "<td colspan='4'>" & ContractList & show1 & "&nbsp;</td>" & _
                "</tr>" & _
                "<tr bgcolor='#7EBAD1' align='left'><td colspan='4'><p><strong>工資</strong></p></td></tr>" & _
                "<tr>" & _
                "<td colspan='4'>" & paymoneyList & "&nbsp;</td>" & _
                "</tr>" & _
                "<tr bgcolor='#7EBAD1' align='left'><td colspan='4'><p><strong>工作時間、休息、休假</strong></p></td></tr>" & _
                "<tr>" & _
                "<td colspan='4'>" & relaxList & show2 & "&nbsp;</td>" & _
                "</tr>" & _
                "<tr bgcolor='#7EBAD1' align='left'><td colspan='4'><p><strong>童工、女工</strong></p></td></tr>" & _
                "<tr>" & _
                "<td colspan='4'>" & childList & "&nbsp;</td>" & _
                "</tr>" & _
                "<tr bgcolor='#7EBAD1' align='left'><td colspan='4'><p><strong>保障與福利</strong></p></td></tr>" & _
                "<tr>" & _
                "<td colspan='4'>" & WelfareList & show3 & "&nbsp;</td>" & _
                "</tr>" & _
                 "<tr>" & _
                "  <td>本爭議事項是否有向其他勞工主管機關申訴：</td>" & _
                "  <td colspan='3'>" & dtype1 & "</td>" & _
                "</tr>" & _
                "<tr>" & _
                "  <td>本爭議事項是否曾申請勞資爭議調解：</td>" & _
                "  <td colspan='3'>" & mtype1 & "</td>" & _
                "</tr>" & _
                "<tr>" & _
                "<td colspan='4'><p>申訴具體內容：<br/>" & _
                    Replace(RepSql(txt_cmp_memo.Text), vbCrLf, "<br/>") & _
                "    </p></td>" & _
                "</tr>" & _
                "<tr>" & _
                "  <td>為利於協助追償受損權益，是否願意將您的姓名提供給事業單位：</td>" & _
                "  <td colspan='3'>" & stype1 & "</td>" & _
                "</tr>" & _
                "<tr>" & _
                "<td>安全密碼：</td>" & _
                "<td colspan='3'>" & cmp_pass.Text & "</td>" & _
                "</tr>" & _
                "<tr>" & _
                "<td colspan='4' align='center'><font color='#666666' style=' size:-1'>謝謝您的來信！本局將以最快速度回覆！！</font></td>" & _
                "</tr>" & _
                "</table></body>"

        Dim sUpFolder As String = Get_SysPara("AP_File")
        Dim sFilePath1 As String = sUpFolder & txtFileName1.Text
        Dim sFilePath2 As String = sUpFolder & txtFileName2.Text
        Dim sFilePath3 As String = sUpFolder & txtFileName3.Text
        Dim sFilePath4 As String = sUpFolder & txtFileName4.Text
        Dim sFilePath5 As String = sUpFolder & txtFileName5.Text
        If SendMail("臺北市政府勞動局勞動條件申訴信箱-具名申訴-送出成功", dtMail, sBody, s_prgcode, sFilePath1, sFilePath2, sFilePath3, sFilePath4, sFilePath5) Then
            Show_Message("送出成功")
        Else
            Show_Message("送出失敗，請通知系統管理員")
            Exit Sub
        End If
    End Sub

    Protected Function Set_Ver_Code() As Boolean
        '產生亂數英數字串
        Dim array As String = "#123456789ABCDEFGHIJKLMN$PQRSTUVWXYZ"
        Dim generator As New Random
        txtVerCode.Text = ""
        For i = 1 To 6
            txtVerCode.Text &= array(Int(generator.Next(0, 36)))
            Session("ComplainA_VerCode") = txtVerCode.Text
        Next
    End Function
End Class
