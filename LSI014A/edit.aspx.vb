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
                        , "", "", "", "" _
                        , "", "", "", "" _
                        , txt_work_type_memo.Text, Get_CKL_Click(ckl_use_tool1), "", "", "" _
                        , "", "", txt_use_tool5_memo.Text, Get_CKL_Click(ckl_save_tool), txt_save_tool_memo.Text _
                        , "", txt_dan_memo_memo.Text, txt_dan_mail.Text, txt_dan_mail2.Text, Session("usr_code"), Format(Now, "yyyy/MM/dd"), _
                        txt_fix_code.Text, ddl_is_holiday.SelectedValue) Then


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

        '檢查重覆資料
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
            Show_Message("MAIL格式不正確，請確認")
        End If
        If Chk_Email(txt_dan_mail2.Text) = False And txt_dan_mail2.Text <> "" Then
            If txt_dan_mail2.Text <> "未填寫" Then
                Chk_data = False
                Show_Message("MAIL2格式不正確，請確認")
            End If
        End If
        If IsDate(txt_pet_date.Text) = False Or IsDate(txt_work_date_s.Text)=False Or IsDate(txt_work_date_e.Text)=False Then
            Chk_data = False
            Show_Message("日期不正確，請確認")
        Else
            If Date.Compare(txt_work_date_s.Text, txt_work_date_e.Text) > 0 Then
                Chk_data = False
                Show_Message("施工期間日期順序不正確，請確認")
            End If
        End If
        

        '檢查必填欄位是否空白
        If txt_com_name.Text = "" Or ddl_att_area.SelectedValue = "" Or ddl_att_road.SelectedValue = "" Or txt_att_add1.Text = "" _
                                  Or txt_pet_tel1.Text = "" Or txt_pet_date.Text = "" Or txt_pet_time.Text = "" Or txt_pet_name.Text = "" _
                                  Or txt_work_date_s.Text = "" Or txt_work_date_e.Text = "" Or txt_att_name1.Text = "" _
                                  Or txt_att_tel1.Text = "" Or txt_att_name2.Text = "" Or txt_att_tel2.Text = "" _
                                  Or txt_work_floor_s.Text = "" Or txt_work_floor_e.Text = "" Or ddl_work_type1.SelectedValue = "" _
                                  Or txt_dan_mail.Text = "" Then
            Chk_data = False
            Show_Message("必填欄位不得為空白，請確認")
        End If
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '接收default頁面變數
        If Not IsPostBack Then
            SetClientCheck()
            Set_DDL_Option("ZipCode", "110", ddl_att_area, "", "---請選擇---", "")
            Set_DDL_Option(ddl_att_area.SelectedValue, "", ddl_att_road, "", "---請選擇---", "")
            Set_CKL_Option("save_tool", "", ckl_save_tool, "", "", "")
            'Set_CKL_Option("dan_memo", "", ckl_dan_memo, "", "", "")
            Set_DDL_Option("YorN", "N", ddl_is_night, "", "---請選擇---", "")
            Set_DDL_Option("YorN", "N", ddl_is_holiday, "", "---請選擇---", "")
            Set_DDL_Option("work_type1", "", ddl_work_type1, "", "---請選擇---", "")
            'Set_DDL_Option("work_type2", "", ddl_work_type2, "", "---請選擇---", "")
            'Set_DDL_Option("work_type3", "", ddl_work_type3, "", "---請選擇---", "")
            'Set_DDL_Option("work_type4", "", ddl_work_type4, "", "---請選擇---", "")
            'Set_DDL_Option("work_type5", "", ddl_work_type5, "", "---請選擇---", "")
            'Set_DDL_Option("work_type6", "", ddl_work_type6, "", "---請選擇---", "")
            'Set_DDL_Option("work_type7", "", ddl_work_type7, "", "---請選擇---", "")
            'Set_DDL_Option("work_type8", "", ddl_work_type8, "", "---請選擇---", "")
            'Set_DDL_Option("work_type9", "", ddl_work_type9, "", "---請選擇---", "")
            Set_CKL_Option("use_tool1", "", ckl_use_tool1, "", "", "")
            'Set_DDL_Option("use_tool1", "", ddl_use_tool1, "", "---請選擇---", "")
            'Set_DDL_Option("use_tool2", "", ddl_use_tool2, "", "---請選擇---", "")
            'Set_DDL_Option("use_tool3", "", ddl_use_tool3, "", "---請選擇---", "")
            'Set_DDL_Option("use_tool4", "", ddl_use_tool4, "", "---請選擇---", "")
            'Set_DDL_Option("use_tool5", "", ddl_use_tool5, "", "---請選擇---", "")

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

            If Pgm_type = "MDY" Then l_PgmType.Text = "[修改模式]" Else l_PgmType.Text = "[新增模式]"

            Select Case Pgm_type
                Case "ADD"
                    txt_pet_date.Text = Format(Now, "yyyy/MM/dd")
                    txt_pet_time.Text = Format(Now, "HH:mm")
                    txt_work_date_s.Text = Format(Now, "yyyy/MM/dd")
                    txt_work_date_e.Text = Format(Now, "yyyy/MM/dd")
                    txt_work_floor_s.Text = "1F"
                    txt_work_floor_e.Text = "1F"
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
                        ddl_att_road.Items.Clear()
                        Set_DDL_Option(ddl_att_area.SelectedValue, "", ddl_att_road, "", "---請選擇---", "")
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
                        ddl_is_holiday.SelectedValue = Trim(Dt_tmp.Rows(0).Item("is_holiday") & "")
                        Set_CKL_Click(Trim(Dt_tmp.Rows(0).Item("save_tool") & ""), ckl_save_tool, "save_tool")
                        ddl_work_type1.SelectedValue = Trim(Dt_tmp.Rows(0).Item("work_type1") & "")
                        'ddl_work_type2.SelectedValue = Trim(Dt_tmp.Rows(0).Item("work_type2") & "")
                        'ddl_work_type3.SelectedValue = Trim(Dt_tmp.Rows(0).Item("work_type3") & "")
                        'ddl_work_type4.SelectedValue = Trim(Dt_tmp.Rows(0).Item("work_type4") & "")
                        'ddl_work_type5.SelectedValue = Trim(Dt_tmp.Rows(0).Item("work_type5") & "")
                        'ddl_work_type6.SelectedValue = Trim(Dt_tmp.Rows(0).Item("work_type6") & "")
                        'ddl_work_type7.SelectedValue = Trim(Dt_tmp.Rows(0).Item("work_type7") & "")
                        'ddl_work_type8.SelectedValue = Trim(Dt_tmp.Rows(0).Item("work_type8") & "")
                        'ddl_work_type9.SelectedValue = Trim(Dt_tmp.Rows(0).Item("work_type9") & "")
                        txt_work_type_memo.Text = Trim(Dt_tmp.Rows(0).Item("work_type_memo") & "")
                        Set_CKL_Click(Trim(Dt_tmp.Rows(0).Item("use_tool1") & ""), ckl_use_tool1, "use_tool1")
                        'ddl_use_tool1.SelectedValue = Trim(Dt_tmp.Rows(0).Item("use_tool1") & "")
                        'ddl_use_tool2.SelectedValue = Trim(Dt_tmp.Rows(0).Item("use_tool2") & "")
                        'ddl_use_tool3.SelectedValue = Trim(Dt_tmp.Rows(0).Item("use_tool3") & "")
                        'ddl_use_tool4.SelectedValue = Trim(Dt_tmp.Rows(0).Item("use_tool4") & "")
                        'ddl_use_tool5.SelectedValue = Trim(Dt_tmp.Rows(0).Item("use_tool5") & "")
                        'txt_use_tool3_memo.Text = Trim(Dt_tmp.Rows(0).Item("use_tool3_memo") & "")
                        txt_use_tool5_memo.Text = Trim(Dt_tmp.Rows(0).Item("use_tool5_memo") & "")
                        'Set_CKL_Click(Trim(Dt_tmp.Rows(0).Item("dan_memo") & ""), ckl_dan_memo, "dan_memo")
                        txt_save_tool_memo.Text = Trim(Dt_tmp.Rows(0).Item("save_tool_memo") & "")
                        txt_dan_memo_memo.Text = Trim(Dt_tmp.Rows(0).Item("dan_memo_memo") & "")
                        txt_dan_mail.Text = Trim(Dt_tmp.Rows(0).Item("dan_mail") & "")
                        txt_dan_mail2.Text = Trim(Dt_tmp.Rows(0).Item("dan_mail2") & "")
                        txt_fix_code.Text = Trim(Dt_tmp.Rows(0).Item("fix_code") & "")
                    End If

                    If Pgm_type = "COPY" Then
                        txt_dan_code.Text = ""
                        txt_dan_code.Enabled = False
                    Else
                        txt_fix_code.Enabled = False
                        txt_dan_code.Enabled = False
                        If Chk_Email(txt_dan_mail.Text) Then
                            Bt_send.Visible = True
                        End If
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

    Protected Sub ddl_att_area_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddl_att_area.SelectedIndexChanged
        ddl_att_road.Items.Clear()
        Set_DDL_Option(ddl_att_area.SelectedValue, "", ddl_att_road, "", "---請選擇---", "")
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
        '加入通報人Mail
        Dim rowMail As DataRow = dtMail.NewRow
        rowMail(0) = txt_dan_mail.Text
        rowMail(1) = txt_pet_name.Text
        dtMail.Rows.Add(rowMail)
        If txt_dan_mail2.Text <> "" Then
            Dim rowMail2 As DataRow = dtMail.NewRow
            rowMail2(0) = txt_dan_mail2.Text
            rowMail2(1) = ""
            dtMail.Rows.Add(rowMail2)
        End If
        '加入承辦人信箱
        Dim MailList As String() = Split(Get_SysPara("DA_receive"), "|")
        For i As Integer = 0 To UBound(MailList)
            If String.IsNullOrEmpty(MailList(i)) = False Then
                Dim rowTmp As DataRow = dtMail.NewRow
                rowTmp(0) = MailList(i)
                rowTmp(1) = "危險性作業通報承辦人"
                dtMail.Rows.Add(rowTmp)
            End If
        Next
        '郵件內容
        '作業樓層附註
        Dim sBuild As String = ""
        If txt_bulid_name.Text <> "" Then
            sBuild = txt_bulid_name.Text
        ElseIf txt_bulid_code.Text <> "" Then
            sBuild = "建號及工程名稱：" & txt_bulid_code.Text
        End If
        '作業類別內容
        Dim bFirst As Boolean = True
        Dim sWorkTypeList As String = ""
        If ddl_work_type1.SelectedValue <> "Z" Then
            'bFirst = False
            sWorkTypeList = ddl_work_type1.SelectedItem.Text
        Else
            sWorkTypeList = "其他：" & txt_work_type_memo.Text
        End If
        '使用機具內容
        bFirst = True
        Dim sUseToolList As String = ""
        For i As Integer = 0 To ckl_use_tool1.Items.Count - 1
            If ckl_use_tool1.Items(i).Selected = True Then
                If bFirst = False Then sUseToolList &= "、"
                bFirst = False
                sUseToolList &= ckl_use_tool1.Items(i).Text
            End If
        Next
        If txt_use_tool5_memo.Text <> "" Then
            If bFirst = False Then sUseToolList &= "、"
            sUseToolList &= "其他：" & txt_use_tool5_memo.Text
        End If
        '安全裝置內容
        bFirst = True
        Dim sSaveToolList As String = ""
        For i As Integer = 0 To ckl_save_tool.Items.Count - 1
            If ckl_save_tool.Items(i).Selected = True Then
                If bFirst = False Then sSaveToolList &= "、"
                bFirst = False
                sSaveToolList &= ckl_save_tool.Items(i).Text
            End If
        Next
        If txt_save_tool_memo.Text <> "" Then
            If bFirst = False Then sSaveToolList &= "、"
            sSaveToolList &= "其他：" & txt_save_tool_memo.Text
        End If
        Dim sBody As String = "" & _
            "<table cellpadding=5 cellspacing=0 bordercolor=#7EBAD1 border=1 bgcolor=#ffffff style='border-collapse:collapse;'" & _
            "width='720' summary='排版表格' align=center class='c12'>" & _
            "<tr bgcolor='#7EBAD1' align='center'><td colspan='4'><p><strong>危險性作業通報表</strong></p></td></tr>" & _
            "<tr>" & _
            "  <td>通報單位名稱：</td>" & _
            "  <td colspan='3'>" & txt_com_name.Text & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td width='20%'>通報人員姓名：</td>" & _
            "  <td width='30%'>" & txt_pet_name.Text & "</td>" & _
            "  <td width='20%'>通報人員電話：</td>" & _
            "  <td width='30%'>" & txt_pet_tel1.Text & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>通　報　日　期：</td>" & _
            "  <td>西元 " & Format(Now, "yyyy") & " 年 " & Format(Now, "MM") & " 月 " & Format(Now, "dd") & " 日 </td>" & _
            "  <td>通　報　時　間：</td><td> " & Format(Now, "HH 時 mm 分") & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td rowspan='2'>施工廠商資料：</td>" & _
            "  <td colspan='3'>廠商名稱：" & txt_att_name1.Text & "</td></tr><tr>" & _
            "  <td colspan='3'>公司電話：" & txt_att_tel1.Text & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>施工人員姓名：</td>" & _
            "  <td>" & txt_att_name2.Text & "</td>" & _
            "  <td>施工人員手機：</td>" & _
            "  <td>" & txt_att_tel2.Text & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td rowspan='2'>作　業　地　址：</td>" & _
            "  <td colspan='3'>臺北市" & ddl_att_area.SelectedItem.Text & ddl_att_road.SelectedItem.Text & txt_att_add1.Text & "</td>" & _
            "  </tr><tr><td colspan='3'>" & sBuild & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>預計施工期間：</td>" & _
            "  <td colspan='3'>開工：西元 " & Mid(txt_work_date_s.Text, 1, 4) & " 年 " & _
                                        Mid(txt_work_date_s.Text, 6, 2) & " 月 " & _
                                        Mid(txt_work_date_s.Text, 9, 2) & " 日 " & _
            "                  <br />完工：西元 " & Mid(txt_work_floor_e.Text, 1, 4) & " 年 " & _
                                        Mid(txt_work_floor_e.Text, 6, 2) & " 月 " & _
                                        Mid(txt_work_floor_e.Text, 9, 2) & " 日 (是否夜間施工：" & ddl_is_night.SelectedItem.Text & ")</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>作　業　樓　層：</td>" & _
            "  <td colspan='3'>" & txt_work_floor_s.Text & "～" & txt_work_floor_e.Text & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>作　業　類　別：</td>" & _
            "  <td colspan='3'>" & sWorkTypeList & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>使　用　機　具：</td>" & _
            "  <td colspan='3'>" & sUseToolList & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>安　全　裝　置：</td>" & _
            "  <td colspan='3'>" & sSaveToolList & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>備　　　　　註：</td>" & _
            "  <td colspan='3'>" & txt_dan_memo_memo.Text & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>通　報　回　覆1：</td>" & _
            "  <td colspan='3'>" & txt_dan_mail.Text & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>通　報　回　覆2：</td>" & _
            "  <td colspan='3'>" & txt_dan_mail2.Text & "</td>" & _
            "</tr>" & _
            "</table>"

        If SendMail("臺北市勞動檢查處-危險性作業通報-通報成功", dtMail, sBody, s_prgcode) Then
            Show_Message("送出成功")
        Else
            Show_Message("送出失敗，請通知系統管理員")
            Exit Sub
        End If
    End Sub
End Class
