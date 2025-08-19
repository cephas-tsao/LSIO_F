Imports System.Data

Partial Class basic_LSI013A_edit
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI013A"
	

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '自動編號
        If txtPgmType.Text <> "MDY" Then txt_fix_code.Text = "FI" & Format(Now, "yyMMdd") & Right("0000" & Get_AutoIntMax("Fix_work", "RIGHT(fix_code, 10)", " WHERE fix_code LIKE 'FI" & Format(Now, "yyMMdd") & "%'") + 1, 4)

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

        If chkEng.checked = True Then
            chkEng.text = "Y"
        Else
            chkEng.text = "N"
        End If

        If chktoher.checked = True Then
            chktoher.text = "Y"
        Else
            chktoher.text = "N"
        End If

        'If SaveEditData(s_prgcode, Pgm_Type, txt_fix_code.Text, txt_com_name.Text, txt_pet_tel1.Text, txt_pet_date.Text _
        '                , txt_pet_time.Text, txt_att_add1.Text, ddl_att_area.SelectedValue, ddl_att_road.SelectedValue _
        '                , txt_work_floor.Text, Get_CKL_Click(ckl_work_floor_type), txt_work_date_s.Text, txt_work_date_e.Text _
        '                , txt_att_cname1.Text, txt_att_name1.Text, txt_att_tel1.Text, txt_pet_name1.Text, txt_pet_name2.Text _
        '                , txt_pet_tel3.Text, Get_CKL_Click(ckl_fix_type), Session("usr_code"), Format(Now, "yyyy/MM/dd"), txt_safe_item.Text, txt_dan_code.Text) Then
        '201409 Allen調整
        If SaveEditData(s_prgcode, Pgm_Type, txt_fix_code.Text, txt_com_name.Text, txt_pet_tel1.Text, txt_pet_date.Text _
                        , txt_pet_time.Text, txt_att_add1.Text, ddl_att_area.SelectedValue, ddl_att_road.SelectedValue _
                        , txt_work_floor.Text, Get_CKL_Click(ckl_work_floor_type), txt_work_date_s.Text, txt_work_date_e.Text _
                        , txt_att_cname1.Text, txt_att_name1.Text, txt_att_tel1.Text, txt_pet_name1.Text, txt_pet_name2.Text _
                        , txt_pet_tel3.Text, Get_CKL_Click(ckl_fix_type), Session("usr_code"), Format(Now, "yyyy/MM/dd"), txt_safe_item.Text _
                        , txt_dan_code.Text, txt_pet_name3.Text, txt_pet_name4.Text, txt_pet_tel5.Text, txt_pet_tel6.Text, txt_pet_tel4.Text _
                        , totheigh.Text, txt_work_type_memo.Text, rad_att_night.SelectedValue, rad_att_holiday.SelectedValue, txt_att_tel2.Text _
                        , rad_att_public.SelectedValue, radchoose.SelectedValue, rad_att_area.SelectedValue, l_pet_man.SelectedValue, txt_att_comment.Text _
                        , txt_pet_email.Text, chkEng.Text, chktoher.Text, txt_pet_name.Text, txt_pet_tel2.Text) Then

            '儲存使用紀錄
            ws.InsUsrRec(PrgMemo, Pgm_Type, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)

            If Pgm_Type <> "MDY" Then
                '取得最新條件
                If txtSubWhere.Text <> "" Then
                    txtSubWhere.Text &= " OR fix_code='" & RepSql(sKeyPrimary) & "' "
                Else
                    txtSubWhere.Text = " WHERE fix_code='" & RepSql(sKeyPrimary) & "' "
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
            If Chk_RelData(s_prgcode, "", txt_fix_code.Text) = False Then
                Chk_data = False
                Show_Message("有重覆參數名稱，請確認")
            End If

        End If
        '檢查資料正確性       
        If Chk_DateForm(txt_pet_date.Text) = False Or Chk_DateForm(txt_work_date_s.Text) = False Or Chk_DateForm(txt_work_date_e.Text) = False Then
            Chk_data = False
            Show_Message("日期格式不正確，請確認")
        End If

        '檢查必填欄位是否空白
        If txt_com_name.Text = "" Or ddl_att_area.SelectedValue = "" Or ddl_att_road.SelectedValue = "" Or txt_att_add1.Text = "" _
                                  Or txt_pet_tel1.Text = "" Or l_pet_man.SelectedValue = "" Or txt_pet_date.Text = "" Or txt_work_floor.Text = "" _
                                  Or txt_work_date_s.Text = "" Or txt_work_date_e.Text = "" Or txt_safe_item.Text = "" _
                                  Or (Get_CKL_Click(ckl_fix_type) = "" And chkEng.checked = False And chktoher.checked = False) Or txt_pet_time.Text = "" Then
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
            'Set_CKL_Option("work_floor", "", ckl_work_floor_type, "", "", "")
            Set_CKL_Option("fix_type", "", ckl_fix_type, "", "", "")
            '設定RadioButtonList
            Set_RBL_Option("YorN", "", rad_att_area, "", "", "B")
            Set_RBL_Option("YorN", "", rad_att_night, "", "", "B")
            Set_RBL_Option("YorN", "", rad_att_holiday, "", "", "B")
            Set_RBL_Option("YorN", "", rad_att_public, "", "", "B")

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
                    txt_pet_time.Text = Format(Now, "HH:mm")
                    txt_work_date_e.Text = Format(Now, "yyyy/MM/dd")
                    txt_work_date_s.Text = Format(Now, "yyyy/MM/dd")
                Case "MDY", "COPY"
                    Dim Dt_tmp As DataTable = Get_DataSet(s_prgcode, "Detail", txtPrimary.Text).Tables(0)
                    If Dt_tmp.Rows.Count > 0 Then
                        txt_fix_code.Text = Trim(Dt_tmp.Rows(0).Item("fix_code") & "")
                        txt_com_name.Text = Trim(Dt_tmp.Rows(0).Item("com_name") & "")
                        txt_pet_tel1.Text = Trim(Dt_tmp.Rows(0).Item("pet_tel1") & "")
                        l_pet_man.SelectedValue = Trim(Dt_tmp.Rows(0).Item("pet_man") & "")
                        txt_pet_name.Text = Trim(Dt_tmp.Rows(0).Item("att_man") & "")
                        txt_pet_tel2.Text = Trim(Dt_tmp.Rows(0).Item("att_tel") & "")
                        txt_pet_email.Text = Trim(Dt_tmp.Rows(0).Item("att_email") & "")
                        txt_pet_date.Text = Trim(Dt_tmp.Rows(0).Item("pet_date") & "")
                        txt_pet_time.Text = Trim(Dt_tmp.Rows(0).Item("pet_time") & "")
                        txt_att_add1.Text = Trim(Dt_tmp.Rows(0).Item("att_add1") & "")
                        ddl_att_area.SelectedValue = Trim(Dt_tmp.Rows(0).Item("att_area") & "")
                        radchoose.SelectedValue = Trim(Dt_tmp.Rows(0).Item("engchoosetype") & "")
                        rad_att_area.SelectedValue = Trim(Dt_tmp.Rows(0).Item("Workplacetype") & "")
                        rad_att_night.SelectedValue = Trim(Dt_tmp.Rows(0).Item("daytype") & "")
                        rad_att_holiday.SelectedValue = Trim(Dt_tmp.Rows(0).Item("holtype") & "")
                        rad_att_public.SelectedValue = Trim(Dt_tmp.Rows(0).Item("engtype") & "")
                        ddl_att_road.Items.Clear()
                        Set_DDL_Option(ddl_att_area.SelectedValue, "", ddl_att_road, "", "---請選擇---", "")
                        ddl_att_road.SelectedValue = Trim(Dt_tmp.Rows(0).Item("att_road") & "")
                        txt_work_floor.Text = Trim(Dt_tmp.Rows(0).Item("work_floor") & "")
                        'Set_CKL_Click(Trim(Dt_tmp.Rows(0).Item("work_floor_type") & ""), ckl_work_floor_type, "work_floor")
                        txt_work_date_s.Text = Trim(Dt_tmp.Rows(0).Item("work_date_s") & "")
                        txt_work_date_e.Text = Trim(Dt_tmp.Rows(0).Item("work_date_e") & "")
                        txt_att_cname1.Text = Trim(Dt_tmp.Rows(0).Item("att_cname1") & "")
                        txt_att_tel2.Text = Trim(Dt_tmp.Rows(0).Item("att_tel21") & "")
                        txt_att_comment.Text = Trim(Dt_tmp.Rows(0).Item("att_common") & "")
                        txt_att_name1.Text = Trim(Dt_tmp.Rows(0).Item("att_name1") & "")
                        txt_att_tel1.Text = Trim(Dt_tmp.Rows(0).Item("att_tel1") & "")
                        txt_pet_name1.Text = Trim(Dt_tmp.Rows(0).Item("att_cname2") & "")
                        txt_pet_name2.Text = Trim(Dt_tmp.Rows(0).Item("att_name2") & "")
                        txt_pet_tel3.Text = Trim(Dt_tmp.Rows(0).Item("att_tel2") & "")
                        txt_pet_tel4.Text = Trim(Dt_tmp.Rows(0).Item("att_mob2") & "")
                        txt_pet_name3.Text = Trim(Dt_tmp.Rows(0).Item("att_cname3") & "")
                        txt_pet_name4.Text = Trim(Dt_tmp.Rows(0).Item("att_name3") & "")
                        txt_pet_tel5.Text = Trim(Dt_tmp.Rows(0).Item("att_tel3") & "")
                        txt_pet_tel6.Text = Trim(Dt_tmp.Rows(0).Item("att_mob3") & "")
                        txt_safe_item.Text = Trim(Dt_tmp.Rows(0).Item("safe_item") & "")
                        txt_dan_code.Text = Trim(Dt_tmp.Rows(0).Item("dan_code") & "")
                        'chkEng.chedked = Trim(Dt_tmp.Rows(0).Item("att_chk") & "")
                        totheigh.Text = Trim(Dt_tmp.Rows(0).Item("totheigh") & "")
                        txt_work_type_memo.Text = Trim(Dt_tmp.Rows(0).Item("othertype") & "")
                        Set_CKL_Click(Trim(Dt_tmp.Rows(0).Item("fix_type") & ""), ckl_fix_type, "fix_type")
                        If Trim(Dt_tmp.Rows(0).Item("att_chk") & "") = "Y" Then
                            chkEng.checked = True
                        End If

                        If Trim(Dt_tmp.Rows(0).Item("att_chkoth") & "") = "Y" Then
                            chktoher.checked = True
                        End If

                    End If

                    If Pgm_type = "COPY" Then
                        txt_fix_code.Text = ""
                        txt_fix_code.Enabled = False
                    Else
                        txt_fix_code.Enabled = False
                        Bt_send.Visible = True
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

        Dim show2 As String
        Dim show3 As String

        If l_pet_man.SelectedValue = "0" Then
            show2 = "業主"
        ElseIf l_pet_man.SelectedValue = "1" Then
            show2 = "主辦機關"
        ElseIf l_pet_man.SelectedValue = "2" Then
            show2 = "施工廠商"
        End If

        If radchoose.SelectedValue = "1" Then
            show3 = "建築新工程"
        ElseIf radchoose.SelectedValue = "2" Then
            show3 = "建築修繕工程"
        ElseIf radchoose.SelectedValue = "3" Then
            show3 = "其他建築工程"
        ElseIf radchoose.SelectedValue = "4" Then
            show3 = "土木新建工程"
        ElseIf radchoose.SelectedValue = "5" Then
            show3 = "土木修繕工程"
        ElseIf radchoose.SelectedValue = "6" Then
            show3 = "其他土木工程"
        End If
        Dim show4 As String

        If rad_att_area.SelectedValue = "Y" Then
            show4 = "是"
        Else
            show4 = "否"
        End If

        '發送Mail通知
        Dim dtMail As New DataTable
        Dim column1 As New DataColumn
        Dim column2 As New DataColumn
        column1.DataType = System.Type.GetType("System.String")
        column2.DataType = System.Type.GetType("System.String")
        dtMail.Columns.Add(column1)
        dtMail.Columns.Add(column2)

        Dim rowMail As DataRow = dtMail.NewRow
        rowMail(0) = txt_pet_email.Text
        rowMail(1) = txt_pet_name.Text
        dtMail.Rows.Add(rowMail)

        '加入承辦人信箱
        Dim MailList As String() = Split(Get_SysPara("FI_receive"), "|")
		
        For i As Integer = 0 To UBound(MailList)-1
            If String.IsNullOrEmpty(MailList(i)) = False   Then
                Dim rowTmp As DataRow = dtMail.NewRow
                rowTmp(0) =MailList(i)
                rowTmp(1) = "裝修通報承辦人"
                dtMail.Rows.Add(rowTmp)
            End If
        Next
        '郵件內容
        Dim bFirst As Boolean = True
        'Dim sWorkFloorList As String = ""
        'For i As Integer = 0 To ckl_work_floor_type.Items.Count - 1
        '    If ckl_work_floor_type.Items(i).Selected = True Then
        '        If bFirst = False Then sWorkFloorList &= "、"
        '        bFirst = False
        '        sWorkFloorList &= ckl_work_floor_type.Items(i).Text
        '    End If
        'Next
        bFirst = True

        Dim sFixTypeList As String = ""
        For i As Integer = 0 To ckl_fix_type.Items.Count - 1
            If ckl_fix_type.Items(i).Selected = True Then
                If bFirst = False Then sFixTypeList &= "、"
                bFirst = False
                sFixTypeList &= ckl_fix_type.Items(i).Text
            End If
        Next
        Dim show5 As String = ""
        If chkEng.checked = True Then
            If sFixTypeList <> "" Then
                show5 = "、建築工程之外牆施工架拆組(總高度：" & totheigh.text & "公尺)"
            Else
                show5 = "建築工程之外牆施工架拆組(總高度：" & totheigh.text & "公尺)"
            End If

        End If

            Dim show6 As String = ""
        If chktoher.checked = True Then
            If sFixTypeList <> "" Or chkEng.checked = True Then
                show6 = "、其他：請加註：" & txt_work_type_memo.text
            Else
                show6 = "其他：請加註：" & txt_work_type_memo.text
            End If

        End If

            Dim show7 As String = ""
            If rad_att_night.SelectedValue = "Y" Then
                show7 = "是"
            Else
                show7 = "否"
            End If

            Dim show8 As String = ""
            If rad_att_holiday.SelectedValue = "Y" Then
                show8 = "是"
            Else
                show8 = "否"
            End If
        Dim sBody As String = "" & _
            "<table cellpadding=5 cellspacing=0 bordercolor=#7EBAD1 border=1 bgcolor=#ffffff style='border-collapse:collapse;'" & _
            "width='720' summary='排版表格' align=center class='c12'>" & _
            "<tr bgcolor='#7EBAD1' align='center'><td colspan='4'><p><strong>營繕、裝修工程及精準檢查通報</strong></p></td></tr>" & _
            "<tr>" & _
            "  <td width='20%'>通報單位：</td>" & _
            "  <td width='30%'>" & txt_com_name.Text & "</td>" & _
            "  <td width='20%'>聯　絡　電　話：</td>" & _
            "  <td width='30%'>" & txt_pet_tel1.Text & "</td>" & _
            "</tr>" & _
            "<tr>" & _
                "  <td>通　報　對　象：</td>" & _
                "  <td colspan='3'>" & show2 & "</td>" & _
                "<tr>" & _
            "<tr>" & _
            "  <td>通　報　日　期：</td>" & _
            "  <td>西元 " & Mid(txt_pet_date.Text, 1, 4) & " 年 " & _
                                        Mid(txt_pet_date.Text, 6, 2) & " 月 " & _
                                        Mid(txt_pet_date.Text, 9, 2) & " 日 </td>" & _
            "  <td>通　報　時　間：</td><td> " & Mid(txt_pet_time.Text, 1, 2) & " 時 " & _
                                        Mid(txt_pet_time.Text, 4, 2) & " 分</td>" & _
            "</tr>" & _
            "<tr>" & _
                "  <td width='20%'>通報人員：</td>" & _
                "  <td width='30%'>" & txt_pet_name.Text & "</td>" & _
                "  <td width='20%'>行　動　電　話：</td>" & _
                "  <td width='30%'>" & txt_pet_tel2.Text & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>工　程　地　址：</td>" & _
            "  <td colspan='3'>臺北市" & ddl_att_area.SelectedItem.Text & ddl_att_road.SelectedItem.Text & txt_att_add1.Text & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>工程名稱：</td>" & _
            "  <td colspan='3'>" & txt_safe_item.Text & "</td>" & _
            "</tr>" & _
            "<tr><td>工程種類：</td>" & _
                "<td colspan='3'>" & show3 & "</td></tr>" & _
            "<tr><td>是否為丁類危險場所：</td>" & _
            "<td colspan='3'>" & show4 & "</td></tr>" & _
            "<tr>" & _
                "  <td>可通報種類：</td>" & _
                "  <td colspan='3'>" & sFixTypeList & show5 & show6 & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>施工期間：</td>" & _
            "  <td colspan='3'>預計開工：西元 " & Mid(txt_work_date_s.Text, 1, 4) & " 年 " & _
                                        Mid(txt_work_date_s.Text, 6, 2) & " 月 " & _
                                        Mid(txt_work_date_s.Text, 9, 2) & " 日 " & _
            "                  <br />預計完工：西元 " & Mid(txt_work_date_e.Text, 1, 4) & " 年 " & _
                                        Mid(txt_work_date_e.Text, 6, 2) & " 月 " & _
                                        Mid(txt_work_date_e.Text, 9, 2) & " 日 </td>" & _
            "</tr>" & _
            "<tr><td>是否夜間施工</td><td colspan='3'>" & show7 & "</td></tr>" & _
            "<tr><td>是否六、日施工</td><td colspan='3'>" & show8 & "</td></tr>" & _
            "<tr>" & _
            "  <td>作業區域：</td>" & _
            "  <td colspan='3'>" & txt_work_floor.Text & "</td>" & _
            "</tr>" & _
             "<tr rowspan='2'>" & _
                "  <td width='20%'>業主資料：</td>" & _
                "  <td colspan='3'><table>" & _
                "  <tr><td>業主名稱</td><td>" & txt_att_cname1.Text & "</td>" & _
                "  <td>聯絡電話</td><td>" & txt_att_tel1.Text & "</td></tr>" & _
                "  <tr><td>聯絡人</td><td>" & txt_att_name1.Text & "</td>" & _
                "  <td>行動電話</td><td>" & txt_att_tel2.Text & "</td></tr>" & _
                "  </tr>" & _
           "</table></td></tr>" & _
           "<tr>" & _
                "  <td>備註：</td>" & _
                "  <td colspan='3'>" & txt_att_comment.text & "</td>" & _
                "</tr>" & _
            "</table>"

        If SendMail("臺北市勞動檢查處-申辦服務-營繕、裝修工程及精準檢查通報-送出成功", dtMail, sBody, s_prgcode) Then
            Show_Message("送出成功")
        Else
		
		 For Tmp_i As Integer = 0 To UBound(MailList)-1 
		 
            Show_Message(MailList(Tmp_i ) &   "-" & Tmp_i+1 &  ":"  & "送出失敗，請通知系統管理員")
		next
            Exit Sub
        End If
    End Sub
End Class
