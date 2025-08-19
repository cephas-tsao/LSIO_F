Imports System.Data

Partial Class basic_LSI020A_edit
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI020A"

      Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '自動編號
        If txtPgmType.Text <> "MDY" Then txt_fix_code.Text = "SH" & Format(Now, "yyMMdd") & Right("0000" & Get_AutoIntMax("SCHOOL_WORK", "RIGHT(fix_code, 10)", " WHERE fix_code LIKE 'SH" & Format(Now, "yyMMdd") & "%'") + 1, 4)
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

        If SaveEditData(s_prgcode, Pgm_Type, txt_fix_code.Text, ddl_com_name.SelectedValue, Format(Now, "yyyy/MM/dd") _
                        , Format(Now, "HH:mm"), txt_work_floor.Text, Get_CKL_Click(ckl_work_floor_type), txt_work_date_s.Text, txt_work_date_e.Text _
                        , txt_att_cname1.Text, txt_att_name1.Text, txt_att_tel1.Text, txt_att_cname2.Text, txt_att_name2.Text _
                        , txt_att_tel2.Text, Get_CKL_Click(ckl_fix_type), Format(Now, "yyyy/MM/dd"), txt_safe_item.Text, txt_work_money.Text, ddl_att_area.SelectedValue, Session("usr_code"), txt_att_mail.Text) Then

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
        If Chk_DateForm(txt_work_date_s.Text) = False Or Chk_DateForm(txt_work_date_e.Text) = False Then
            Chk_data = False
            Show_Message("日期格式不正確，請確認")
        End If

        '檢查必填欄位是否空白
        If txt_safe_item.Text = "" Or txt_work_floor.Text = "" Or txt_work_money.Text = "" Or txt_work_date_s.Text = "" Or txt_work_date_e.Text = "" Or txt_work_money.Text = "" Or Get_CKL_Click(ckl_fix_type) = "" Then
            Chk_data = False
            Show_Message("必填欄位不得為空白，請確認")
        End If
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '接收default頁面變數
        If Not IsPostBack Then
            SetClientCheck()
            '設定下拉式選單
            Set_DDL_Option("ZipCode1", "110", ddl_att_area, "", "", "B")
            Set_DDL_Option(ddl_att_area.SelectedValue, "", ddl_com_name, "", "", "B")
            '設定CheckBoxList
            Set_CKL_Option("work_floor", "", ckl_work_floor_type, "", "", "B")
            Set_CKL_Option("fix_type1", "", ckl_fix_type, "", "", "B")

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
                    'txt_pet_date.Text = Format(Now, "yyyy/MM/dd")
                    'txt_pet_time.Text = Format(Now, "HH:mm")
                    txt_work_date_e.Text = Format(Now, "yyyy/MM/dd")
                    txt_work_date_s.Text = Format(Now, "yyyy/MM/dd")
                Case "MDY", "COPY"
                    Dim Dt_tmp As DataTable = Get_DataSet(s_prgcode, "Detail", txtPrimary.Text).Tables(0)
                    If Dt_tmp.Rows.Count > 0 Then
                        txt_fix_code.Text = Trim(Dt_tmp.Rows(0).Item("fix_code") & "")
                        ddl_att_area.SelectedValue = Trim(Dt_tmp.Rows(0).Item("att_area") & "")
                        ddl_com_name.Items.Clear()
                        Set_DDL_Option(ddl_att_area.SelectedValue, "", ddl_com_name, "", "---請選擇---", "")
                        ddl_com_name.SelectedValue = Trim(Dt_tmp.Rows(0).Item("com_name") & "")
                        txt_work_floor.Text = Trim(Dt_tmp.Rows(0).Item("work_floor") & "")
                        Set_CKL_Click(Trim(Dt_tmp.Rows(0).Item("work_floor_type") & ""), ckl_work_floor_type, "work_floor")
                        txt_work_date_s.Text = Trim(Dt_tmp.Rows(0).Item("work_date_s") & "")
                        txt_work_date_e.Text = Trim(Dt_tmp.Rows(0).Item("work_date_e") & "")
                        txt_att_cname1.Text = Trim(Dt_tmp.Rows(0).Item("att_cname1") & "")
                        txt_att_name1.Text = Trim(Dt_tmp.Rows(0).Item("att_name1") & "")
                        txt_att_tel1.Text = Trim(Dt_tmp.Rows(0).Item("att_tel1") & "")
                        txt_att_cname2.Text = Trim(Dt_tmp.Rows(0).Item("att_cname2") & "")
                        txt_att_name2.Text = Trim(Dt_tmp.Rows(0).Item("att_name2") & "")
                        txt_att_tel2.Text = Trim(Dt_tmp.Rows(0).Item("att_tel2") & "")
                        txt_safe_item.Text = Trim(Dt_tmp.Rows(0).Item("safe_item") & "")
                        txt_work_money.Text = Trim(Dt_tmp.Rows(0).Item("work_money") & "")
                        txt_att_mail.Text = Trim(Dt_tmp.Rows(0).Item("att_mail") & "")
                        Set_CKL_Click(Trim(Dt_tmp.Rows(0).Item("fix_type") & ""), ckl_fix_type, "fix_type")
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
        'img_date1.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_pet_date")
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
        ddl_com_name.Items.Clear()
        Set_DDL_Option(ddl_att_area.SelectedValue, "", ddl_com_name, "", "---請選擇---", "B")

        'ddl_att_road.Items.Clear()
        'Set_DDL_Option(ddl_att_area.SelectedValue, "", ddl_att_road, "", "---請選擇---", "")
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
        rowMail(0) = txt_att_mail.Text
        rowMail(1) = txt_att_cname1.Text
        dtMail.Rows.Add(rowMail)

        '加入承辦人信箱
        Dim MailList As String() = Split(Get_SysPara("FI_receive"), "|")
        For i As Integer = 0 To UBound(MailList)
            If String.IsNullOrEmpty(MailList(i)) = False Then
                Dim rowTmp As DataRow = dtMail.NewRow
                rowTmp(0) = MailList(i)
                rowTmp(1) = "校園營繕工程暨修繕通報承辦人"
                dtMail.Rows.Add(rowTmp)
            End If
        Next
        '郵件內容
        Dim bFirst As Boolean = True
        Dim sWorkFloorList As String = ""
        For i As Integer = 0 To ckl_work_floor_type.Items.Count - 1
            If ckl_work_floor_type.Items(i).Selected = True Then
                If bFirst = False Then sWorkFloorList &= "、"
                bFirst = False
                sWorkFloorList &= ckl_work_floor_type.Items(i).Text
            End If
        Next
        bFirst = True
        Dim sFixTypeList As String = ""
        For i As Integer = 0 To ckl_fix_type.Items.Count - 1
            If ckl_fix_type.Items(i).Selected = True Then
                If bFirst = False Then sFixTypeList &= "、"
                bFirst = False
                sFixTypeList &= ckl_fix_type.Items(i).Text
            End If
        Next
        Dim sBody As String = "" & _
                "<table cellpadding=5 cellspacing=0 bordercolor=#7EBAD1 border=1 bgcolor=#ffffff style='border-collapse:collapse;'" & _
                "width='720' summary='排版表格' align=center class='c12'>" & _
                "<tr bgcolor='#7EBAD1' align='center'><td colspan='4'><p><strong>校園營繕工程暨修繕作業</strong></p></td></tr>" & _
                 "<tr>" & _
                "  <td>施　工　地　點：</td>" & _
                "  <td colspan='3'>臺北市" & ddl_att_area.SelectedItem.Text & ddl_com_name.SelectedItem.Text() & "</td>" & _
                "</tr>" & _
                "  <tr><td>合約工程名稱：</td>" & _
                "  <td colspan='3'>" & txt_safe_item.Text & "</td>" & _
                "</tr>" & _
                 "<tr>" & _
                "  <td>作業區域及樓層：</td>" & _
                "  <td>" & txt_work_floor.Text & "</td>" & _
                "  <td colspan='2'>" & sWorkFloorList & "</td>" & _
                "</tr>" & _
                "<tr>" & _
                "  <td>種　　　　　類：</td>" & _
                "  <td colspan='3'>" & sFixTypeList & "</td>" & _
                "</tr>" & _
                "<tr>" & _
                "  <td>施　工　期　間：</td>" & _
               "  <td colspan='3'>開工：西元 " & Mid(txt_work_date_s.Text, 1, 4) & " 年 " & _
                                        Mid(txt_work_date_s.Text, 6, 2) & " 月 " & _
                                        Mid(txt_work_date_s.Text, 9, 2) & " 日 " & _
              "                  <br />完工：西元 " & Mid(txt_work_date_e.Text, 1, 4) & " 年 " & _
                                        Mid(txt_work_date_e.Text, 6, 2) & " 月 " & _
                                        Mid(txt_work_date_e.Text, 9, 2) & " 日 </td>" & _
                "</tr>" & _
              "<tr>" & _
                "  <td>承　攬　金　額：</td>" & _
                "  <td colspan='3'>" & txt_work_money.Text & "</td>" & _
                "</tr>" & _
                  "<tr>" & _
                "  <td>通　報　日　期：</td>" & _
                "  <td>西元 " & Format(Now, "yyyy/MM/dd") & " </td>" & _
                "  <td>通　報　時　間：</td><td> " & Format(Now, "HH 時 mm 分") & "</td>" & _
                "</tr>" & _
                "<tr>" & _
                "  <td>聯　絡　資　料：<br />(如有未詳者免填)</td>" & _
                "  <td colspan='3'><table summary='排版用表格' width='100%' border='0' >" & _
                "      <tr><td align='center'>類別</td><td>名稱</td><td>聯絡人/工地負責人</td><td>電話</td></tr><tr>" & _
                "        <th>主辦單位</th>" & _
                "        <td>" & txt_att_cname1.Text & "</td>" & _
                "        <td>" & txt_att_name1.Text & "</td>" & _
                "        <td>" & txt_att_tel1.Text & "</td>" & _
                "      </tr><tr>" & _
                "        <th>承覽廠商</th>" & _
                "        <td>" & txt_att_cname2.Text & "</td>" & _
                "        <td>" & txt_att_name2.Text & "</td>" & _
                "        <td>" & txt_att_tel2.Text & "</td>" & _
                "      </tr></table></td>" & _
                "</tr>" & _
                "</table>"

        If SendMail("臺北市勞動檢查處-申辦服務-校園營繕工程暨修繕作業-送出成功", dtMail, sBody, s_prgcode) Then
            Show_Message("送出成功")
        Else
            Show_Message("送出失敗，請通知系統管理員")
            Exit Sub
        End If
    End Sub
End Class
