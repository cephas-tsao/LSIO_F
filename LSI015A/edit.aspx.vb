Imports System.Data

Partial Class basic_LSI015A_edit
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI015A"

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '自動編號
        If txtPgmType.Text <> "MDY" Then txt_danc_code.Text = "DC" & Format(Now, "yyMMdd") & Right("0000" & Get_AutoIntMax("Danger_check", "RIGHT(danc_code, 10)", " WHERE danc_code LIKE 'DC" & Format(Now, "yyMMdd") & "%'") + 1, 4)

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

        If SaveEditData(s_prgcode, Pgm_Type, txt_danc_code.Text, txt_eng_name.Text, txt_pet_date.Text, txt_pet_time.Text, txt_eng_add.Text _
                        , Get_CKL_Click(ckl_danc_kind), rbl_danc_step.SelectedValue, txt_danc_div.Text, rbl_danc_type.SelectedValue, txt_com_name.Text _
                        , txt_com_idno.Text, txt_com_add.Text, txt_com_tel1.Text, txt_att_name1.Text, txt_att_name2.Text, txt_con_name.Text, txt_con_tel1.Text _
                        , txt_con_tel2.Text, txt_con_email.Text, rbl_danc_state.SelectedValue, rbl_result.SelectedValue, Session("usr_code"), Format(Now, "yyyy/MM/dd") _
                        , ddl_att_area.SelectedValue, ddl_att_road.SelectedValue, txt_con_pw.Text) Then

            '儲存使用紀錄
            ws.InsUsrRec(PrgMemo, Pgm_Type, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)

            If Pgm_Type <> "MDY" Then
                '取得最新條件
                If txtSubWhere.Text <> "" Then
                    txtSubWhere.Text &= " OR danc_code='" & RepSql(sKeyPrimary) & "' "
                Else
                    txtSubWhere.Text = " WHERE danc_code='" & RepSql(sKeyPrimary) & "' "
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
            If Chk_RelData(s_prgcode, "", txt_danc_code.Text) = False Then
                Chk_data = False
                Show_Message("有重覆參數名稱，請確認")
            End If

        End If
        '檢查資料正確性       
        If Chk_DateForm(txt_pet_date.Text) = False Then
            Chk_data = False
            Show_Message("日期格式不正確，請確認")
        End If
        If Chk_TimeForm(txt_pet_time.Text) = False Then
            Chk_data = False
            Show_Message("時間格式不正確，請確認")
        End If
        If Chk_CompanyNo(txt_com_idno.Text) = False Then
            Chk_data = False
            Show_Message("統編不正確，請確認")
        End If
        If Chk_Email(txt_con_email.Text) = False Then
            Chk_data = False
            Show_Message("MAIL格式不正確，請確認")
        End If

        '檢查必填欄位是否空白
        If txt_com_name.Text = "" Or ddl_att_area.SelectedValue = "" Or ddl_att_road.SelectedValue = "" Or txt_com_add.Text = "" _
                                  Or txt_eng_name.Text = "" Or txt_pet_date.Text = "" Or txt_pet_time.Text = "" Or txt_eng_add.Text = "" _
                                  Or txt_com_idno.Text = "" Or txt_com_tel1.Text = "" Or txt_att_name1.Text = "" _
                                  Or txt_con_name.Text = "" Or txt_att_name2.Text = "" Or txt_con_tel1.Text = "" _
                                  Or txt_con_tel2.Text = "" Or txt_con_email.Text = "" Or ckl_danc_kind.SelectedValue = "" _
                                  Or rbl_danc_step.SelectedValue = "" Or rbl_danc_type.SelectedValue = "" Then
            Chk_data = False
            Show_Message("必填欄位不得為空白，請確認")
        End If
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '接收default頁面變數
        If Not IsPostBack Then
            SetClientCheck()
            '設定下拉式選單
            Set_DDL_Option("ZipCode", "110", ddl_att_area, "", "---請選擇---", "")
            Set_DDL_Option(ddl_att_area.SelectedValue, "", ddl_att_road, "", "---請選擇---", "")

            Set_RBL_Option("result", "", rbl_result, "", "", "")
            Set_RBL_Option("danc_state", "", rbl_danc_state, "", "", "")
            Set_CKL_Option("danc_kind", "", ckl_danc_kind, "", "", "")
            Set_RBL_Option("danc_type", "", rbl_danc_type, "", "", "")
            Set_RBL_Option("danc_step", "", rbl_danc_step, "", "", "")

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
                Case "MDY", "COPY"
                    Dim Dt_tmp As DataTable = Get_DataSet(s_prgcode, "Detail", txtPrimary.Text).Tables(0)
                    If Dt_tmp.Rows.Count > 0 Then
                        txt_danc_code.Text = Trim(Dt_tmp.Rows(0).Item("danc_code") & "")
                        txt_eng_name.Text = Trim(Dt_tmp.Rows(0).Item("eng_name") & "")
                        txt_pet_date.Text = Trim(Dt_tmp.Rows(0).Item("pet_date") & "")
                        txt_pet_time.Text = Trim(Dt_tmp.Rows(0).Item("pet_time") & "")
                        txt_eng_add.Text = Trim(Dt_tmp.Rows(0).Item("eng_add") & "")
                        rbl_danc_step.SelectedValue = Trim(Dt_tmp.Rows(0).Item("danc_step") & "")
                        txt_danc_div.Text = Trim(Dt_tmp.Rows(0).Item("danc_div") & "")
                        rbl_danc_type.SelectedValue = Trim(Dt_tmp.Rows(0).Item("danc_type") & "")
                        txt_com_name.Text = Trim(Dt_tmp.Rows(0).Item("com_name") & "")
                        txt_com_idno.Text = Trim(Dt_tmp.Rows(0).Item("com_idno") & "")
                        txt_com_add.Text = Trim(Dt_tmp.Rows(0).Item("com_add") & "")
                        txt_com_tel1.Text = Trim(Dt_tmp.Rows(0).Item("com_tel1") & "")
                        txt_att_name1.Text = Trim(Dt_tmp.Rows(0).Item("att_name1") & "")
                        txt_att_name2.Text = Trim(Dt_tmp.Rows(0).Item("att_name2") & "")
                        txt_con_name.Text = Trim(Dt_tmp.Rows(0).Item("con_name") & "")
                        txt_con_tel1.Text = Trim(Dt_tmp.Rows(0).Item("con_tel1") & "")
                        txt_con_tel2.Text = Trim(Dt_tmp.Rows(0).Item("con_tel2") & "")
                        Set_CKL_Click(Trim(Dt_tmp.Rows(0).Item("danc_kind") & ""), ckl_danc_kind, "danc_kind")
                        txt_con_email.Text = Trim(Dt_tmp.Rows(0).Item("con_email") & "")
                        rbl_danc_state.SelectedValue = Trim(Dt_tmp.Rows(0).Item("danc_state") & "")
                        rbl_result.SelectedValue = Trim(Dt_tmp.Rows(0).Item("result") & "")
                        ddl_att_area.SelectedValue = Trim(Dt_tmp.Rows(0).Item("att_area") & "")
                        ddl_att_area_SelectedIndexChanged(sender, e)
                        ddl_att_road.SelectedValue = Trim(Dt_tmp.Rows(0).Item("att_road") & "")
                        txt_con_pw.Text = Trim(Dt_tmp.Rows(0).Item("con_pw") & "")
                    End If

                    If Pgm_type = "COPY" Then
                        txt_danc_code.Text = ""
                        txt_danc_code.Enabled = False
                    Else
                        txt_danc_code.Enabled = False
                        If Chk_Email(txt_con_email.Text) Then
                            Bt_send.Visible = True
                        End If
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
        '加入聯絡人Mail
        Dim rowMail As DataRow = dtMail.NewRow
        rowMail(0) = txt_con_email.Text
        rowMail(1) = txt_con_name.Text
        dtMail.Rows.Add(rowMail)
        '加入承辦人信箱
        Dim MailList As String() = Split(Get_SysPara("DC_receive"), "|")
        For i As Integer = 0 To UBound(MailList)
            If String.IsNullOrEmpty(MailList(i)) = False Then
                Dim rowTmp As DataRow = dtMail.NewRow
                rowTmp(0) = MailList(i)
                rowTmp(1) = "危評審查承辦人"
                dtMail.Rows.Add(rowTmp)
            End If
        Next
        '郵件內容
        Dim bFirst As Boolean = True
        Dim sDancKindList As String = ""
        For i As Integer = 0 To ckl_danc_kind.Items.Count - 1
            If ckl_danc_kind.Items(i).Selected = True Then
                If bFirst = False Then sDancKindList &= "、"
                bFirst = False
                sDancKindList &= ckl_danc_kind.Items(i).Text
            End If
        Next
        Dim sDancStep As String = ""
        If rbl_danc_step.SelectedValue = "A" Then
            sDancStep = rbl_danc_step.SelectedItem.Text
        Else
            sDancStep = rbl_danc_step.SelectedItem.Text & "(第" & txt_danc_div.Text & "階段)"
        End If
        Dim sBody As String = "" & _
            "<table cellpadding=5 cellspacing=0 bordercolor=#7EBAD1 border=1 bgcolor=#ffffff style='border-collapse:collapse;'" & _
            "width='720' summary='排版表格' align=center class='c12'>" & _
            "<tr bgcolor='#7EBAD1' align='center'><td colspan='4'><p><strong>危險性工作場所基本資料</strong></p></td></tr>" & _
            "<tr>" & _
            "  <td>工　程　名　稱：</td>" & _
            "  <td colspan='3'>" & txt_eng_name.Text & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>事業單位名稱：</td>" & _
            "  <td colspan='3'>" & txt_com_name.Text & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>公　司　地　址：</td>" & _
            "  <td colspan='3'>臺北市" & ddl_att_area.SelectedItem.Text & ddl_att_road.SelectedItem.Text & txt_com_add.Text & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td width='20%'>公　司　電　話：</td>" & _
            "  <td width='30%'>" & txt_com_tel1.Text & "</td>" & _
            "  <td width='20%'>事業單位統編：</td>" & _
            "  <td width='30%'>" & txt_com_idno.Text & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>事業經營負責人：</td>" & _
            "  <td>" & txt_att_name1.Text & "</td>" & _
            "  <td>工作場所負責人：</td>" & _
            "  <td>" & txt_att_name2.Text & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>聯　　絡　　人：</td>" & _
            "  <td colspan='3'>姓　　名：" & txt_con_name.Text & "　室內電話：" & txt_con_tel1.Text & "<br />" & _
                               "行動電話：" & txt_con_tel2.Text & "　　E-Mail：" & txt_con_email.Text & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>通　報　日　期：</td>" & _
            "  <td>西元 " & Mid(txt_pet_date.Text, 1, 4) & " 年 " & _
                                        Mid(txt_pet_date.Text, 6, 2) & " 月 " & _
                                        Mid(txt_pet_date.Text, 9, 2) & " 日 </td>" & _
            "  <td>通　報　時　間：</td><td> " & Mid(txt_pet_time.Text, 1, 2) & " 時 " & _
                                        Mid(txt_pet_time.Text, 4, 2) & " 分</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>丁類危險性工作<br />場所類別：</td>" & _
            "  <td colspan='3'>" & sDancKindList & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>整體或分階段<br />送審：</td>" & _
            "  <td colspan='3'>" & sDancStep & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>送　審　類　別：</td>" & _
            "  <td colspan='3'>" & rbl_danc_type.SelectedItem.Text & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>查詢進度密碼：</td>" & _
            "  <td colspan='3'>" & txt_con_pw.Text & "</td>" & _
            "</tr>" & _
            "</table>"

        If SendMail("臺北市勞動檢查處-申辦服務-危評審查-送出成功", dtMail, sBody, s_prgcode) Then
            Show_Message("送出成功")
        Else
            Show_Message("送出失敗，請通知系統管理員")
            Exit Sub
        End If
    End Sub
End Class
