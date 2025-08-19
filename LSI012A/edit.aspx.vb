Imports System.Data
Imports System.Web
Imports System.Text.RegularExpressions

Partial Class basic_LSI012A_edit
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI012A"

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '自動編號
        If txtPgmType.Text <> "MDY" Then txt_sma_code.Text = "SS" & Format(Now, "yyMMdd") & Right("0000" & Get_AutoIntMax("Small_space", "RIGHT(sma_code, 10)", " WHERE sma_code LIKE 'SS" & Format(Now, "yyMMdd") & "%'") + 1, 4)

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

        If SaveEditData(s_prgcode, Pgm_Type, txt_sma_code.Text, txt_com_name.Text, ddl_eng_area.SelectedValue, txt_eng_add.Text _
                        , txt_eng_date_s.Text, txt_eng_date_e.Text, txt_pet_name.Text, txt_pet_date.Text, txt_pet_tel1.Text _
                        , txt_pet_mail.Text, txt_att_name1.Text, txt_att_add1.Text, txt_att_tel1.Text, txt_att_name2.Text _
                        , txt_att_add2.Text, txt_att_tel2.Text, txt_att_name3.Text, txt_att_add3.Text, txt_att_tel3.Text _
                        , Get_CKL_Click(ckl_pet_list), txt_pet_other.Text, Session("usr_code"), Format(Now, "yyyy/MM/dd"), ddl_eng_road.SelectedValue) Then
            '儲存使用紀錄
            ws.InsUsrRec(PrgMemo, Pgm_Type, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)

            If Pgm_Type <> "MDY" Then
                '取得最新條件
                If txtSubWhere.Text <> "" Then
                    txtSubWhere.Text &= " OR sma_code='" & RepSql(sKeyPrimary) & "' "
                Else
                    txtSubWhere.Text = " WHERE sma_code='" & RepSql(sKeyPrimary) & "' "
                End If
            End If
            '存檔正確導回default頁面
            Response.Redirect("Default.aspx")


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
            If Chk_RelData(s_prgcode, "", txt_sma_code.Text) = False Then
                Chk_data = False
                Show_Message("有重覆參數名稱，請確認")
            End If

        End If
        '檢查資料正確性       
        If Chk_DateForm(txt_pet_date.Text) = False Or Chk_DateForm(txt_eng_date_s.Text) = False Or Chk_DateForm(txt_eng_date_e.Text) = False Then
            Chk_data = False
            Show_Message("日期格式不正確，請確認")
        End If
        If Chk_Email(txt_pet_mail.Text) = False Then
            Chk_data = False
            Show_Message("信箱格式不正確，請確認")
        End If

        '檢查必填欄位是否空白
        If txt_com_name.Text = "" Or ddl_eng_area.SelectedValue = "" Or ddl_eng_road.SelectedValue = "" Or txt_eng_add.Text = "" _
                                  Or txt_eng_date_s.Text = "" Or txt_eng_date_e.Text = "" Or txt_pet_name.Text = "" _
                                  Or txt_pet_date.Text = "" Or txt_pet_tel1.Text = "" Or txt_pet_mail.Text = "" Then
            Chk_data = False
            Show_Message("必填欄位不得為空白，請確認")
        End If
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '接收default頁面變數
        If Not IsPostBack Then
            SetClientCheck()
            Set_DDL_Option("ZipCode", "110", ddl_eng_area, "", "---請選擇---", "")
            Set_DDL_Option(ddl_eng_area.SelectedValue, "", ddl_eng_road, "", "---請選擇---", "")
            Set_CKL_Option("Pet_List", "", ckl_pet_list, "", "", "")

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
                    txt_eng_date_s.Text = Format(Now, "yyyy/MM/dd")
                    txt_eng_date_e.Text = Format(Now, "yyyy/MM/dd")
                Case "MDY", "COPY"
                    Dim Dt_tmp As DataTable = Get_DataSet(s_prgcode, "Detail", txtPrimary.Text).Tables(0)
                    If Dt_tmp.Rows.Count > 0 Then
                        txt_sma_code.Text = Trim(Dt_tmp.Rows(0).Item("sma_code") & "")
                        txt_com_name.Text = Trim(Dt_tmp.Rows(0).Item("com_name") & "")
                        txt_eng_add.Text = Trim(Dt_tmp.Rows(0).Item("eng_add") & "")
                        txt_eng_date_s.Text = Trim(Dt_tmp.Rows(0).Item("eng_date_s") & "")
                        txt_eng_date_e.Text = Trim(Dt_tmp.Rows(0).Item("eng_date_e") & "")
                        txt_pet_name.Text = Trim(Dt_tmp.Rows(0).Item("pet_name") & "")
                        txt_pet_date.Text = Trim(Dt_tmp.Rows(0).Item("pet_date") & "")
                        txt_pet_tel1.Text = Trim(Dt_tmp.Rows(0).Item("pet_tel1") & "")
                        txt_pet_mail.Text = Trim(Dt_tmp.Rows(0).Item("pet_mail") & "")
                        txt_att_name1.Text = Trim(Dt_tmp.Rows(0).Item("att_name1") & "")
                        txt_att_add1.Text = Trim(Dt_tmp.Rows(0).Item("att_add1") & "")
                        txt_att_tel1.Text = Trim(Dt_tmp.Rows(0).Item("att_tel1") & "")
                        txt_att_name2.Text = Trim(Dt_tmp.Rows(0).Item("att_name2") & "")
                        txt_att_add2.Text = Trim(Dt_tmp.Rows(0).Item("att_add2") & "")
                        txt_att_tel2.Text = Trim(Dt_tmp.Rows(0).Item("att_tel2") & "")
                        txt_att_name3.Text = Trim(Dt_tmp.Rows(0).Item("att_name3") & "")
                        txt_att_add3.Text = Trim(Dt_tmp.Rows(0).Item("att_add3") & "")
                        txt_att_tel3.Text = Trim(Dt_tmp.Rows(0).Item("att_tel3") & "")
                        Set_CKL_Click(Trim(Dt_tmp.Rows(0).Item("pet_list") & ""), ckl_pet_list, "Pet_List")
                        txt_pet_other.Text = Trim(Dt_tmp.Rows(0).Item("pet_other") & "")
                        ddl_eng_area.SelectedValue = Trim(Dt_tmp.Rows(0).Item("eng_area") & "")
                        ddl_eng_area_SelectedIndexChanged(sender, e)
                        ddl_eng_road.SelectedValue = Trim(Dt_tmp.Rows(0).Item("eng_road") & "")
                    End If

                    If Pgm_type = "COPY" Then
                        txt_sma_code.Text = ""
                        txt_sma_code.Enabled = False
                    Else
                        txt_sma_code.Enabled = False
                        Bt_send.Visible = True
                    End If

            End Select
        End If
                    img_date1.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_pet_date")
        img_date2.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_eng_date_s")
        img_date3.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_eng_date_e")
        l_ErrMsg.Text = ""
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
        'rowMail(0) = txt_att_mail.Text
        rowMail(1) = txt_att_name1.Text
        dtMail.Rows.Add(rowMail)

        '加入承辦人信箱
        Dim MailList As String() = Split(Get_SysPara("PC_receive"), "|")
        For i As Integer = 0 To UBound(MailList)
            If String.IsNullOrEmpty(MailList(i)) = False Then
                Dim rowTmp As DataRow = dtMail.NewRow
                rowTmp(0) = MailList(i)
                rowTmp(1) = "局限空間申報承辦人"
                dtMail.Rows.Add(rowTmp)
            End If
        Next
        '郵件內容
        Dim bFirst As Boolean = True

        bFirst = True


        Dim nSpilt = Split(Trim(txt_pet_date.Text & ""), "/")

        Dim sBody As String = "您好，局限空間申報通報 在日期" & RepSql(txt_pet_date.Text) & "，已通報成功，若需查詢，請與臺北市勞檢處連絡。謝謝。" &
                "<table cellpadding=5 cellspacing=0 bordercolor=#7EBAD1 border=1 bgcolor=#ffffff style='border-collapse:collapse;'" &
                "width='720' summary='排版表格' align=center class='c12'>" &
                "<tr bgcolor='#7EBAD1' align='center'><td colspan='20'><p><strong>局限空間申報通報</strong></p></td></tr>" &
                "<tr><td colspan='20' bgcolor='#9EDAF1' align='center'><font color=red>事業單位資料</font></td></tr>" &
                "<tr>" &
                "  <td colspan='4'>通報單位：</td>" &
                "  <td colspan='16'>" & RepSql(txt_com_name.Text) & "</td>" &
                "</tr>" &
                "<tr>" &
                "  <td colspan='4'>聯絡人：</td>" &
                "  <td colspan='6'>" & RepSql(txt_att_name1.Text) & "</td>" &
        "  <td colspan='4'>電子信箱：</td>" &
        "  <td colspan='6'>" & RepSql(txt_pet_mail.Text) & "</td>" &
        "</tr>" &
        "<tr>" &
        "  <td colspan='4'>聯絡電話：</td>" &
        "  <td colspan='6'>" & RepSql(txt_att_tel1.Text) & "</td>" &
        "  <td colspan='4'>手機：</td>" &
        "  <td colspan='6'>" & txt_att_tel2.Text & "</td>" &
        "</tr>" &
        "<tr>" &
        "  <td colspan='4'>通報日期：</td>" &
        "  <td colspan='16'>西元 " & nSpilt(0) & " 年 " &
                                    nSpilt(1) & " 月 " &
                                    nSpilt(2) & " 日 </td>" &
        "</tr>" &
        "<tr><td colspan='20' bgcolor='#9EDAF1' align='center'><font color=red>承攬資料</font></td></tr>" &
        "<tr>" &
        "  <td colspan='4'>工程名稱：</td>" &
        "  <td colspan='6'>" & RepSql(txt_com_name.Text) & "</td>" &
        "  <td colspan='4'>建號：</td>" &
        "  <td colspan='6'>" & txt_sma_code.Text & "</td>" &
        "</tr>" &
        "<tr>" &
        "  <td colspan='4'>工程地址：</td>" &
        "  <td colspan='16'>臺北市" & ddl_eng_area.SelectedItem.Text & ddl_eng_road.SelectedItem.Text & RepSql(txt_eng_add.Text) & "</td>" &
        "</tr>" &
        "</table>"


        If SendMail("臺北市勞動檢查處-申辦服務-局限空間申報通報-送出成功", dtMail, sBody, s_prgcode) Then
            Show_Message("送出成功")
        Else
            Show_Message("送出失敗，請通知系統管理員")
            Exit Sub
        End If
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

    Protected Sub ddl_eng_area_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddl_eng_area.SelectedIndexChanged
        ddl_eng_road.Items.Clear()
        Set_DDL_Option(ddl_eng_area.SelectedValue, "", ddl_eng_road, "", "---請選擇---", "")
    End Sub
End Class
