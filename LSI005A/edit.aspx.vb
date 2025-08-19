Imports System.Data

Partial Class basic_LSI005A_edit
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI005A"

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '自動編號
        If txtPgmType.Text <> "MDY" Then txt_sch_code.Text = "SC" & Format(Now, "yyMMdd") & Right("0000" & Get_AutoIntMax("e_paper_sch", "RIGHT(sch_code, 10)", " WHERE sch_code LIKE 'SC" & Format(Now, "yyMMdd") & "%'") + 1, 4)

        '檢查資料正確性
        If Chk_data() = False Then Exit Sub

        '設定Key欄位來源
        Dim sKeyPrimary As String = txt_sch_code.Text '*修改處*

        '設定記錄資料
        Dim sSql As String = Get_SqlStr(s_prgcode, "Detail", sKeyPrimary)
        Dim sFieldName As String = ws.Get_FieldName(s_prgcode)
        Dim sOldData As String = ws.Get_OldData(sSql, s_prgcode)
        Dim sNewData As String = Get_NewData(s_prgcode)

        '資料存檔
        Dim Pgm_Type As String = txtPgmType.Text
        Dim PrgMemo As String = RepSql(Label1.Text & ":" & sKeyPrimary)

        If SaveEditData(s_prgcode, Pgm_Type, txt_sch_code.Text, txt_paper_code.Text, txt_send_date.Text, ddl_send_time.SelectedValue, _
                        rbl_send_state.SelectedValue, Session("usr_code"), Format(Now, "yyyy/MM/dd"), "", "Y") Then
            '"sch_code,paper_code,send_date,send_time,send_state,last_date,usr_code,ins_date,count"
            ''***更新HTML資料************
            'detail_edit.Set_Paper_Html()
            '***************************
            Show_Message("存檔成功")
            '儲存使用紀錄
            ws.InsUsrRec(PrgMemo, Pgm_Type, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)

            If Pgm_Type <> "MDY" Then
                txtPgmType.Text = "MDY"
                detail_edit.setKeyName(sKeyPrimary)
                Panel1.Visible = True
                Bt_send.Visible = True
                Bt_Manual.Visible = True

                '取得最新條件
                If txtSubWhere.Text <> "" Then
                    txtSubWhere.Text &= " OR sch_code='" & RepSql(sKeyPrimary) & "' "
                Else
                    txtSubWhere.Text = " WHERE sch_code='" & RepSql(sKeyPrimary) & "' "
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

        '檢查重覆資料
        If txtPgmType.Text = "ADD" Then
            If Chk_RelData(s_prgcode, "", txt_sch_code.Text) = False Then
                Chk_data = False
                Show_Message("有重覆資料，請確認")
            End If
        End If

        '檢查必填欄位是否空白
        If txt_paper_name.Text = "" Or txt_send_date.Text = "" Then
            Chk_data = False
            Show_Message("必填欄位不得為空白，請確認")
        End If

        '檢查日期格式是否正確
        If Chk_DateForm(txt_send_date.Text) = False Then
            Chk_data = False
            Show_Message("日期格式錯誤，請確認")
        End If
        'If Chk_TimeForm(txt_tn_time_s.Text, ":") = False Or Chk_TimeForm(txt_tn_time_e.Text, ":") = False Then
        '    Chk_data = False
        '    Show_Message("時間格式錯誤，請確認")
        'End If
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '接收default頁面變數
        If Not IsPostBack Then
            Call SetClientCheck()

            '標頭
            Label1.Text = CType(PreviousPage.FindControl("Label1"), Label).Text
            '編輯模式
            txtPgmType.Text = CType(PreviousPage.FindControl("txtPgmType"), TextBox).Text
            '首頁條件
            txtSubWhere.Text = CType(PreviousPage.FindControl("txtsubwhere"), TextBox).Text
            txtOrder.Text = CType(PreviousPage.FindControl("txtOrder"), TextBox).Text
            '
            Set_RBL_Option("sendstate", "", rbl_send_state, "", "", "")

            '非新增模式則預設取得資料
            Dim Pgm_type As String = txtPgmType.Text
            txtPrimary.Text = CType(PreviousPage.FindControl("txtPrimary"), TextBox).Text
            txt_sch_code.Attributes.Add("ReadOnly", "ReadOnly")
            txt_paper_code.Attributes.Add("ReadOnly", "ReadOnly")
            txt_paper_name.Attributes.Add("ReadOnly", "ReadOnly")
            rbl_send_state.Attributes.Add("onclick", "return false;")
            '時間格式設定
            Set_DDL_time()

            Select Case Pgm_type
                Case "ADD"
                    txt_send_date.Text = Format(Now, "yyyy/MM/dd")
                    rbl_send_state.SelectedValue = "A"
                Case "MDY", "COPY"
                    Dim Dt_tmp As DataTable = Get_DataSet(s_prgcode, "Detail", txtPrimary.Text).Tables(0)
                    If Dt_tmp.Rows.Count > 0 Then
                        If Pgm_type = "COPY" Then
                            txt_sch_code.Text = ""
                        Else
                            txt_sch_code.Text = txtPrimary.Text
                            txt_sch_code.Enabled = False
                            detail_edit.setKeyName(txtPrimary.Text)
                            Panel1.Visible = True
                            Bt_send.Visible = True
                            Bt_Manual.Visible = True
                        End If
                        rbl_send_state.SelectedValue = Trim(Dt_tmp.Rows(0).Item("send_state") & "")
                        ddl_send_time.SelectedValue = Trim(Dt_tmp.Rows(0).Item("send_time") & "")
                        txt_send_date.Text = Trim(Dt_tmp.Rows(0).Item("send_date") & "")
                        txt_paper_code.Text = Trim(Dt_tmp.Rows(0).Item("paper_code") & "")
                        txt_paper_name.Text = ws.QueryData(txt_paper_code.Text, "e_paper", "paper_code", "paper_name")
                    End If
            End Select
        End If
        ib_usr_code.Attributes("onclick") = ws.Set_AddForm("../../public/addform.aspx", "EPC-1", "1|2", "txt_paper_code|txt_paper_name", "")
        img_date1.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_send_date")
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

    Protected Sub Err_Message(ByVal sMsg As String)
        Dim ErrMsg_Type As String = ws.Get_SysPara("ErrMsg_Type", "0")

        If ErrMsg_Type = "0" Or ErrMsg_Type = "2" Then
            '彈出式對話方塊
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('" & sMsg & "');</script>"))
        End If

        If ErrMsg_Type = "1" Or ErrMsg_Type = "2" Then
            '下方紅字訊息
            l_ErrMsg.Text = Replace(sMsg, "\n", "<br />")
        End If
    End Sub

    Protected Sub txt_usr_code_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt_paper_code.TextChanged
        txt_paper_name.Text = ws.QueryData(txt_paper_code.Text, "e_paper", "paper_code", "paper_name")
    End Sub
    Protected Sub Set_DDL_time()
        For i As Integer = 0 To 23
            ddl_send_time.Items.Add(New ListItem(i.ToString("00") & ":00"))
        Next
    End Sub

    Protected Sub Bt_send_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Bt_send.Click
        '取得Mail
        Dim dtSendLog As DataTable = Get_DataSet("SendNewsMail", "SendLog", txt_sch_code.Text).Tables(0)
        Dim sMessage As String = ""

        If dtSendLog.Rows.Count > 0 Then
            '取得電子報內容
            Dim dtNews As DataTable = Get_DataSet("SendNewsMail", "News", txt_paper_code.Text).Tables(0)

            If dtNews.Rows.Count > 0 Then
                Dim sTitle As String = dtNews.Rows(0).Item(0)
                Dim sBody As String = dtNews.Rows(0).Item(1)

                For i As Integer = 0 To dtSendLog.Rows.Count - 1
                    Dim sMail As String = dtSendLog.Rows(i).Item(0) & ""

                    If Chk_Email(sMail) = False Then
                        sMessage &= sMail & ":電子郵件格式錯誤\n"
                    Else
                        '發送Mail
                        Dim dtMail As New DataTable
                        Dim column1 As New DataColumn
                        Dim column2 As New DataColumn
                        column1.DataType = System.Type.GetType("System.String")
                        column2.DataType = System.Type.GetType("System.String")
                        dtMail.Columns.Add(column1)
                        dtMail.Columns.Add(column2)
                        Dim rowMail As DataRow = dtMail.NewRow
                        rowMail(0) = sMail
                        rowMail(1) = ""
                        dtMail.Rows.Add(rowMail)

                        If SendNewsMail(sTitle, dtMail, sBody, s_prgcode) Then
                            sMessage &= sMail & ":發送成功\n"
                        Else
                            sMessage &= sMail & ":發送失敗\n"
                        End If
                    End If
                Next

                '更新資料
                Dim sDate = ""
                '若非第一次發信，則寫入重送日期
                If Chk_RelData(s_prgcode, "", txt_sch_code.Text) = True Then sDate = Format(Now, "yyyy/MM/dd")
                SaveEditData(s_prgcode, "Result", txt_sch_code.Text, sDate)
            Else
                sMessage = "無電子報資料，請確認"
            End If
        Else
            sMessage = "無Mail資料，請確認"
        End If
            Err_Message(sMessage)
    End Sub

    Protected Sub Bt_Manual_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Bt_Manual.Click
        '新增隱藏排程工作'自動編號
        Dim sSchCode As String = "MA" & Format(Now, "yyMMdd") & Right("0000" & Get_AutoIntMax("e_paper_sch", "RIGHT(sch_code, 10)", _
                                 " WHERE sch_code LIKE 'MA" & Format(Now, "yyMMdd") & "%'") + 1, 4)

        SaveEditData(s_prgcode, "ADD", sSchCode, txt_paper_code.Text, Format(Now, "yyyy/MM/dd"), Format(Now, "HH:") & "00", _
                        "A", Session("usr_code"), Format(Now, "yyyy/MM/dd"), "", "N")
        '更新資料
        Dim sDate = ""
        '若非第一次發信，則寫入重送日期
        If Chk_RelData(s_prgcode, "", txt_sch_code.Text) = True Then sDate = Format(Now, "yyyy/MM/dd")
        SaveEditData(s_prgcode, "Result", txt_sch_code.Text, sDate)


        ''取得訂戶Mail List
        'Dim dtMailList As DataTable = Get_DataSet("SendNewsMail", "MailList").Tables(0)
        'Dim sMessage As String = ""
        'Dim bResult As String = "B"

        'If dtMailList.Rows.Count > 0 Then
        '    '取得電子報內容
        '    Dim dtNews As DataTable = Get_DataSet("SendNewsMail", "News", txt_paper_code.Text).Tables(0)

        '    If dtNews.Rows.Count > 0 Then
        '        Dim iTotalCount As Integer = 0
        '        Dim iFailCount As Integer = 0
        '        Dim sTitle As String = dtNews.Rows(0).Item(0)
        '        Dim sBody As String = dtNews.Rows(0).Item(1)

        '        For i As Integer = 0 To dtMailList.Rows.Count - 1
        '            iTotalCount += 1
        '            Dim sMail As String = dtMailList.Rows(i).Item(0) & ""
        '            Dim sName As String = dtMailList.Rows(i).Item(1) & ""

        '            If Chk_Email(sMail) = False Then
        '                bResult = "C"
        '                sMessage &= sMail & ":電子郵件格式錯誤\n"
        '                '寫入Log檔
        '                SaveEditData("LSI005A_detail", "ADD", "", txt_sch_code.Text, sMail, sName)
        '            Else
        '                '發送Mail
        '                Dim dtMail As New DataTable
        '                Dim column1 As New DataColumn
        '                Dim column2 As New DataColumn
        '                column1.DataType = System.Type.GetType("System.String")
        '                column2.DataType = System.Type.GetType("System.String")
        '                dtMail.Columns.Add(column1)
        '                dtMail.Columns.Add(column2)
        '                Dim rowMail As DataRow = dtMail.NewRow
        '                rowMail(0) = sMail
        '                rowMail(1) = sName
        '                dtMail.Rows.Add(rowMail)

        '                If SendNewsMail(sTitle, dtMail, sBody, s_prgcode) Then
        '                    'sMessage &= sMail & ":發送成功\n"
        '                Else
        '                    iFailCount += 1
        '                    sMessage &= sMail & ":發送失敗\n"
        '                End If
        '            End If
        '        Next

        '        sMessage &= "總共發送: " & iTotalCount & " 件\n"
        '        sMessage &= "發送成功: " & iTotalCount - iFailCount & " 件\n"
        '        sMessage &= "發送失敗 " & iFailCount & " 件\n"

        '        Dim sDate = ""
        '        '若非第一次發信，則寫入重送日期
        '        If Chk_RelData(s_prgcode, "", txt_sch_code.Text) = True Then sDate = Format(Now, "yyyy/MM/dd")
        '        '更新發送標籤
        '        SaveEditData("SendNewsMail", "Result", txt_sch_code.Text, bResult, sDate)
        '        rbl_send_state.SelectedValue = bResult
        '    Else
        '        sMessage = "無電子報資料，請確認"
        '    End If
        'Else
        '    sMessage = "無Mail資料，請確認"
        'End If
        'Err_Message(sMessage)
    End Sub
End Class

