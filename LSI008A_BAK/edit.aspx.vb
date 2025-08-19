Imports System.Data

Partial Class basic_LSI008A_edit
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI008A"
    Dim s_prgname As String = ws.Get_PrgName(s_prgcode)

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        'If Session("UsrListDo") = "N" Then  '避免reload重覆執行
        '    Session("UsrListDo") = "Y"
        '    If txtPgmType.Text <> "MDY" Then
        '        'Response.Write("<script>opener.window.location.href = opener.window.location.href;window.close();</script>")
        '        txt_mrd_code.Text = Session("LSI008A_MrdCode")
        '        txtPgmType.Text = "MDY"
        '        detail_edit.setKeyName(txt_mrd_code.Text)
        '        Panel1.Visible = True
        '    End If
        '    'txtSaved.Text = "YES"
        '    Exit Sub
        'End If

        '自動編號
        If txtPgmType.Text <> "MDY" Then txt_mrd_code.Text = "MR" & Format(Now, "yyMMdd") & Right("0000" & Get_AutoIntMax("Meeting_order", "RIGHT(mrd_code, 10)", " WHERE mrd_code LIKE 'MR" & Format(Now, "yyMMdd") & "%'") + 1, 4)

        '檢查資料正確性
        If Chk_data() = False Then Exit Sub

        '設定Key欄位來源
        Dim sKeyPrimary As String = txt_mrd_code.Text '*修改處*

        '設定記錄資料
        Dim sSql As String = Get_SqlStr(s_prgcode, "Detail", sKeyPrimary)
        Dim sFieldName As String = ws.Get_FieldName(s_prgcode)
        Dim sOldData As String = ws.Get_OldData(sSql, s_prgcode)
        Dim sNewData As String = Get_NewData(s_prgcode)

        '資料存檔
        Dim Pgm_Type As String = txtPgmType.Text
        Dim PrgMemo As String = RepSql(Label1.Text & ":" & sKeyPrimary)

        If SaveEditData(s_prgcode, Pgm_Type, txt_mrd_code.Text, ddl_room_code.SelectedValue, txt_mrd_date.Text, _
                        ddl_mrd_time_s.SelectedValue, ddl_mrd_time_e.SelectedValue, txt_usr_code.Text, txt_mrd_name.Text, _
                        txt_mrd_host.Text, txt_mrd_memo.Text, Get_CKL_Click(ckl_send_list), ddl_is_show.SelectedValue) Then
            '儲存使用紀錄
            ws.InsUsrRec(PrgMemo, Pgm_Type, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)

            If Pgm_Type <> "MDY" Then
                txtPgmType.Text = "MDY"
                detail_edit.setKeyName(sKeyPrimary)
                Panel1.Visible = True
            End If

            If Pgm_Type = "MDY" Then
                '發送mail 寫在這
                Show_Message("發送mail成功")
            End If

            'txtSaved.Text = "YES"
            Show_Message("儲存成功!")
                '存檔成功重整父頁面
                Response.Write("<script>opener.window.location.href = opener.window.location.href;</script>")
            Else
                '存檔出錯就秀出錯誤訊息
                Show_Message("儲存失敗!")
        End If
        'Session("UsrListDo") = "Y"
    End Sub

    Protected Function Chk_data() As Boolean
        'Server端資料檢查 自訂函數
        Chk_data = True

        ''檢查重覆資料
        If txtPgmType.Text = "ADD" Then
            If Chk_RelData(s_prgcode, "MrdCode", txt_mrd_code.Text) = False Then
                Chk_data = False
                Show_Message("有重覆參數資料，請確認")
            End If
        End If

        '檢查資料正確性
        If txt_mrd_date.Text < Format(Now, "yyyy/MM/dd") Then
            Show_Message("開會日期不能選擇小於今日，請確認")
            Return False
        End If
        If Chk_RelData(s_prgcode, "MrdTime", txt_mrd_code.Text, ddl_room_code.SelectedValue, txt_mrd_date.Text, ddl_mrd_time_s.SelectedValue, ddl_mrd_time_e.SelectedValue) = False Then
            Show_Message("您所選擇時段已有人租用，請確認")
            Return False
        End If
        If ddl_mrd_time_e.SelectedValue <= ddl_mrd_time_s.SelectedValue Then
            Chk_data = False
            Show_Message("租用截止時間需大於租用開始時間，請確認")
        End If
        'If not (txt_mrd_code.Text.Contains("危評") ) 
        '	Show_Message("不是危評會議，請將危評選項選項改為否!")
        'end if 


        ''檢查必填欄位是否空白
        'If txt_room_code.Text = "" Or txt_mrd_code.Text = "" Then
        '    Chk_data = False
        '    Show_Message("必填欄位不得為空白，請確認")
        'End If
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '接收default頁面變數
        If Not IsPostBack Then
            Set_DDL_Option("", "3A", ddl_room_code, Get_SqlStr(s_prgcode, "RoomCode"), "", "")
            'Set_DDL_Option("", "", ddl_usr_code, Get_SqlStr(s_prgcode, "UsrCode"), "", "")
            Set_CKL_Option("Send_List", "", ckl_send_list, "", "", "")
            Set_DDL_Option("YorN", "N", ddl_is_show, "", "", "")
            Set_DDL_time()

            '標頭
            Label1.Text = s_prgname & "(" & s_prgcode & ")"

            '設定唯讀欄位
            txt_mrd_date.Attributes.Add("ReadOnly", "ReadOnly")
            txt_usr_name.Attributes.Add("ReadOnly", "ReadOnly")

            '設定隱藏欄位
            txt_usr_code.Attributes.Add("style", "display:none")

            '取得資訊
            txtPgmType.Text = Request.QueryString("type")
            txt_mrd_code.Text = Request.QueryString("key")
            txt_mrd_date.Text = Request.QueryString("date")

            '避免reload重覆執行
            If Session("UsrListDo") = "Y" Then
                If txtPgmType.Text <> "MDY" Then
                    txt_mrd_code.Text = Session("LSI008A_MrdCode")
                    txtPgmType.Text = "MDY"
                End If
                Session("UsrListDo") = "N"
            End If

            Select Case txtPgmType.Text
                Case "ADD"
                    txt_usr_code.Text = Session("usr_code")
                    txt_usr_name.Text = Session("usr_name")

                Case "MDY", "COPY"
                    Dim Dt_tmp As DataTable = Get_DataSet(s_prgcode, "Detail", txt_mrd_code.Text).Tables(0)
                    If Dt_tmp.Rows.Count > 0 Then
                        'txt_mrd_code.Text = Trim(Dt_tmp.Rows(0).Item("mrd_code") & "")
                        txt_mrd_date.Text = Trim(Dt_tmp.Rows(0).Item("mrd_date") & "")
                        ddl_room_code.SelectedValue = Trim(Dt_tmp.Rows(0).Item("room_code") & "")
                        ddl_mrd_time_s.SelectedValue = Trim(Dt_tmp.Rows(0).Item("mrd_time_s") & "")
                        ddl_mrd_time_e.SelectedValue = Trim(Dt_tmp.Rows(0).Item("mrd_time_e") & "")
                        'ddl_usr_code.SelectedValue = Trim(Dt_tmp.Rows(0).Item("usr_code") & "")
                        txt_usr_code.Text = Trim(Dt_tmp.Rows(0).Item("usr_code") & "")
                        txt_usr_name.Text = ws.QueryData(txt_usr_code.Text, "BDP080", "usr_code", "usr_name")
                        txt_mrd_name.Text = Trim(Dt_tmp.Rows(0).Item("mrd_name") & "")
                        txt_mrd_host.Text = Trim(Dt_tmp.Rows(0).Item("mrd_host") & "")
                        ddl_is_show.SelectedValue = Trim(Dt_tmp.Rows(0).Item("is_show") & "")
                        txt_mrd_memo.Text = Trim(Dt_tmp.Rows(0).Item("mrd_memo") & "")
                        Set_CKL_Click(Trim(Dt_tmp.Rows(0).Item("send_list") & ""), ckl_send_list, "Send_List")
                    End If

                    If txtPgmType.Text = "COPY" Then
                        txt_mrd_code.Text = ""
                    Else
                        detail_edit.setKeyName(txt_mrd_code.Text)
                        Panel1.Visible = True
                        img_date1.Visible = True
                    End If

                    If Session("usr_code") <> txt_usr_code.Text Then
                        Button1.Visible = False
                        Panel1.Visible = False
                    End If

                    '科室管理者與系統管理員可強制修改
                    If Session("grp_code") = "G001" Or Session("grp_code") = "G002" Then
                        Button1.Visible = True
                        Panel1.Visible = True
                    End If

                Case Else
                    Button1.Visible = False
                    Button1.Enabled = False
                    '關閉視窗
                    Response.Write("<script>window.close();</script>")
            End Select

            txt_mrd_code.Enabled = False
        End If

        l_ErrMsg.Text = ""
        img_date1.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_mrd_date")
    End Sub

    Protected Sub Bt_BackUp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Bt_BackUp.Click
        'If txtSaved.Text <> "" Then
        '    '關閉視窗並重整父頁面
        '    Response.Write("<script>opener.window.location.href = opener.window.location.href;window.close();</script>")
        'Else
        '關閉視窗
        Response.Write("<script>window.close();</script>")
        'End If
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

    Protected Sub Set_DDL_time()
        Dim sSpan As String = ws.Get_SysPara("LSI008A_TimeSpan", "15")

        '00 整點,30半點,15一刻

        Select Case True
            Case sSpan = "15"
                For i As Integer = 8 To 22
                    ddl_mrd_time_s.Items.Add(New ListItem(i.ToString("00") & ":00"))
                    ddl_mrd_time_s.Items.Add(New ListItem(i.ToString("00") & ":15"))
                    ddl_mrd_time_s.Items.Add(New ListItem(i.ToString("00") & ":30"))
                    ddl_mrd_time_s.Items.Add(New ListItem(i.ToString("00") & ":45"))
                    ddl_mrd_time_e.Items.Add(New ListItem(i.ToString("00") & ":00"))
                    ddl_mrd_time_e.Items.Add(New ListItem(i.ToString("00") & ":15"))
                    ddl_mrd_time_e.Items.Add(New ListItem(i.ToString("00") & ":30"))
                    ddl_mrd_time_e.Items.Add(New ListItem(i.ToString("00") & ":45"))
                Next
            Case sSpan = "30"
                For i As Integer = 8 To 22
                    ddl_mrd_time_s.Items.Add(New ListItem(i.ToString("00") & ":00"))
                    ddl_mrd_time_s.Items.Add(New ListItem(i.ToString("00") & ":30"))
                    ddl_mrd_time_e.Items.Add(New ListItem(i.ToString("00") & ":00"))
                    ddl_mrd_time_e.Items.Add(New ListItem(i.ToString("00") & ":30"))
                Next
            Case sSpan = "00"
                For i As Integer = 8 To 22
                    ddl_mrd_time_s.Items.Add(New ListItem(i.ToString("00") & ":00"))
                    ddl_mrd_time_e.Items.Add(New ListItem(i.ToString("00") & ":00"))
                Next
        End Select
    End Sub
End Class
