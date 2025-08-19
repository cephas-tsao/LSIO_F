Imports System.Data

Partial Class basic_LSI011A_edit
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI011A"

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '自動編號
        If txtPgmType.Text <> "MDY" Then txt_self_code.Text = "SF" & Format(Now, "yyMMdd") & Right("0000" & Get_AutoIntMax("Self_check", "RIGHT(self_code, 10)", " WHERE self_code LIKE 'SF" & Format(Now, "yyMMdd") & "%'") + 1, 4)

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

        If SaveEditData(s_prgcode, Pgm_Type, txt_self_code.Text, txt_com_name.Text, txt_eng_name.Text, txt_eng_add.Text, _
                        txt_com_tel1.Text, txt_com_tel2.Text, txt_self_date.Text, txt_att_name.Text, txt_att_dep.Text, _
                        ddl_eng_area.SelectedValue, "", Session("usr_code"), Format(Now, "yyyy/MM/dd"), ddl_eng_road.SelectedValue) Then
            '儲存使用紀錄
            ws.InsUsrRec(PrgMemo, Pgm_Type, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)

            If Pgm_Type <> "MDY" Then
                '取得最新條件
                If txtSubWhere.Text <> "" Then
                    txtSubWhere.Text &= " OR self_code='" & RepSql(sKeyPrimary) & "' "
                Else
                    txtSubWhere.Text = " WHERE self_code='" & RepSql(sKeyPrimary) & "' "
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
            If Chk_RelData(s_prgcode, "", txt_self_code.Text) = False Then
                Chk_data = False
                Show_Message("有重覆參數名稱，請確認")
            End If

        End If
        '檢查資料正確性       
        If Chk_DateForm(txt_self_date.Text) = False Then
            Chk_data = False
            Show_Message("日期格式不正確，請確認")
        End If

        '檢查必填欄位是否空白
        If txt_com_name.Text = "" Or txt_eng_name.Text = "" Or txt_eng_add.Text = "" Or ddl_eng_area.SelectedValue = "" _
                                  Or ddl_eng_road.SelectedValue = "" Or txt_com_tel1.Text = "" Or txt_com_tel2.Text = "" _
                                  Or txt_self_date.Text = "" Or txt_att_name.Text = "" Or txt_att_dep.Text = "" Then
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
                    txt_self_date.Text = Format(Now, "yyyy/MM/dd")
                Case "MDY", "COPY"
                    Dim Dt_tmp As DataTable = Get_DataSet(s_prgcode, "Detail", txtPrimary.Text).Tables(0)
                    If Dt_tmp.Rows.Count > 0 Then
                        txt_self_code.Text = Trim(Dt_tmp.Rows(0).Item("self_code") & "")
                        txt_com_name.Text = Trim(Dt_tmp.Rows(0).Item("com_name") & "")
                        txt_eng_name.Text = Trim(Dt_tmp.Rows(0).Item("eng_name") & "")
                        txt_eng_add.Text = Trim(Dt_tmp.Rows(0).Item("eng_add") & "")
                        txt_com_tel1.Text = Trim(Dt_tmp.Rows(0).Item("com_tel1") & "")
                        txt_com_tel2.Text = Trim(Dt_tmp.Rows(0).Item("com_tel2") & "")
                        txt_self_date.Text = Trim(Dt_tmp.Rows(0).Item("self_date") & "")
                        txt_att_name.Text = Trim(Dt_tmp.Rows(0).Item("att_name") & "")
                        txt_att_dep.Text = Trim(Dt_tmp.Rows(0).Item("att_dep") & "")
                        lab_Start_path1.Text = Set_Files("File_LSI011A", Trim(Dt_tmp.Rows(0).Item("start_path1") & ""))
                        ddl_eng_area.SelectedValue = Trim(Dt_tmp.Rows(0).Item("eng_area") & "")
                        ddl_eng_area_SelectedIndexChanged(sender, e)
                        ddl_eng_road.SelectedValue = Trim(Dt_tmp.Rows(0).Item("eng_road") & "")
                    End If

                    If Pgm_type = "COPY" Then
                        txt_self_code.Text = ""
                        txt_self_code.Enabled = False
                    Else
                        txt_self_code.Enabled = False
                    End If
            End Select
        End If
        img_date1.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_self_date")
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

    Protected Sub ddl_eng_area_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddl_eng_area.SelectedIndexChanged
        ddl_eng_road.Items.Clear()
        Set_DDL_Option(ddl_eng_area.SelectedValue, "", ddl_eng_road, "", "---請選擇---", "")
    End Sub
End Class
