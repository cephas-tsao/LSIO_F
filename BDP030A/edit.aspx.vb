Imports System.Data

Partial Class management_BDP030A_edit
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "BDP030A"

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
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

        If SaveEditData(s_prgcode, Pgm_Type, txt_prg_code.Text, _
                        txt_prg_name.Text, ddl_prg_type.SelectedValue, ddl_is_use.SelectedValue, _
                        ddl_is_group.SelectedValue, txt_scr_no.Text, txt_menu_level.Text, _
                        ddl_up_level.SelectedValue, ddl_prg_folder.SelectedValue, txt_s_date.Text, _
                        txt_e_date.Text _
                        ) Then
            '儲存使用紀錄
            ws.InsUsrRec(PrgMemo, Pgm_Type, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)

            If Pgm_Type <> "MDY" Then
                '取得最新條件
                If txtSubWhere.Text <> "" Then
                    txtSubWhere.Text &= " OR prg_code='" & RepSql(sKeyPrimary) & "' "
                Else
                    txtSubWhere.Text = " WHERE prg_code='" & RepSql(sKeyPrimary) & "' "
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
            If Chk_RelData(s_prgcode, "", txt_prg_code.Text) = False Then
                Chk_data = False
                Show_Message("有重覆編號資料，請確認")
            End If
        End If

        '檢查必填欄位是否空白
        If txt_prg_code.Text = "" Or txt_prg_name.Text = "" Then
            Chk_data = False
            Show_Message("必填欄位不得為空白，請確認")
        End If

        '檢查日期格式是否正確
        If Chk_DateForm(txt_s_date.Text) = False Then
            Chk_data = False
            Show_Message("日期格式錯誤，請確認")
        End If

        If Chk_DateForm(txt_e_date.Text) = False Then
            Chk_data = False
            Show_Message("日期格式錯誤，請確認")
        End If


    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '接收default頁面變數
        If Not IsPostBack Then
            SetClientCheck()

            '控項預設值
            Call Set_DDL_Option("PrgType", "", ddl_prg_type, "", "", "")
            Call Set_DDL_Option("IsUse", "", ddl_is_use, "", "", "")
            Call Set_DDL_Option("YorN", "", ddl_is_group, "", "", "")
            Call Set_DDL_Option("folder", "", ddl_prg_folder, "", "--請選擇--", "B")
            Call Set_DDL_Option("", "", ddl_up_level, Get_SqlStr("BDP030option1", "", Request("prg_code")), "--請選擇--", "B")

            '日期欄位處理
            txt_s_date.Text = Format(Now, "yyyy/MM/dd")
            txt_e_date.Text = Format(Now, "yyyy/MM/dd")

            img_date1.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_s_date")
            img_date2.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_e_date")

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
                        txt_prg_name.Text = Trim(Dt_tmp.Rows(0).Item("prg_name") & "")
                        ddl_prg_type.SelectedValue = Trim(Dt_tmp.Rows(0).Item("prg_type") & "")
                        ddl_is_use.SelectedValue = Trim(Dt_tmp.Rows(0).Item("is_use") & "")
                        ddl_is_group.SelectedValue = Trim(Dt_tmp.Rows(0).Item("is_group") & "")
                        txt_scr_no.Text = Trim(Dt_tmp.Rows(0).Item("scr_no") & "")
                        txt_menu_level.Text = Trim(Dt_tmp.Rows(0).Item("menu_level") & "")
                        ddl_up_level.SelectedValue = Trim(Dt_tmp.Rows(0).Item("up_level") & "")
                        ddl_prg_folder.SelectedValue = Trim(Dt_tmp.Rows(0).Item("prg_folder") & "")
                        txt_s_date.Text = Trim(Dt_tmp.Rows(0).Item("s_date") & "")
                        txt_e_date.Text = Trim(Dt_tmp.Rows(0).Item("e_date") & "")
                    End If

                    If Pgm_type = "COPY" Then
                        txt_prg_code.Text = ""
                        txt_prg_code.Enabled = True
                    Else
                        txt_prg_code.Text = txtPrimary.Text
                        txt_prg_code.Enabled = False
                    End If
            End Select
        End If
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
