Imports System.Data

Partial Class management_BDP150A_edit
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "BDP150A"

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

        If SaveEditData(s_prgcode, Pgm_Type, txtPrimary.Text, ddl_ip_type.SelectedValue, txt_ip_name.Text, txt_ip_s1.Text, txt_ip_s2.Text, txt_ip_s3.Text, _
                         txt_ip_s4.Text, txt_ip_e1.Text, txt_ip_e2.Text, txt_ip_e3.Text, txt_ip_e4.Text) Then
            '儲存使用紀錄
            ws.InsUsrRec(PrgMemo, Pgm_Type, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)

            If Pgm_Type <> "MDY" Then
                '取得最新條件
                If txtSubWhere.Text <> "" Then
                    txtSubWhere.Text &= " OR serial_no='" & RepSql(sKeyPrimary) & "' "
                Else
                    txtSubWhere.Text = " WHERE serial_no='" & RepSql(sKeyPrimary) & "' "
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
        'If txtPgmType.Text = "ADD" Then
        '    If Chk_RelData(s_prgcode, "", txt_par_name.Text) = False Then
        '        Chk_data = False
        '        Show_Message("有重覆參數名稱，請確認")
        '    End If
        'End If

        '檢查必填欄位是否空白
        If txt_ip_e1.Text = "" Or txt_ip_e2.Text = "" Or txt_ip_e3.Text = "" Or txt_ip_e4.Text = "" _
            Or txt_ip_s1.Text = "" Or txt_ip_s2.Text = "" Or txt_ip_s3.Text = "" Or txt_ip_s4.Text = "" Or ddl_ip_type.SelectedValue = "" Then
            Chk_data = False
            Show_Message("必填欄位不得為空白，請確認")
        End If
        '檢查IP輸入是否為數字
        If CHK_val(txt_ip_s1.Text) = False Or CHK_val(txt_ip_s2.Text) = False Or CHK_val(txt_ip_s3.Text) = False Or CHK_val(txt_ip_s4.Text) = False Or _
             CHK_val(txt_ip_e1.Text) = False Or CHK_val(txt_ip_e2.Text) = False Or CHK_val(txt_ip_e3.Text) = False Or CHK_val(txt_ip_e4.Text) = False Then
            Chk_data = False
            Show_Message("IP輸入不正確，請確認")
        End If
        
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '接收default頁面變數
        If Not IsPostBack Then
            SetClientCheck()
            Call Set_DDL_Option("IP_tp", "", ddl_ip_type, "", "---", "")
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
                    ddl_ip_type.SelectedValue = "B"
                Case "MDY", "COPY"
                    Dim Dt_tmp As DataTable = Get_DataSet(s_prgcode, "Detail", txtPrimary.Text).Tables(0)
                    If Dt_tmp.Rows.Count > 0 Then
                        txt_ip_s1.Text = Trim(Dt_tmp.Rows(0).Item("ip_s1") & "")
                        txt_ip_s2.Text = Trim(Dt_tmp.Rows(0).Item("ip_s2") & "")
                        txt_ip_s3.Text = Trim(Dt_tmp.Rows(0).Item("ip_s3") & "")
                        txt_ip_s4.Text = Trim(Dt_tmp.Rows(0).Item("ip_s4") & "")
                        txt_ip_e1.Text = Trim(Dt_tmp.Rows(0).Item("ip_e1") & "")
                        txt_ip_e2.Text = Trim(Dt_tmp.Rows(0).Item("ip_e2") & "")
                        txt_ip_e3.Text = Trim(Dt_tmp.Rows(0).Item("ip_e3") & "")
                        txt_ip_e4.Text = Trim(Dt_tmp.Rows(0).Item("ip_e4") & "")
                        ddl_ip_type.SelectedValue = Trim(Dt_tmp.Rows(0).Item("ip_type") & "")
                        txt_ip_name.Text = Trim(Dt_tmp.Rows(0).Item("ip_name") & "")
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
    Function CHK_val(ByVal p_val As String) As Boolean
        CHK_val = True
        If Val(p_val) < 0 Or Val(p_val) > 255 Then
            CHK_val = False
        End If
    End Function
End Class
