Imports System.Data
Imports Microsoft.Security.Application

Partial Class basic_LSI003A_edit
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "Downtime" 'Downtime

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        txt_Down_range.Attributes.Add("maxlength", "300")
        txt_Down_range.Attributes.Add("onkeyup", "return ismaxlength(this)")
        txt_Down_reason.Attributes.Add("maxlength", "300")
        txt_Down_reason.Attributes.Add("onkeyup", "return ismaxlength(this)")
        txt_Down_remarks.Attributes.Add("maxlength", "300")
        txt_Down_remarks.Attributes.Add("onkeyup", "return ismaxlength(this)")

        '接收default頁面變數
        If Not IsPostBack Then
            SetClientCheck()

            'Set_DDL_Option("book_autho", "1", ddl_book_author, "", "---請選擇---", "")
            'Set_DDL_Option("areacode", "are01", ddl_book_area, "", "---請選擇---", "B")
            'Set_DDL_Option(ddl_book_area.SelectedValue, "", ddl_book_road, "", "---請選擇---", "")
            'chk_com_type1_1.Text = "勞工安全衛生人員"

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
            'Sanitizer.GetSafeHtmlFragment(AntiXss.HtmlEncode(Format(Now, "yyyy/MM/dd")))
            Select Case Pgm_type
                Case "ADD"
                    txt_Down_cDate.Text = Format(Now, "yyyy/MM/dd")
                    txt_Down_creatorID.Text = Session("usr_code")
                Case "MDY", "COPY"
                    Dim Dt_tmp As DataTable = Get_DataSet(s_prgcode, "Detail", txtPrimary.Text).Tables(0)

                    If Dt_tmp.Rows.Count > 0 Then
                        txt_Down_code.Text = Trim(Dt_tmp.Rows(0).Item("Down_code") & "")
                        txt_Down_institutions.Text = Trim(Dt_tmp.Rows(0).Item("Down_institutions") & "")
                        txt_Down_title.Text = Trim(Dt_tmp.Rows(0).Item("Down_title") & "")
                        txt_Down_address.Text = Trim(Dt_tmp.Rows(0).Item("Down_address") & "")
                        txt_Down_X.Text = Trim(Dt_tmp.Rows(0).Item("Down_X") & "")
                        txt_Down_Y.Text = Trim(Dt_tmp.Rows(0).Item("Down_Y") & "")
                        txt_Down_stopDate.Text = Trim(Dt_tmp.Rows(0).Item("Down_stopDate") & "")
                        txt_Down_returnDate.Text = Trim(Dt_tmp.Rows(0).Item("Down_returnDate") & "")
                        txt_Down_range.Text = Trim(Dt_tmp.Rows(0).Item("Down_range") & "")
                        txt_Down_reason.Text = Trim(Dt_tmp.Rows(0).Item("Down_reason") & "")
                        txt_Down_remarks.Text = Trim(Dt_tmp.Rows(0).Item("Down_remarks") & "")
                        'txt_Down_sDate.Text = Trim(Dt_tmp.Rows(0).Item("Down_sDate") & "")
                        'txt_Down_eDate.Text = Trim(Dt_tmp.Rows(0).Item("Down_eDate") & "")
                        'txt_Down_pDate.Text = Trim(Dt_tmp.Rows(0).Item("Down_pDate") & "")
                        txt_Down_cDate.Text = Trim(Dt_tmp.Rows(0).Item("Down_cDate") & "")
                        txt_Down_uDate.Text = Trim(Dt_tmp.Rows(0).Item("Down_uDate") & "")
                        txt_Down_creatorID.Text = Trim(Dt_tmp.Rows(0).Item("Down_creatorID") & "")
                        txt_Down_modifyID.Text = IIf(Trim(Dt_tmp.Rows(0).Item("Down_modifyID") & "") = "0", "", Trim(Dt_tmp.Rows(0).Item("Down_modifyID") & ""))
                        'txt_Down_creatorIP.Text = Trim(Dt_tmp.Rows(0).Item("Down_creatorIP") & "")
                        'txt_Down_modifyIP.Text = Trim(Dt_tmp.Rows(0).Item("Down_modifyIP") & "")
                        'Set_CKL_Click(Trim(Dt_tmp.Rows(0).Item("book_print") & ""), ck_book_print, "book_print") '版本
                    End If

                    If Pgm_type = "COPY" Then
                        txt_Down_code.Text = ""
                        txt_Down_code.Enabled = False
                    Else
                        txt_Down_code.Enabled = False
                    End If
            End Select
        End If

        img_date2.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_Down_stopDate")
        img_date3.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_Down_returnDate")
        l_ErrMsg.Text = ""
    End Sub

    '存檔
    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '自動編號
        'If txtPgmType.Text <> "MDY" Then txt_Down_code.Text = "BK" & Format(Now, "yyMMdd") & Right("0000" & Get_AutoIntMax("yumi2", "RIGHT(Down_code, 10)", " WHERE Down_code LIKE 'BK" & Format(Now, "yyMMdd") & "%'") + 1, 4) '****

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

        'If SaveEditData(s_prgcode, Pgm_Type, txt_Down_institutions.Text, _
        '        txt_Down_title.Text, txt_Down_address.Text, _
        '        txt_Down_stopDate.Text, txt_Down_returnDate.Text, txt_Down_range.Text, _
        '        txt_Down_reason.Text, txt_Down_remarks.Text, _
        '        Session("usr_code"), _
        '        txt_Down_code.Text, txt_Down_X.Text, txt_Down_Y.Text) Then
        If SaveEditData(s_prgcode, Pgm_Type, txt_Down_institutions.Text, _
                txt_Down_title.Text, txt_Down_address.Text, _
                txt_Down_stopDate.Text, txt_Down_returnDate.Text, txt_Down_range.Text, _
                txt_Down_reason.Text, txt_Down_remarks.Text, _
                Session("usr_code"), _
                txt_Down_code.Text, txt_Down_X.Text, txt_Down_Y.Text) Then
            '儲存使用紀錄
            ws.InsUsrRec(PrgMemo, Pgm_Type, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)

            If Pgm_Type <> "MDY" Then
                '取得最新條件
                If txtSubWhere.Text <> "" Then
                    txtSubWhere.Text &= " OR Down_code='" & RepSql(sKeyPrimary) & "' "
                Else
                    txtSubWhere.Text = " WHERE Down_code='" & RepSql(sKeyPrimary) & "' "
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
        'If txtPgmType.Text = "ADD" Or txtPgmType.Text = "COPY" Then
        '    If Chk_RelData(s_prgcode, "", txt_Down_code.Text) = False Then
        '        Chk_data = False
        '        Show_Message("有重覆參數名稱，請確認")
        '    End If
        'End If

        '檢查資料正確性 
        'If Chk_Tel(txt_per_tel1.Text) = False Then
        '    Chk_data = False
        '    Show_Message("行動電話格式不正確，請確認")
        'End If
        If Chk_DateForm(txt_Down_cDate.Text) = False Then
            Chk_data = False
            Show_Message("日期格式不正確，請確認")
        End If

        '檢查必填欄位是否空白
        If txt_Down_institutions.Text = "" Then
            Chk_data = False
            Show_Message("必填欄位不得為空白，請確認")
        End If
    End Function

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

    '儲存作者
    Protected Function Get_Author_List() As String
        Dim sStr As String = ""
        Dim bFirst As Boolean = True
        If bFirst <> True Then sStr &= "|"
        bFirst = False
        'sStr &= ddl_book_author.SelectedValue

        Return sStr
    End Function

End Class
