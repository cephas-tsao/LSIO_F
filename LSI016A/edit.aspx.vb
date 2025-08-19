Imports System.Data

Partial Class basic_LSI016A_edit
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI016A"

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '自動編號
        If txtPgmType.Text <> "MDY" Then txt_news_code.Text = "NW" & Format(Now, "yyMMdd") & Right("0000" & Get_AutoIntMax("e_news", "RIGHT(news_code, 10)", " WHERE news_code LIKE 'NW" & Format(Now, "yyMMdd") & "%'") + 1, 4)

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

        '上傳圖片

        If fup_news_pic1.HasFile = True Then
            If UploadFile(fup_news_pic1, "File_LSI016A", txt_news_code.Text & ".jpg", s_prgcode) = True Then
                txt_news_pic1.Text = txt_news_code.Text & ".jpg"
            End If
        End If


        If SaveEditData(s_prgcode, Pgm_Type, txt_news_code.Text, txt_news_title.Text, rbl_news_type.SelectedValue, txt_news_contect.Text _
                        , txt_news_pic1.Text, txt_usr_code.Text, txt_ins_date.Text) Then

            '儲存使用紀錄
            ws.InsUsrRec(PrgMemo, Pgm_Type, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)

            If Pgm_Type <> "MDY" Then
                txtPgmType.Text = "MDY"
                Bt_preview.Visible = True
                Bt_preview.Attributes("onclick") = "subwindow=window.open('preview.aspx?code=" & txt_news_code.Text & "','preview','height=768,width=1024,top=200,left=400,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no','');subwindow.focus();return false;"

                '    '取得最新條件
                '    If txtSubWhere.Text <> "" Then
                '        txtSubWhere.Text &= " OR news_code='" & RepSql(sKeyPrimary) & "' "
                '    Else
                '        txtSubWhere.Text = " WHERE news_code='" & RepSql(sKeyPrimary) & "' "
                '    End If
            End If
            ''存檔正確導回default頁面
            'Response.Redirect("Default.aspx?subwhere=" & Replace(txtSubWhere.Text, "%", "$") & "&order=" & txtOrder.Text)
            Show_Message("儲存成功!")
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
            If Chk_RelData(s_prgcode, "", txt_news_code.Text) = False Then
                Chk_data = False
                Show_Message("有重覆參數名稱，請確認")
            End If

        End If
        If rbl_news_type.SelectedValue = "" Then
            Chk_data = False
            Show_Message("請選擇新聞類型")
        End If
        If txt_news_title.Text = "" Or txt_news_contect.Text = "" Then
            Chk_data = False
            Show_Message("必填欄位請確認輸入")
        End If
        '檢查資料正確性       
        'If Chk_DateForm(txt_pet_date.Text) = False Or Chk_DateForm(txt_work_date_s.Text) = False Or Chk_DateForm(txt_work_date_e.Text) = False Then
        '    Chk_data = False
        '    Show_Message("日期格式不正確，請確認")
        'End If

    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '接收default頁面變數
        If Not IsPostBack Then
            SetClientCheck()
            Set_RBL_Option("news_type", "", rbl_news_type, "", "", "")

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

            txt_ins_date.Attributes.Add("ReadOnly", "ReadOnly")
            txt_usr_code.Attributes.Add("ReadOnly", "ReadOnly")

            Select Case Pgm_type
                Case "ADD"
                    txt_ins_date.Text = Format(Now, "yyyy/MM/dd")
                    txt_usr_code.Text = Session("usr_code")
                Case "MDY", "COPY"
                    Dim Dt_tmp As DataTable = Get_DataSet(s_prgcode, "Detail", txtPrimary.Text).Tables(0)
                    If Dt_tmp.Rows.Count > 0 Then
                        txt_news_code.Text = Trim(Dt_tmp.Rows(0).Item("news_code") & "")
                        txt_news_title.Text = Trim(Dt_tmp.Rows(0).Item("news_title") & "")
                        rbl_news_type.SelectedValue = Trim(Dt_tmp.Rows(0).Item("news_type") & "")
                        txt_news_contect.Text = Trim(Dt_tmp.Rows(0).Item("news_contect") & "")
                        txt_usr_code.Text = Trim(Dt_tmp.Rows(0).Item("usr_code") & "")
                        txt_ins_date.Text = Trim(Dt_tmp.Rows(0).Item("ins_date") & "")
                        lab_news_pic1.Text = Set_Files("File_LSI016A", Trim(Dt_tmp.Rows(0).Item("news_pic1") & ""))
                        txt_news_pic1.Text = Trim(Dt_tmp.Rows(0).Item("news_pic1") & "")
                    End If
                    If Pgm_type = "COPY" Then
                        txt_news_code.Text = ""
                        txt_news_code.Enabled = False
                    Else
                        txt_news_code.Enabled = False
                        Bt_preview.Visible = True
                        Bt_preview.Attributes("onclick") = "subwindow=window.open('preview.aspx?code=" & txt_news_code.Text & "','preview','height=768,width=1024,top=200,left=400,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no','');subwindow.focus();return false;"
                    End If
            End Select
        End If

        l_ErrMsg.Text = ""


        'txt_news_contect.Attributes.Add("ONSELECT", "storeCaret(this);")
        'txt_news_contect.Attributes.Add("ONCLICK", "storeCaret(this);")
        'txt_news_contect.Attributes.Add("ONKEYUP", "storeCaret(this);")

        'toolBtn_B.Attributes.Add("onclick", "insertAtCaret(this.form.txt_news_contect, '<b></b>');return false;")
        'toolBtn_I.Attributes.Add("onclick", "insertAtCaret(this.form.txt_news_contect, '<i></i>');return false;")
        'toolBtn_U.Attributes.Add("onclick", "insertAtCaret(this.form.txt_news_contect, '<u></u>');return false;")
        'toolBtn_CENTER.Attributes.Add("onclick", "insertAtCaret(this.form.txt_news_contect, '<center></center>');return false;")
        'toolBtn_FSET.Attributes.Add("onclick", "insertAtCaret(this.form.txt_news_contect, '<font color=red></font>');return false;")

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
