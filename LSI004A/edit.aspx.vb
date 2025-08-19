Imports System.Data
Imports System.IO
Imports System.Web
Imports System.Text.RegularExpressions

Partial Class basic_LSI004A_edit
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI004A"

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '自動編號
        If txtPgmType.Text <> "MDY" Then txt_paper_code.Text = "PP" & Format(Now, "yyMMdd") & Right("0000" & Get_AutoIntMax("e_paper", "RIGHT(paper_code, 10)", " WHERE paper_code LIKE 'PP" & Format(Now, "yyMMdd") & "%'") + 1, 4)

        '檢查資料正確性
        If Chk_data() = False Then Exit Sub

        '檢查上傳圖檔
        If Chk_UpFile() = False Then Exit Sub

        '設定Key欄位來源
        Dim sKeyPrimary As String = txt_paper_code.Text '*修改處*

        '設定記錄資料
        Dim sSql As String = Get_SqlStr(s_prgcode, "Detail", sKeyPrimary)
        Dim sFieldName As String = ws.Get_FieldName(s_prgcode)
        Dim sOldData As String = ws.Get_OldData(sSql, s_prgcode)
        Dim sNewData As String = Get_NewData(s_prgcode)

        '資料存檔
        Dim Pgm_Type As String = txtPgmType.Text
        Dim PrgMemo As String = RepSql(Label1.Text & ":" & sKeyPrimary)

        If SaveEditData(s_prgcode, Pgm_Type, txt_paper_code.Text, txt_paper_name.Text, txt_paper_num.Text, txt_paper_logo.Text, _
                        txt_paper_bottom1.Text, txt_paper_bottom2.Text, txt_usr_code.Text, txt_ins_date.Text, ddl_paper_type.SelectedValue, _
                         "", ddl_paper_css.SelectedValue) Then
            '***更新HTML資料************
            detail_edit.Set_Paper_Html()
            '***************************
            Show_Message("存檔成功")
            '儲存使用紀錄
            ws.InsUsrRec(PrgMemo, Pgm_Type, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)

            If Pgm_Type <> "MDY" Then
                txtPgmType.Text = "MDY"
                detail_edit.setKeyName(sKeyPrimary)
                Panel1.Visible = True

                '取得最新條件
                If txtSubWhere.Text <> "" Then
                    txtSubWhere.Text &= " OR paper_code='" & RepSql(sKeyPrimary) & "' "
                Else
                    txtSubWhere.Text = " WHERE paper_code='" & RepSql(sKeyPrimary) & "' "
                End If
                '設定預覽按鈕
                Bt_preview.Attributes("onclick") = ("window.open('preview.aspx?code=" & txt_paper_code.Text & "','','height=768,width=1024,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no','');return false;")
                Bt_preview.Visible = True
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
            If Chk_RelData(s_prgcode, "", txt_paper_code.Text) = False Then
                Chk_data = False
                Show_Message("有重覆資料，請確認")
            End If
            If fu_paper_logo.HasFile = False Then
                Chk_data = False
                Show_Message("未選擇上傳圖片，請確認")
            End If
        End If

        '檢查必填欄位是否空白
        If txt_paper_name.Text = "" Or txt_paper_num.Text = "" Then
            Chk_data = False
            Show_Message("必填欄位不得為空白，請確認")
        End If

        ''檢查日期格式是否正確
        'If Chk_DateForm(txt_tn_date.Text) = False Then
        '    Chk_data = False
        '    Show_Message("日期格式錯誤，請確認")
        'End If
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

            '非新增模式則預設取得資料
            Dim Pgm_type As String = txtPgmType.Text
            txtPrimary.Text = CType(PreviousPage.FindControl("txtPrimary"), TextBox).Text
            txt_paper_code.Attributes.Add("ReadOnly", "ReadOnly")
            txt_ins_date.Attributes.Add("ReadOnly", "ReadOnly")
            Set_DDL_Option("paper_type", "", ddl_paper_type, "", "", "")
            Set_DDL_Option("paper_css", "", ddl_paper_css, "", "", "")

            Select Case Pgm_type
                Case "ADD"
                    txt_ins_date.Text = Format(Now, "yyyy/MM/dd")
                    txt_usr_code.Text = Session("usr_code")
                Case "MDY", "COPY"
                    Dim Dt_tmp As DataTable = Get_DataSet(s_prgcode, "Detail", txtPrimary.Text).Tables(0)
                    If Dt_tmp.Rows.Count > 0 Then
                        If Pgm_type = "COPY" Then
                            txt_paper_code.Text = ""
                        Else
                            txt_paper_code.Text = txtPrimary.Text
                            txt_paper_code.Enabled = False
                            detail_edit.setKeyName(txtPrimary.Text)
                            Panel1.Visible = True
                            '設定預覽按鈕
                            Bt_preview.Attributes("onclick") = ("window.open('preview.aspx?code=" & txt_paper_code.Text & "','','height=768,width=1024,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no','');return false;")
                            Bt_preview.Visible = True
                        End If
                        txt_paper_name.Text = Trim(Dt_tmp.Rows(0).Item("paper_name") & "")
                        txt_paper_num.Text = Trim(Dt_tmp.Rows(0).Item("paper_num") & "")
                        txt_paper_logo.Text = Trim(Dt_tmp.Rows(0).Item("paper_logo") & "")
                        txt_paper_bottom1.Text = Trim(Dt_tmp.Rows(0).Item("paper_bottom1") & "")
                        txt_paper_bottom2.Text = Trim(Dt_tmp.Rows(0).Item("paper_bottom2") & "")
                        txt_usr_code.Text = Trim(Dt_tmp.Rows(0).Item("usr_code") & "")
                        txt_ins_date.Text = Trim(Dt_tmp.Rows(0).Item("ins_date") & "")
                        l_paper_logo.Text = Set_Files("File_LSI004A", txt_paper_logo.Text)
                        txt_usr_name.Text = ws.QueryData(txt_usr_code.Text, "BDP080", "usr_code", "usr_name")
                        ddl_paper_type.SelectedValue = Trim(Dt_tmp.Rows(0).Item("paper_type") & "")
                        ddl_paper_css.SelectedValue = Trim(Dt_tmp.Rows(0).Item("paper_css") & "")
                    End If
            End Select
        End If
        ib_usr_code.Attributes("onclick") = ws.Set_AddForm("../../public/addform.aspx", "USR-1", "1|2", "txt_usr_code|txt_usr_name", "")
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

    Protected Sub txt_usr_code_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt_usr_code.TextChanged
        txt_usr_name.Text = ws.QueryData(txt_usr_code.Text, "BDP080", "usr_code", "usr_name")
    End Sub

    Protected Function Chk_UpFile() As Boolean
        If Not fu_paper_logo.HasFile Then
            Show_Message("請選擇要上傳的檔案！")
            Return False
        End If

        Dim posted = fu_paper_logo.PostedFile

        ' 1) 只取基名，防止路徑穿越
        Dim originalName As String = Path.GetFileName(Trim(posted.FileName))
        If String.IsNullOrEmpty(originalName) Then
            Show_Message("檔名不可為空。")
            Return False
        End If

        ' 2) 副檔名白名單
        Dim ext As String = Path.GetExtension(originalName).ToLowerInvariant()
        Dim allowed As String() = {".jpg", ".jpeg", ".png", ".gif", ".pdf"} ' ← 依需求調整
        Dim ok As Boolean = False
        For Each e As String In allowed
            If e = ext Then ok = True : Exit For
        Next
        If Not ok Then
            Show_Message("不允許的檔案類型。")
            Return False
        End If

        ' 3) 檔案大小限制（例：10 MB）
        If posted.ContentLength <= 0 OrElse posted.ContentLength > 10 * 1024 * 1024 Then
            Show_Message("檔案大小不合法（上限 10MB）。")
            Return False
        End If

        ' 4) 以自產安全檔名儲存（以主鍵/代號 + 固定字樣 + 副檔名）
        Dim safeCode As String = Regex.Replace(If(txt_paper_code.Text, ""), "[^A-Za-z0-9_-]", "")
        If safeCode = "" Then safeCode = DateTime.Now.ToString("yyyyMMddHHmmss")
        Dim saveName As String = safeCode & "_logo" & ext

        ' 5) 實體/虛擬路徑
        Dim folderVirtual As String = "~/FileUpload/File_LSI004A/"
        Dim folderPhysical As String = Server.MapPath(folderVirtual)
        If Not Directory.Exists(folderPhysical) Then Directory.CreateDirectory(folderPhysical)

        ' 6) 儲存
        Dim fullPath As String = Path.Combine(folderPhysical, saveName)
        posted.SaveAs(fullPath)

        ' 7) 記錄與顯示（文字框只留檔名；畫面上的連結用安全版 Set_Files 產生）
        txt_paper_logo.Text = saveName
        l_paper_logo.Text = Set_Files_safe("File_LSI004A", saveName, originalName)  ' 原檔名只作顯示文字，會被 HtmlEncode

        Return True
        If fu_paper_logo.HasFile = False Then Return True

        Dim PostedFile As HttpPostedFile = fu_paper_logo.PostedFile
        Dim sFileName As String = Trim(PostedFile.FileName)
        Dim sExt As String = IO.Path.GetExtension(sFileName).ToLower
        Dim sUpFolder As String = "File_LSI004A"
        Dim sSaveName = txt_paper_code.Text & "_logo" & sExt

        '檢查附檔名
        If Chk_Ext(sExt) = False Then
            Show_Message("上傳檔案 " & sFileName & " 附檔名非允許格式,請重新選擇檔案")
            Return False
        End If

        If UploadFile(fu_paper_logo, sUpFolder, sSaveName, s_prgcode) = False Then
            Show_Message("檔案上傳失敗，請確認")
            Return False
        End If

        txt_paper_logo.Text = sSaveName
        l_paper_logo.Text = Set_Files("File_LSI004A", txt_paper_logo.Text)
        Return True
    End Function

    Protected Function Chk_Ext(ByVal sExt As String) As Boolean
        If sExt <> ".jpg" And sExt <> ".gif" Then Return False
        Return True
    End Function
End Class

