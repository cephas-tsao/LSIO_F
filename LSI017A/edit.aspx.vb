Imports System.Data
Imports System.Web
Imports System.Text.RegularExpressions


Partial Class basic_LSI017A_edit
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI017A"

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '自動編號
        If txtPgmType.Text <> "MDY" Then txt_app_code.Text = "AP" & Format(Now, "yyMMdd") & Right("0000" & Get_AutoIntMax("APP_VERSION", "RIGHT(app_code, 10)", " WHERE app_code LIKE 'AP" & Format(Now, "yyMMdd") & "%'") + 1, 4)

        '檢查資料正確性
        If Chk_data() = False Then Exit Sub

        '檢查上傳圖檔
        If Chk_UpFile() = False Then Exit Sub

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

        If SaveEditData(s_prgcode, Pgm_Type, txt_app_code.Text, txt_verison.Text, txt_version_name.Text, ddl_app_type.SelectedValue,
                        txt_cmemo.Text, txt_update_date.Text, Session("usr_code"), Format(Now, "yyyy/MM/dd")) Then
            '儲存使用紀錄
            ws.InsUsrRec(PrgMemo, Pgm_Type, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)

            If Pgm_Type <> "MDY" Then
                '取得最新條件
                If txtSubWhere.Text <> "" Then
                    txtSubWhere.Text &= " OR app_code='" & RepSql(sKeyPrimary) & "' "
                Else
                    txtSubWhere.Text = " WHERE app_code='" & RepSql(sKeyPrimary) & "' "
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
            If Chk_RelData(s_prgcode, "", txt_app_code.Text) = False Then
                Chk_data = False
                Show_Message("有重覆參數名稱，請確認")
            End If

            If FileUpload1.HasFile = False Then
                Chk_data = False
                Show_Message("未選擇上傳檔案，請確認")
            End If
        End If
        '檢查資料正確性       
        If Chk_DateForm(txt_update_date.Text) = False Then
            Chk_data = False
            Show_Message("日期格式不正確，請確認")
        End If
        If ddl_app_type.SelectedValue = "" Then
            Chk_data = False
            Show_Message("未選擇APP種類，請確認")
        End If

        '檢查必填欄位是否空白
        If txt_verison.Text = "" Then
            Chk_data = False
            Show_Message("必填欄位不得為空白，請確認")
        End If
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '接收default頁面變數
        If Not IsPostBack Then
            SetClientCheck()
            Set_DDL_Option("app_type", "", ddl_app_type, "", "---請選擇---", "")

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
            txt_update_date.Attributes.Add("ReadOnly", "ReadOnly")

            Select Case Pgm_type
                Case "ADD"
                    txt_update_date.Text = Format(Now, "yyyy/MM/dd")
                Case "MDY", "COPY"
                    Dim Dt_tmp As DataTable = Get_DataSet(s_prgcode, "Detail", txtPrimary.Text).Tables(0)
                    If Dt_tmp.Rows.Count > 0 Then
                        txt_app_code.Text = Trim(Dt_tmp.Rows(0).Item("app_code") & "")
                        txt_verison.Text = Trim(Dt_tmp.Rows(0).Item("version") & "")
                        txt_cmemo.Text = Trim(Dt_tmp.Rows(0).Item("cmemo") & "")
                        txt_update_date.Text = Trim(Dt_tmp.Rows(0).Item("update_date") & "")
                        lab_app_path.Text = Set_Files("File_LSI017A", Trim(Dt_tmp.Rows(0).Item("app_path") & ""))
                        txt_version_name.Text = Trim(Dt_tmp.Rows(0).Item("app_path") & "")
                        ddl_app_type.SelectedValue = Trim(Dt_tmp.Rows(0).Item("app_type") & "")
                    End If

                    If Pgm_type = "COPY" Then
                        txt_app_code.Text = ""
                        txt_app_code.Enabled = False
                    Else
                        txt_app_code.Enabled = False
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
    Protected Function Chk_UpFile() As Boolean
        If FileUpload1.HasFile = False Then Return True

        Dim PostedFile As HttpPostedFile = FileUpload1.PostedFile
        Dim sFileName As String = Trim(PostedFile.FileName)
        Dim sExt As String = IO.Path.GetExtension(sFileName).ToLower
        Dim sUpFolder As String = "File_LSI017A"
        'Dim sSaveName = GetSaveName() & sExt

        '檢查附檔名
        If Chk_Ext(sExt) = False Then
            Show_Message("上傳檔案 " & sFileName & " 附檔名非允許格式,請重新選擇檔案")
            Return False
        End If

        If UploadFile(FileUpload1, sUpFolder, GetSaveName(), s_prgcode) = False Then
            Show_Message("檔案上傳失敗，請確認")
            Return False
        End If

        txt_version_name.Text = GetSaveName()
        lab_app_path.Text = Set_Files("File_LSI017A", GetSaveName())
        Return True
    End Function
    Protected Function Chk_Ext(ByVal sExt As String) As Boolean
        If sExt <> ".apk" Then Return False
        Return True
    End Function
    Protected Function GetSaveName() As String
        '上傳檔案固定檔名
        Select Case ddl_app_type.SelectedValue
            Case "1"
                Return "LSIO_gps.apk"
            Case "2"
                Return "WorkTime.apk"
            Case "3"
                Return "LSIO.apk"
            Case Else
                Return ""
        End Select
    End Function
End Class
