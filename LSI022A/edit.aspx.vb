Imports System.Data

Partial Class basic_LSI002A_edit
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI022A"

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '自動編號
        If txtPgmType.Text <> "MDY" Then txt_cmp_code.Text = "AP" & Format(Now, "yyMMdd") & Right("0000" & Get_AutoIntMax("mail", "RIGHT(cmp_code, 10)", " WHERE cmp_code LIKE 'AP" & Format(Now, "yyMMdd") & "%'") + 1, 4)
        If txtPgmType.Text <> "MDY" Then
            ins_date.Text = Format(Now, "yyyy/MM/dd")
            cmp_date.Text = DateTime.Now.ToString("HH:mm:ss")
            'cmp_pass.Text = Session("ComplainA_VerCode")
        End If
        '檢查資料正確性
        If Chk_data() = False Then Exit Sub

        '設定Key欄位來源
        Dim sKeyPrimary As String = txtPrimary.Text '*修改處*

        '設定記錄資料
        Dim sSql As String = Get_SqlStr(s_prgcode, "Detail", sKeyPrimary)
        Dim sFieldName As String = ws.Get_FieldName(s_prgcode)
        Dim sOldData As String = ws.Get_OldData(sSql, s_prgcode)
        Dim sNewData As String = Get_NewData(s_prgcode)

        Dim showtype As String = ""
        

        '資料存檔
        Dim Pgm_Type As String = txtPgmType.Text
        Dim PrgMemo As String = RepSql(Label1.Text & ":" & sKeyPrimary)

        If SaveEditData(s_prgcode, Pgm_Type, txt_cmp_code.Text, Format(Now, "yyyy/MM/dd"), cmp_date.Text, txt_cmp_memo.Text, txt_cmp_name.Text, txt_cmp_email.Text, txt_cmp_tel1.Text, "", "", "", "", "", _
                        ) Then
            '儲存使用紀錄
            ws.InsUsrRec(PrgMemo, Pgm_Type, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)

            If Pgm_Type <> "MDY" Then
                '取得最新條件
                If txtSubWhere.Text <> "" Then
                    txtSubWhere.Text &= " OR cmp_code='" & RepSql(sKeyPrimary) & "' "
                Else
                    txtSubWhere.Text = " WHERE cmp_code='" & RepSql(sKeyPrimary) & "' "
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
        If txtPgmType.Text = "ADD" Then
            If Chk_RelData(s_prgcode, "", txt_cmp_code.Text) = False Then
                Chk_data = False
                Show_Message("有重覆參數名稱，請確認")
            End If
        End If
        '檢查資料正確性
        
        If txt_cmp_memo.Text = "" Then
            Chk_data = False
            Show_Message("信件內容不得為空白，請確認")
        End If

        If txt_cmp_name.Text = "" Then
            Chk_data = False
            Show_Message("申訴人姓名不得為空白，請確認")
        End If
        If Chk_Email(txt_cmp_email.Text) = False Then
            Chk_data = False
            Show_Message("電子郵件格式不正確，請確認")
        End If
        If Chk_Tel(txt_cmp_tel1.Text) = False Then
            Chk_data = False
            Show_Message("聯絡電話格式不正確，請確認")
        End If



        '檢查必填欄位是否空白

    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '接收default頁面變數
        If Not IsPostBack Then

            SetClientCheck()
            '設定下拉式選單
            'Set_DDL_Option("ZipCode", "110", ddl_att_area, "", "---請選擇---", "")
            'Set_DDL_Option("areacode", "are01", ddl_att_area, "", "---請選擇---", "")
            'Set_DDL_Option(ddl_att_area.SelectedValue, "", ddl_att_road, "", "---請選擇---", "")

            'Set_DDL_Option("ZipCode", "110", ddl_att_area1, "", "---請選擇---", "")
            'Set_DDL_Option("areacode", "are01", ddl_att_area1, "", "---請選擇---", "")
            'Set_DDL_Option(ddl_att_area1.SelectedValue, "", ddl_att_road1, "", "---請選擇---", "")

            'Set_DDL_Option("CmpType", "", ddl_cmp_type, "", "", "")

            'Set_CKL_Option("com_Contra", "", com_Contract, "", "", "")

            'Set_CKL_Option("com_paymon", "", com_paymon, "", "", "")

            'Set_CKL_Option("com_relax", "", com_relax, "", "", "")

            'Set_CKL_Option("com_child", "", com_child, "", "", "")

            'Set_CKL_Option("com_Welfar", "", com_Welfar, "", "", "")

            If txtPgmType.Text <> "MDY" Then
                Set_Ver_Code()
                ins_date.Text = Format(Now, "yyyy/MM/dd")
                cmp_date.Text = DateTime.Now.ToString("HH:mm:ss")
                'cmp_pass.Text = Session("ComplainA_VerCode")
            End If
            'Set_DDL_Option("YorN", "", ddl_is_agree, "", "---請選擇---", "")
            '標頭 
            Label1.Text = CType(PreviousPage.FindControl("Label1"), Label).Text
            '編輯模式
            txtPgmType.Text = CType(PreviousPage.FindControl("txtPgmType"), TextBox).Text
            ddl_is_agree.SelectedValue = "Y"
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
                        ins_date.Text = Trim(Dt_tmp.Rows(0).Item("ins_date") & "")
                        cmp_date.Text = Trim(Dt_tmp.Rows(0).Item("cmp_date") & "")
                        txt_cmp_code.Text = Trim(Dt_tmp.Rows(0).Item("cmp_code") & "")
                        txt_cmp_name.Text = Trim(Dt_tmp.Rows(0).Item("cmp_name") & "")
                        txt_cmp_tel1.Text = Trim(Dt_tmp.Rows(0).Item("cmp_tel1") & "")
                        txt_cmp_email.Text = Trim(Dt_tmp.Rows(0).Item("cmp_email") & "")
                        txt_cmp_memo.Text = Trim(Dt_tmp.Rows(0).Item("cmp_memo") & "")
                        lab_cmp_file1.Text = Set_Files("File_LSI002A", Trim(Dt_tmp.Rows(0).Item("cmp_file1") & ""))
                        lab_cmp_file2.Text = Set_Files("File_LSI002A", Trim(Dt_tmp.Rows(0).Item("cmp_file2") & ""))
                        lab_cmp_file3.Text = Set_Files("File_LSI002A", Trim(Dt_tmp.Rows(0).Item("cmp_file3") & ""))
                        lab_cmp_file4.Text = Set_Files("File_LSI002A", Trim(Dt_tmp.Rows(0).Item("cmp_file4") & ""))
                        lab_cmp_file5.Text = Set_Files("File_LSI002A", Trim(Dt_tmp.Rows(0).Item("cmp_file5") & ""))
                    End If

                    If Pgm_type = "COPY" Then
                        txt_cmp_code.Text = ""
                        txt_cmp_code.Enabled = False
                    Else
                        txt_cmp_code.Enabled = False
                        '具名申訴則顯示補發按鈕
                        If Chk_Email(txt_cmp_email.Text) Then
                            txtFileName1.Text = Trim(Dt_tmp.Rows(0).Item("cmp_file1") & "")
                            txtFileName2.Text = Trim(Dt_tmp.Rows(0).Item("cmp_file2") & "")
                            txtFileName3.Text = Trim(Dt_tmp.Rows(0).Item("cmp_file3") & "")
                            txtFileName4.Text = Trim(Dt_tmp.Rows(0).Item("cmp_file4") & "")
                            txtFileName5.Text = Trim(Dt_tmp.Rows(0).Item("cmp_file5") & "")
                            Bt_send.Visible = True
                        End If
                        
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

    Protected Sub Bt_send_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Bt_send.Click
        '發送Mail通知
        Dim dtMail As New DataTable
        Dim column1 As New DataColumn
        Dim column2 As New DataColumn
        column1.DataType = System.Type.GetType("System.String")
        column2.DataType = System.Type.GetType("System.String")
        dtMail.Columns.Add(column1)
        dtMail.Columns.Add(column2)
        '加入申訴人Mail
        Dim rowMail As DataRow = dtMail.NewRow
        rowMail(0) = txt_cmp_email.Text
        rowMail(1) = txt_cmp_name.Text
        dtMail.Rows.Add(rowMail)
        '加入承辦人信箱
        Dim MailList As String() = Split(Get_SysPara("mail_receive"), "|")
        For i As Integer = 0 To UBound(MailList)
            If String.IsNullOrEmpty(MailList(i)) = False Then
                Dim rowTmp As DataRow = dtMail.NewRow
                rowTmp(0) = MailList(i)
                rowTmp(1) = "申訴承辦人"
                dtMail.Rows.Add(rowTmp)
            End If
        Next
        'Dim sIs_agree As String = "不同意本處繼續使用資料"
        'If ddl_is_agree.SelectedValue = "Y" Then sIs_agree = "同意本處繼續使用資料"
        '郵件內容

        Dim showtype1 As String = ""
        

        Dim bFirst As Boolean = True
        bFirst = True

        

        Dim sBody As String = "" & _
            "<html xmlns='http://www.w3.org/1999/xhtml'>" & _
            "<head><meta http-equiv='Content-Type' content='text/html; charset=utf-8' /></head><body>" & _
            "<table cellpadding=5 cellspacing=0 bordercolor=#7EBAD1 border=1 bgcolor=#ffffff style='border-collapse:collapse;'" & _
            "        width='720' summary='排版表格' align=center class='c12'>" & _
            "<tr bgcolor='#7EBAD1' align='center'><td colspan='4'><p><strong>具名申訴</strong></p></td></tr>" & _
                "<tr>" & _
                "  <td width='25%'>受理日期：</td>" & _
                "  <td colspan='3'>" & ins_date.Text & "</td>" & _
                "</tr>" & _
                "<tr>" & _
                "  <td>受理時間：</td>" & _
                "  <td colspan='3'>" & cmp_date.Text & "</td>" & _
                "<tr>" & _
                "  <td>案件編號：</td>" & _
                "  <td colspan='3'>" & txt_cmp_code.Text & "</td>" & _
                "</tr>" & _
                "<tr>" & _
                "  <td>信件內容：</td>" & _
                "  <td colspan='3'>" & txt_cmp_memo.Text & "</td>" & _
                "</tr>" & _
                "<tr bgcolor='#7EBAD1' align='left'><td colspan='4'><p><strong>當事人基本資料</strong></p></td></tr>" & _
                "<tr>" & _
                "  <td>姓　　　　　名：</td>" & _
                "  <td colspan='3'>" & txt_cmp_name.Text & "</td>" & _
                "</tr>" & _
                "<tr>" & _
                "<td>電子信箱E-mail：</td>" & _
                "<td colspan='3'>" & txt_cmp_email.Text & "</td>" & _
                "</tr>" & _
                "<tr>" & _
                "  <td>聯　絡　電　話：</td>" & _
                "  <td colspan='3'>" & txt_cmp_tel1.Text & "</td>" & _
                "</tr>" & _
                "<tr>" & _
                "<td colspan='4' align='center'><font color='#666666' style=' size:-1'>謝謝您的來信！本局將以最快速度回覆！！</font></td>" & _
                "</tr>" & _
                "</table></body>"

        Dim sUpFolder As String = Get_SysPara("AP_File")
        Dim sFilePath1 As String = sUpFolder & txtFileName1.Text
        Dim sFilePath2 As String = sUpFolder & txtFileName2.Text
        Dim sFilePath3 As String = sUpFolder & txtFileName3.Text
        Dim sFilePath4 As String = sUpFolder & txtFileName4.Text
        Dim sFilePath5 As String = sUpFolder & txtFileName5.Text
        If SendMail("臺北市勞動檢查處-意見信箱-送出成功", dtMail, sBody, s_prgcode, sFilePath1, sFilePath2, sFilePath3, sFilePath4, sFilePath5) Then
            Show_Message("送出成功")
        Else
            Show_Message("送出失敗，請通知系統管理員")
            Exit Sub
        End If
    End Sub

    Protected Function Set_Ver_Code() As Boolean
        '產生亂數英數字串
        Dim array As String = "#123456789ABCDEFGHIJKLMN$PQRSTUVWXYZ"
        Dim generator As New Random
        txtVerCode.Text = ""
        For i = 1 To 6
            txtVerCode.Text &= array(Int(generator.Next(0, 36)))
            Session("ComplainA_VerCode") = txtVerCode.Text
        Next
    End Function
End Class
