Imports System.Data
Imports System.IO
Partial Class basic_LSI010A_edit
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI010A"

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '自動編號
        If txtPgmType.Text <> "MDY" Then txt_start_code.Text = "ST" & Format(Now, "yyMMdd") & Right("0000" & Get_AutoIntMax("Work_start", "RIGHT(start_code, 10)", " WHERE start_code LIKE 'ST" & Format(Now, "yyMMdd") & "%'") + 1, 4)

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

        If SaveEditData(s_prgcode, Pgm_Type, txt_start_code.Text, txt_stop_number.Text, "", "", _
                        "", "", txt_eng_name.Text, "", "", _
                        "", txt_ins_date.Text, "", "", txt_att_name.Text, _
                        txt_att_tel1.Text, "", "", "", "", "", _
                        Session("usr_code")) Then
            '儲存使用紀錄
            ws.InsUsrRec(PrgMemo, Pgm_Type, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)

            If Pgm_Type <> "MDY" Then
                '取得最新條件
                If txtSubWhere.Text <> "" Then
                    txtSubWhere.Text &= " OR start_code='" & RepSql(sKeyPrimary) & "' "
                Else
                    txtSubWhere.Text = " WHERE start_code='" & RepSql(sKeyPrimary) & "' "
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
            If Chk_RelData(s_prgcode, "start_code", txt_start_code.Text) = False Then
                Chk_data = False
                Show_Message("有重覆參數名稱，請確認")
            End If

            If Chk_RelData(s_prgcode, "stop_number", txt_stop_number.Text) = False Then
                Chk_data = False
                Show_Message("有重覆停工文號，請確認")
            End If
        End If

        '檢查停工文號是否正確
        If Chk_StopNumber(txt_stop_number.Text) = False Then
            Chk_data = False
            Show_Message("停工文號格式不正確，請確認")
        End If

        
        '檢查必填欄位是否空白
        If txt_stop_number.Text = "" Or txt_eng_name.Text = ""  Or txt_att_tel1.Text = "" Then
            Chk_data = False
            Show_Message("必填欄位不得為空白，請確認")
        End If
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '接收default頁面變數
        If Not IsPostBack Then
            SetClientCheck()

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
                    txt_ins_date.Text = Format(Now, "yyyy/MM/dd")
                Case "MDY", "COPY"
                    Dim Dt_tmp As DataTable = Get_DataSet(s_prgcode, "Detail", txtPrimary.Text).Tables(0)
                    If Dt_tmp.Rows.Count > 0 Then
                        txt_start_code.Text = Trim(Dt_tmp.Rows(0).Item("start_code") & "")
                        txt_stop_number.Text = Trim(Dt_tmp.Rows(0).Item("stop_number") & "")
                        txt_eng_name.Text = Trim(Dt_tmp.Rows(0).Item("Eng_name") & "")
                        txt_att_name.Text = Trim(Dt_tmp.Rows(0).Item("Att_name") & "")
                        txt_att_tel1.Text = Trim(Dt_tmp.Rows(0).Item("Att_tel1") & "")
                        txt_ins_date.Text = Trim(Dt_tmp.Rows(0).Item("ins_date") & "")
                       ' lab_Start_path1.Text = Set_Files("File_LSI010A", Trim(Dt_tmp.Rows(0).Item("start_path1") & "")) 
                       ' lab_start_path2.Text = Set_Files("File_LSI010A", Trim(Dt_tmp.Rows(0).Item("start_path2") & ""))
                       
                        Dim imgRoot = "/FileUpload/File_LSI010A/" & Trim(Dt_tmp.Rows(0).Item("start_code"))
						
						 If Not File.Exists(imgRoot & "_1.pdf ") Then
						  ' Objec標籤
           ' Dim strCLSID As String =  "clsid:CA8A9780-280D-11CF-A24D-444553540000" 'pdf clsid
            Dim fileName As String = imgRoot & "_1.pdf"  'pdf檔案位置  
			'Dim pdfOcx As String =  "<embed id=""pdf1"" height=""1024""  width=""700""   "  
            'pdfOcx += " data-render-src="""   & fileName & """   "   
			'pdfOcx += " type=""application/pdf ""   >"   
			
			Dim pdfOcx As String =  "<OBJECT id=""pdf1"" height=""1024""  width=""700""  "  
				pdfOcx += " data="""   &imgRoot & "_1.pdf" & """   "   
				pdfOcx += " type=""application/pdf ""   >"   	
				pdfOcx += " </OBJECT>"
              
           If  File.Exists(imgRoot & "_1.pdf") Then
            Me.Literal1.Text = pdfOcx
			else 
			 Me.Literal1.Text="沒有資料。"
			end if 
			'img_att1.visible=false

						end if 
					
						
                        If Not File.Exists(imgRoot & "_1.jpg") Then
                            img_att1.ImageUrl = imgRoot & "_1.jpg"

                            img_att1.Width = 700

                        Else
                            img_att1.ImageUrl =  imgRoot  & "nofile.jpg"
                            img_att1.Width = 300
                        End If


                        If Not File.Exists(imgRoot & "_2.jpg") Then
                            img_att2.ImageUrl = imgRoot & "_2.jpg"
                            img_att2.Width = 700
                        Else
                            img_att2.ImageUrl = "/FileUpload/File_LSI010A/" & "nofile.jpg"
                            img_att1.Width = 300
                        End If


                    End If

                    If Pgm_type = "COPY" Then
                        txt_start_code.Text = ""
                        txt_start_code.Enabled = False
                    Else
                        txt_start_code.Enabled = False
                        'MAIL欄位隱藏()
                        'If Chk_Email(txt_att_email.Text) Then
                        '    txtFileName1.Text = Trim(Dt_tmp.Rows(0).Item("start_path1") & "")
                        '    txtFileName2.Text = Trim(Dt_tmp.Rows(0).Item("start_path2") & "")
                        '    Bt_send.Visible = True
                        'End If
                    End If
            End Select
        End If
        ib_code1.Attributes("onclick") = ws.Set_AddForm("../../public/addform.aspx", "SPN-1", "1", "txt_stop_number", "")
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

    'Protected Sub Bt_send_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Bt_send.Click
    '    '發送Mail通知
    '    Dim dtMail As New DataTable
    '    Dim column1 As New DataColumn
    '    Dim column2 As New DataColumn
    '    column1.DataType = System.Type.GetType("System.String")
    '    column2.DataType = System.Type.GetType("System.String")
    '    dtMail.Columns.Add(column1)
    '    dtMail.Columns.Add(column2)
    '    '加入申訴人Mail
    '    Dim rowMail As DataRow = dtMail.NewRow
    '    rowMail(0) = txt_att_email.Text
    '    rowMail(1) = txt_att_name.Text
    '    dtMail.Rows.Add(rowMail)
    '    '加入承辦人信箱
    '    Dim MailList As String() = Split(Get_SysPara("ST_receive"), "|")
    '    For i As Integer = 0 To UBound(MailList)
    '        If String.IsNullOrEmpty(MailList(i)) = False Then
    '            Dim rowTmp As DataRow = dtMail.NewRow
    '            rowTmp(0) = MailList(i)
    '            rowTmp(1) = "線上復工申請承辦人"
    '            dtMail.Rows.Add(rowTmp)
    '        End If
    '    Next
    '    '郵件內容
    '    Dim sBody As String = "" & _
    '        "<table cellpadding=5 cellspacing=0 bordercolor=#7EBAD1 border=1 bgcolor=#ffffff style='border-collapse:collapse;'" & _
    '        "width='720' summary='排版表格' align=center class='c12'>" & _
    '        "<tr bgcolor='#7EBAD1' align='center'><td colspan='4'><p><strong>線上復工申請資料表</strong></p></td></tr>" & _
    '        "<tr>" & _
    '        "  <td>申　請　單　位：</td>" & _
    '        "  <td colspan='3'>" & txt_com_name.Text & "</td>" & _
    '        "</tr>" & _
    '        "<tr>" & _
    '        "  <td>負　　責　　人：</td>" & _
    '        "  <td colspan='3'>" & txt_com_per.Text & "</td>" & _
    '        "</tr>" & _
    '        "<tr>" & _
    '        "  <td>事業單位發文日期：</td>" & _
    '        "  <td colspan='3'>西元 " & Mid(txt_start_date.Text, 1, 4) & " 年 " & _
    '                                    Mid(txt_start_date.Text, 6, 2) & " 月 " & _
    '                                    Mid(txt_start_date.Text, 9, 2) & " 日 </td>" & _
    '        "</tr>" & _
    '        "<tr>" & _
    '        "  <td>事業單位發文文號：</td>" & _
    '        "  <td colspan='3'>" & txt_start_number.Text & "</td>" & _
    '        "</tr>" & _
    '        "<tr>" & _
    '        "  <td>工　程　名　稱：</td>" & _
    '        "  <td colspan='3'>" & txt_eng_name.Text & "</td>" & _
    '        "</tr>" & _
    '        "<tr>" & _
    '        "  <td>建　　　　　號：</td>" & _
    '        "  <td colspan='3'>" & txt_eng_code.Text & "</td>" & _
    '        "</tr>" & _
    '        "<tr>" & _
    '        "  <td>座　落　位　置：</td>" & _
    '        "  <td colspan='3'>" & txt_eng_gps.Text & "</td>" & _
    '        "</tr>" & _
    '        "<tr>" & _
    '        "  <td>檢　查　日　期：</td>" & _
    '        "  <td colspan='3'>西元 " & Mid(txt_chk_date.Text, 1, 4) & " 年 " & _
    '                                    Mid(txt_chk_date.Text, 6, 2) & " 月 " & _
    '                                    Mid(txt_chk_date.Text, 9, 2) & " 日 </td>" & _
    '        "</tr>" & _
    '        "<tr>" & _
    '        "  <td>申　請　日　期：</td>" & _
    '        "  <td colspan='3'>西元 " & Mid(txt_ins_date.Text, 1, 4) & " 年 " & _
    '                                    Mid(txt_ins_date.Text, 6, 2) & " 月 " & _
    '                                    Mid(txt_ins_date.Text, 9, 2) & " 日 </td>" & _
    '        "</tr>" & _
    '        "<tr>" & _
    '        "  <td>停　工　文　號：</td>" & _
    '        "  <td colspan='3'>" & txt_stop_number.Text & "</td>" & _
    '        "</tr>" & _
    '        "<tr>" & _
    '        "  <td>停　工　時　間：</td>" & _
    '        "  <td colspan='3'>自西元 " & Mid(txt_stop_date.Text, 1, 4) & " 年 " & _
    '                                    Mid(txt_stop_date.Text, 6, 2) & " 月 " & _
    '                                    Mid(txt_stop_date.Text, 9, 2) & " 日 " & _
    '                                    Mid(txt_stop_time.Text, 1, 2) & " 時 " & _
    '                                    Mid(txt_stop_time.Text, 4, 2) & " 分起<br />" & _
    '                          "至西元 " & Mid(txt_stop_date_e.Text, 1, 4) & " 年 " & _
    '                                    Mid(txt_stop_date_e.Text, 6, 2) & " 月 " & _
    '                                    Mid(txt_stop_date_e.Text, 9, 2) & " 日 " & _
    '                                    Mid(txt_stop_time_e.Text, 1, 2) & " 時 " & _
    '                                    Mid(txt_stop_time_e.Text, 4, 2) & " 分止<br /></td>" & _
    '        "</tr>" & _
    '        "<tr>" & _
    '        "  <td width='24%'>聯　　絡　　人：</td>" & _
    '        "  <td width='26%'>" & txt_att_name.Text & "</td>" & _
    '        "  <td width='24%'>聯　絡　電　話：</td>" & _
    '        "  <td width='26%'>" & txt_att_tel1.Text & "</td>" & _
    '        "</tr>" & _
    '        "<tr>" & _
    '        "  <td>聯絡E-mail：</td>" & _
    '        "  <td colspan='3'>" & txt_att_email.Text & "</td>" & _
    '        "</tr>" & _
    '        "</table>"

    '    Dim sUpFolder As String = Get_SysPara("ST_File")
    '    Dim sFilePath1 As String = sUpFolder & txtFileName1.Text
    '    Dim sFilePath2 As String = sUpFolder & txtFileName2.Text
    '    If SendMail("臺北市勞動檢查處-申辦服務-線上復工申請-送出成功", dtMail, sBody, s_prgcode, sFilePath1, sFilePath2) Then
    '        Show_Message("送出成功")
    '    Else
    '        Show_Message("送出失敗，請通知系統管理員")
    '        Exit Sub
    '    End If
    'End Sub
End Class
