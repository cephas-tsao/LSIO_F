Imports System.Data

Partial Class basic_LSI021A_edit
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI021A"

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '自動編號
        If txtPgmType.Text <> "MDY" Then txt_roof_cod.Text = "RF" & Format(Now, "yyMMdd") & Right("0000" & Get_AutoIntMax("Roof_work", "RIGHT(roof_cod, 10)", " WHERE roof_cod LIKE 'RF" & Format(Now, "yyMMdd") & "%'") + 1, 4)

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

        If SaveEditData(s_prgcode, Pgm_Type, txt_roof_cod.Text, txt_com_name.Text, ddl_att_area.SelectedValue, ddl_att_road.SelectedValue _
                        , txt_att_add1.Text, txt_att_name1.Text, txt_att_tel1.Text, txt_att_name2.Text, txt_att_tel2.Text, txt_roof_mail.Text _
                        , txt_pet_date.Text, Get_CKL_Click(chk_roof_tool1), txt_roof_sdate.Text, txt_roof_edate.Text, txt_roof_height.Text, txt_floor_height.Text, txt_roof_memo.Text _
                        , Get_CKL_Click(chk_tear_tool1), txt_tearon_sdate.Text, txt_tearon_edate.Text, txt_tearoff_sdate.Text, txt_tearoff_edate.Text, txt_tear_height.Text, txt_tearfloor_height.Text, txt_tear_memo.Text _
                        , Get_CKL_Click(chk_cage_tool1), txt_cage_sdate.Text, txt_cage_edate.Text, txt_cageroof_height.Text, txt_cagefloor_height.Text, txt_cage_memo.Text, selwork.SelectedValue _
                        ) Then


            '儲存使用紀錄
            ws.InsUsrRec(PrgMemo, Pgm_Type, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)

            If Pgm_Type <> "MDY" Then
                '取得最新條件
                If txtSubWhere.Text <> "" Then
                    txtSubWhere.Text &= " OR roof_cod='" & RepSql(sKeyPrimary) & "' "
                Else
                    txtSubWhere.Text = " WHERE roof_cod='" & RepSql(sKeyPrimary) & "' "
                End If
            End If
            '存檔正確導回default頁面
            Response.Redirect("Default.aspx?subwhere=" & Replace(txtSubWhere.Text, "%", "$") & "&order=" &  HttpUtility.HtmlEncode(txtOrder.Text ) )
        Else
            '存檔出錯就秀出錯誤訊息
            Show_Message("儲存失敗!")
        End If
    End Sub

    Protected Function Chk_data() As Boolean
        'Server端資料檢查 自訂函數
        Chk_data = True

        If selwork.SelectedValue = "" Then
            Show_Message("請挑選資料來源")
        End If
        '檢查重覆資料
        If txtPgmType.Text = "ADD" Or txtPgmType.Text = "COPY" Then
            If Chk_RelData(s_prgcode, "", txt_roof_cod.Text) = False Then
                Chk_data = False
                Show_Message("有重覆參數名稱，請確認")
            End If

        End If
        '檢查資料正確性       
        If Chk_DateForm(txt_pet_date.Text) = False Then
            Chk_data = False
            Show_Message("通報日期日期格式不正確，請確認")
        End If

        If chk_roof_tool1.SelectedValue = "" And chk_tear_tool1.SelectedValue = "" And chk_cage_tool1.SelectedValue = "" Then
            Chk_data = False
            Show_Message("請挑選通報類型")
        End If


        If txt_roof_sdate.Text <> "" Or txt_roof_edate.Text <> "" Then
            If Chk_DateForm(txt_roof_sdate.Text) = False Then
                Chk_data = False
                Show_Message("輕質屋頂作業預定施工期起日期格式不正確，請確認")
            End If
            If Chk_DateForm(txt_roof_edate.Text) = False Then
                Chk_data = False
                Show_Message("輕質屋頂作業預定施工期迄日期格式不正確，請確認")
            End If
        End If

        If txt_tearon_sdate.Text <> "" Or txt_tearon_edate.Text <> "" Or txt_tearoff_sdate.Text <> "" Or txt_tearoff_edate.Text <> "" Then
            If Chk_DateForm(txt_tearon_sdate.Text) = False Then
                Chk_data = False
                Show_Message("施工架組配與拆除預定組配時間起日期格式不正確，請確認")
            End If
            If Chk_DateForm(txt_tearon_edate.Text) = False Then
                Chk_data = False
                Show_Message("施工架組配與拆除預定組配時間迄日期格式不正確，請確認")
            End If
            If Chk_DateForm(txt_tearoff_sdate.Text) = False Then
                Chk_data = False
                Show_Message("施工架組配與拆除預定拆除時間起日期格式不正確，請確認")
            End If
            If Chk_DateForm(txt_tearoff_edate.Text) = False Then
                Chk_data = False
                Show_Message("施工架組配與拆除預定拆除時間迄日期格式不正確，請確認")
            End If
        End If

        If txt_cage_sdate.Text <> "" Or txt_cage_edate.Text <> "" Then
            If Chk_DateForm(txt_cage_sdate.Text) = False Then
                Chk_data = False
                Show_Message("使用吊籠作業預定施工期起日期格式不正確，請確認")
            End If
            If Chk_DateForm(txt_cage_edate.Text) = False Then
                Chk_data = False
                Show_Message("使用吊籠作業預定施工期迄日期格式不正確，請確認")
            End If
        End If

        If (txt_roof_mail.Text) <> "" Then
            If Chk_Email(txt_roof_mail.Text) = False Then
                Chk_data = False
                Show_Message("MAIL格式不正確，請確認")
            End If
        End If

        '檢查必填欄位是否空白
        If txt_com_name.Text = "" Or ddl_att_area.SelectedValue = "" Or ddl_att_road.SelectedValue = "" Or txt_att_add1.Text = "" _
                                  Or txt_att_name1.Text = "" Or txt_att_tel1.Text = "" Or txt_att_name2.Text = "" Or txt_att_tel2.Text = "" _
                                  Or txt_pet_date.Text = "" Then
            Chk_data = False
            Show_Message("必填欄位不得為空白，請確認")
        End If
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '接收default頁面變數
        If Not IsPostBack Then
            SetClientCheck()
            Set_DDL_Option("ZipCode", "110", ddl_att_area, "", "---請選擇---", "")
            Set_DDL_Option(ddl_att_area.SelectedValue, "", ddl_att_road, "", "---請選擇---", "")
            Set_CKL_Option("roof_tool1", "", chk_roof_tool1, "", "", "")
            Set_CKL_Option("tear_tool1", "", chk_tear_tool1, "", "", "")
            Set_CKL_Option("cage_tool1", "", chk_cage_tool1, "", "", "")
	    Set_CKL_Option("rtype2", "",chk_rtype2, "", "", "")

            'Set_DDL_Option("work_type2", "", ddl_work_type2, "", "---請選擇---", "")
            'Set_DDL_Option("work_type3", "", ddl_work_type3, "", "---請選擇---", "")
            'Set_DDL_Option("work_type4", "", ddl_work_type4, "", "---請選擇---", "")
            'Set_DDL_Option("work_type5", "", ddl_work_type5, "", "---請選擇---", "")
            'Set_DDL_Option("work_type6", "", ddl_work_type6, "", "---請選擇---", "")
            'Set_DDL_Option("work_type7", "", ddl_work_type7, "", "---請選擇---", "")
            'Set_DDL_Option("work_type8", "", ddl_work_type8, "", "---請選擇---", "")
            'Set_DDL_Option("work_type9", "", ddl_work_type9, "", "---請選擇---", "")
            'Set_DDL_Option("use_tool1", "", ddl_use_tool1, "", "---請選擇---", "")
            'Set_DDL_Option("use_tool2", "", ddl_use_tool2, "", "---請選擇---", "")
            'Set_DDL_Option("use_tool3", "", ddl_use_tool3, "", "---請選擇---", "")
            'Set_DDL_Option("use_tool4", "", ddl_use_tool4, "", "---請選擇---", "")
            'Set_DDL_Option("use_tool5", "", ddl_use_tool5, "", "---請選擇---", "")

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

            If Pgm_type = "MDY" Then l_PgmType.Text = "[修改模式]" Else l_PgmType.Text = "[新增模式]"

            Select Case Pgm_type
                Case "ADD"
                    txt_pet_date.Text = Format(Now, "yyyy/MM/dd")
                Case "MDY", "COPY"
                    Dim Dt_tmp As DataTable = Get_DataSet(s_prgcode, "Detail", txtPrimary.Text).Tables(0)
                    If Dt_tmp.Rows.Count > 0 Then
                        txt_roof_cod.Text = Trim(Dt_tmp.Rows(0).Item("roof_cod") & "")
                        txt_com_name.Text = Trim(Dt_tmp.Rows(0).Item("com_name") & "")
                        ddl_att_area.SelectedValue = Trim(Dt_tmp.Rows(0).Item("att_area") & "")
                        ddl_att_road.Items.Clear()
                        Set_DDL_Option(ddl_att_area.SelectedValue, "", ddl_att_road, "", "---請選擇---", "")
                        ddl_att_road.SelectedValue = Trim(Dt_tmp.Rows(0).Item("att_road") & "")
                        txt_att_add1.Text = Trim(Dt_tmp.Rows(0).Item("att_add1") & "")
                        txt_att_name1.Text = Trim(Dt_tmp.Rows(0).Item("att_name1") & "")
                        txt_att_tel1.Text = Trim(Dt_tmp.Rows(0).Item("att_tel1") & "")
                        txt_att_name2.Text = Trim(Dt_tmp.Rows(0).Item("att_name2") & "")
                        txt_att_tel2.Text = Trim(Dt_tmp.Rows(0).Item("att_tel2") & "")
                        txt_roof_mail.Text = Trim(Dt_tmp.Rows(0).Item("roof_mail") & "")
                        txt_pet_date.Text = Trim(Dt_tmp.Rows(0).Item("pet_date") & "")
                        Set_CKL_Click(Trim(Dt_tmp.Rows(0).Item("roof_tool1") & ""), chk_roof_tool1, "roof_tool1")
                        txt_roof_sdate.Text = Trim(Dt_tmp.Rows(0).Item("roof_sdate") & "")
                        txt_roof_edate.Text = Trim(Dt_tmp.Rows(0).Item("roof_edate") & "")
                        txt_roof_height.Text = Trim(Dt_tmp.Rows(0).Item("roof_height") & "")
                        txt_floor_height.Text = Trim(Dt_tmp.Rows(0).Item("floor_height") & "")
                        txt_roof_memo.Text = Trim(Dt_tmp.Rows(0).Item("roof_memo") & "")
                        Set_CKL_Click(Trim(Dt_tmp.Rows(0).Item("tear_tool1") & ""), chk_tear_tool1, "tear_tool1")
			Set_CKL_Click(Trim(Dt_tmp.Rows(0).Item("rtype2") & ""), chk_rtype2, "rtype2")
                        txt_tearon_sdate.Text = Trim(Dt_tmp.Rows(0).Item("tearon_sdate") & "")
                        txt_tearon_edate.Text = Trim(Dt_tmp.Rows(0).Item("tearon_edate") & "")
                        txt_tearoff_sdate.Text = Trim(Dt_tmp.Rows(0).Item("tearoff_sdate") & "")
                        txt_tearoff_edate.Text = Trim(Dt_tmp.Rows(0).Item("tearoff_edate") & "")
                        txt_tear_height.Text = Trim(Dt_tmp.Rows(0).Item("tear_height") & "")
                        txt_tearfloor_height.Text = Trim(Dt_tmp.Rows(0).Item("tearfloor_height") & "")
                        txt_tear_memo.Text = Trim(Dt_tmp.Rows(0).Item("tear_memo") & "")
                        Set_CKL_Click(Trim(Dt_tmp.Rows(0).Item("cage_tool1") & ""), chk_cage_tool1, "cage_tool1")
                        txt_cage_sdate.Text = Trim(Dt_tmp.Rows(0).Item("cage_sdate") & "")
                        txt_cage_edate.Text = Trim(Dt_tmp.Rows(0).Item("cage_edate") & "")
                        txt_cageroof_height.Text = Trim(Dt_tmp.Rows(0).Item("cageroof_height") & "")
                        txt_cagefloor_height.Text = Trim(Dt_tmp.Rows(0).Item("cagefloor_height") & "")
                        txt_cage_memo.Text = Trim(Dt_tmp.Rows(0).Item("cage_memo") & "")
                        selwork.SelectedValue = Trim(Dt_tmp.Rows(0).Item("roof_source") & "")

                        REM 呈現照片
                        Dim sPath As String = ws.Get_SysPara("LSI021A_PATH", "")
                        REM Image1.ImageUrl = sPath & Trim(Dt_tmp.Rows(0).Item("roof_path2") & "")
                        Dim sTmpArr = Split(Trim(Dt_tmp.Rows(0).Item("roof_path1") & ""), ",")

Dim sTmpArr2 = Split(Trim(Dt_tmp.Rows(0).Item("roof_path2") & ""), ",")


                        REM IMG1.ImageUrl = "no.jpg"
IMG2.ImageUrl = "no.jpg"
IMG3.ImageUrl = "no.jpg"
IMG4.ImageUrl = "no.jpg"
IMG5.ImageUrl = "no.jpg"
IMG6.ImageUrl = "no.jpg"
IMG7.ImageUrl = "no.jpg"
IMG8.ImageUrl = "no.jpg"
IMG9.ImageUrl = "no.jpg"

IMG2.Visible="False"
IMG3.Visible="False"
IMG4.Visible="False"
IMG5.Visible="False"
IMG6.Visible="False"
IMG7.Visible="False"
IMG8.Visible="False"
IMG9.Visible="False"
IMG10.Visible="False"
IMG11.Visible="False"
IMG12.Visible="False"
IMG13.Visible="False"
IMG14.Visible="False"
IMG15.Visible="False"
IMG16.Visible="False"
IMG17.Visible="False"

if (Dt_tmp.Rows(0).Item("roof_path1")<>"" ) then 
                        For i = LBound(sTmpArr) To UBound(sTmpArr)
                        
			dim vfilename=Split(Trim(sTmpArr(i) & ""), ".")

                                REM If (vfilename(1) = "pdf" Or vfilename(1) = "doc" Or vfilename(1) = "docx") Then
                                REM     IMG1.ImageUrl = "file.jpg"
                                REM     HyperLink0.NavigateUrl = sPath & sTmpArr(i)
                                REM Else
                                REM     IMG1.ImageUrl = sPath & sTmpArr(i)
                                REM     HyperLink0.NavigateUrl = sPath & sTmpArr(i)
                                REM End If


                            Next

                        End If
						REM 2024.02.25 修改路徑參數為指定 s_Image_Navigate_Url		
						Dim s_Image_Navigate_Url As String =  ""
				
            If (Dt_tmp.Rows(0).Item("roof_path2")<>"" ) then
                        For j = LBound(sTmpArr2) To UBound(sTmpArr2)
								REM 2024.02.25，修補輸入為過濾或編碼漏洞，路徑參數變更		
						 s_Image_Navigate_Url =  sPath & HttpUtility.HtmlEncode(sTmpArr2(j)) 		
				if (j=0) then 
					IMG2.ImageUrl = s_Image_Navigate_Url
 					HyperLink2.NavigateUrl =s_Image_Navigate_Url
 					IMG2.Visible="True"
				end if
				if (j=1) then 
				IMG3.Visible="True"
					IMG3.ImageUrl = s_Image_Navigate_Url 
				HyperLink3.NavigateUrl =s_Image_Navigate_Url		
				end if
				
				if (j=2) then 
					IMG4.Visible="True"
					IMG4.ImageUrl = s_Image_Navigate_Url
					HyperLink4.NavigateUrl =s_Image_Navigate_Url 
				end if

				if (j=3) then 
					IMG5.Visible="True"
					IMG5.ImageUrl = s_Image_Navigate_Url
					HyperLink4.NavigateUrl =s_Image_Navigate_Url 
				end if

				if (j=4) then 
					IMG6.Visible="True"
					IMG6.ImageUrl = s_Image_Navigate_Url
					HyperLink4.NavigateUrl =s_Image_Navigate_Url 
				end if
				if (j=5) then 
					IMG7.Visible="True"
					IMG7.ImageUrl = s_Image_Navigate_Url
					HyperLink4.NavigateUrl =s_Image_Navigate_Url 
				end if
				if (j=6) then 
					IMG8.Visible="True"
					IMG8.ImageUrl = s_Image_Navigate_Url
					HyperLink4.NavigateUrl =s_Image_Navigate_Url 
				end if
				if (j=7) then 
					IMG9.Visible="True"
					IMG9.ImageUrl = s_Image_Navigate_Url
					HyperLink4.NavigateUrl =s_Image_Navigate_Url 
				end if
				if (j=8) then 
					IMG10.Visible="True"
					IMG10.ImageUrl = s_Image_Navigate_Url
					HyperLink4.NavigateUrl =s_Image_Navigate_Url 
				end if
				if (j=9) then 
					IMG11.Visible="True"
					IMG11.ImageUrl = s_Image_Navigate_Url
					HyperLink4.NavigateUrl =s_Image_Navigate_Url 
				end if
				if (j=10) then 
					IMG12.Visible="True"
					IMG12.ImageUrl = s_Image_Navigate_Url
					HyperLink4.NavigateUrl =s_Image_Navigate_Url 
				end if
				if (j=11) then 
					IMG13.Visible="True"
					IMG13.ImageUrl = s_Image_Navigate_Url
					HyperLink4.NavigateUrl =s_Image_Navigate_Url 
				end if
				if (j=12) then 
					IMG14.Visible="True"
					IMG14.ImageUrl = s_Image_Navigate_Url
					HyperLink4.NavigateUrl =s_Image_Navigate_Url 
				end if
				if (j=13) then 
				IMG14.Visible="True"
					IMG15.ImageUrl = s_Image_Navigate_Url
					HyperLink4.NavigateUrl =s_Image_Navigate_Url 
				end if
				if (j=14) then 
					IMG14.Visible="True"
					IMG16.ImageUrl = s_Image_Navigate_Url
					HyperLink4.NavigateUrl =s_Image_Navigate_Url 
				end if
				if (j=15) then 
					IMG14.Visible="True"
					IMG17.ImageUrl = s_Image_Navigate_Url
					HyperLink4.NavigateUrl =s_Image_Navigate_Url 
				end if




                        Next
end if
                        If Trim(Dt_tmp.Rows(0).Item("roof_path1") & "") <> "" Then
                            REM 點檢表
                            HyperLink1.Text = "點此下載"
                            HyperLink1.NavigateUrl = sPath & Trim(Dt_tmp.Rows(0).Item("roof_path1") & "")
                        End If

                    End If

                    If Pgm_type = "COPY" Then
                        txt_roof_cod.Text = ""
                        txt_roof_cod.Enabled = False
                    Else
                        txt_roof_cod.Enabled = False
                        Bt_send.Visible = True
                    End If
            End Select
        End If
        img_date1.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_pet_date")
        roof_sdate1.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_roof_sdate")
        roof_edate1.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_roof_edate")
        tearon_sdate1.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_tearon_sdate")
        tearon_edate1.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_tearon_edate")
        tearoff_sdate1.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_tearoff_sdate")
        tearoff_edate1.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_tearoff_edate")
        cage_sdate1.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_cage_sdate")
        cage_edate1.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_cage_edate")
        l_ErrMsg.Text = ""
    End Sub

    Protected Sub SetClientCheck()
        REM 網頁伺服器端檢查點建置
        REM txt_par_value.Attributes("onblur") = "javascript:IsEmpty(this);"
    End Sub
    Protected Sub Bt_BackUp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Bt_BackUp.Click
        Response.Redirect("Default.aspx?subwhere=" & Replace(txtSubWhere.Text, "%", "$") & "&order=" &  HttpUtility.HtmlEncode(txtOrder.Text )  )
    End Sub
    Protected Sub Show_Message(ByVal sMsg As String)
        Dim ErrMsg_Type As String = ws.Get_SysPara("ErrMsg_Type", "0")

        If ErrMsg_Type = "0" Or ErrMsg_Type = "2" Then
            REM 彈出式對話方塊
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('" & sMsg & "');</script>"))
        End If

        If ErrMsg_Type = "1" Or ErrMsg_Type = "2" Then
            REM 下方紅字訊息
            l_ErrMsg.Text &= sMsg & "<br />"
        End If
    End Sub

    Protected Sub ddl_att_area_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddl_att_area.SelectedIndexChanged
        ddl_att_road.Items.Clear()
        Set_DDL_Option(ddl_att_area.SelectedValue, "", ddl_att_road, "", "---請選擇---", "")
    End Sub

    Protected Sub Bt_send_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Bt_send.Click
        REM 發送Mail通知
        Dim dtMail As New DataTable
        Dim column1 As New DataColumn
        Dim column2 As New DataColumn
        column1.DataType = System.Type.GetType("System.String")
        column2.DataType = System.Type.GetType("System.String")
        dtMail.Columns.Add(column1)
        dtMail.Columns.Add(column2)
        REM 加入通報人Mail
        Dim rowMail As DataRow = dtMail.NewRow
        rowMail(0) = txt_roof_mail.Text
        rowMail(1) = txt_att_name2.Text
        dtMail.Rows.Add(rowMail)
        REM 加入承辦人信箱
        Dim MailList As String() = Split(Get_SysPara("RF_receive"), "|")
        For i As Integer = 0 To UBound(MailList)
            If String.IsNullOrEmpty(MailList(i)) = False Then
                Dim rowTmp As DataRow = dtMail.NewRow
                rowTmp(0) = MailList(i)
                rowTmp(1) = "輕質屋頂與施工架及吊籠作業通報承辦人"
                dtMail.Rows.Add(rowTmp)
            End If
        Next
        REM 郵件內容
        REM 郵件內容
        Dim bFirst As Boolean = True

        REM Dim sWorkFloorList As String = ""
        REM For i As Integer = 0 To ckl_work_floor_type.Items.Count - 1
        REM     If ckl_work_floor_type.Items(i).Selected = True Then
        REM         If bFirst = False Then sWorkFloorList &= "、"
        REM         bFirst = False
        REM         sWorkFloorList &= ckl_work_floor_type.Items(i).Text
        REM     End If
        REM Next
        bFirst = True
        Dim rooftoollist As String = ""
        For i As Integer = 0 To chk_roof_tool1.Items.Count - 1
            If chk_roof_tool1.Items(i).Selected = True Then
                If bFirst = False Then rooftoollist &= "、"
                bFirst = False
                rooftoollist &= chk_roof_tool1.Items(i).Text
            End If
        Next
 Dim rtype2list As String = ""
      For i As Integer = 0 To chk_rtype2.Items.Count - 1
            If chk_rtype2.Items(i).Selected = True Then
                If bFirst = False Then rtype2list &= "、"
                bFirst = False
                rtype2list &= chk_rtype2.Items(i).Text
            End If
        Next

        Dim show1 As String = ""
        Dim show11 As String = ""
        If txt_roof_memo.Text <> "" Then
            If rooftoollist <> "" Then
                show1 = "、其他：" & txt_roof_memo.Text
            Else
                show1 = "其他：" & txt_roof_memo.Text
            End If
        Else
            show11 = "無"
        End If

        Dim teartoollist As String = ""
        For j As Integer = 0 To chk_tear_tool1.Items.Count - 1
            If chk_tear_tool1.Items(j).Selected = True Then
                If bFirst = False Then teartoollist &= "、"
                bFirst = False
                teartoollist &= chk_tear_tool1.Items(j).Text
            End If
        Next
        Dim show2 As String = ""
        Dim show21 As String = ""
        If txt_tear_memo.Text <> "" Then
            If teartoollist <> "" Then
                show2 = "、其他：" & txt_tear_memo.Text
            Else
                show2 = "其他：" & txt_tear_memo.Text
            End If
        Else
            show21 = "無"
        End If

        Dim cagetoollist As String = ""
        For k As Integer = 0 To chk_cage_tool1.Items.Count - 1
            If chk_cage_tool1.Items(k).Selected = True Then
                If bFirst = False Then cagetoollist &= "、"
                bFirst = False
                cagetoollist &= chk_cage_tool1.Items(k).Text
            End If
        Next
        Dim show3 As String = ""
        Dim show31 As String = ""
        If txt_cage_memo.Text <> "" Then
            If cagetoollist <> "" Then
                show3 = "、其他：" & txt_cage_memo.Text
            Else
                show3 = "其他：" & txt_cage_memo.Text
            End If
        Else
            show31 = "無"
        End If
        Dim showdet As String = ""
        If rooftoollist <> "" Then
            showdet = "<tr><td rowspan='3'>輕質屋頂作業：</td>"
            showdet = showdet & "<td colspan='3'>預定施工時間" & txt_roof_sdate.Text & "至" & txt_roof_edate.Text & "</td></tr>"
            showdet = showdet & "<tr><td colspan='3'>屋頂高度：" & txt_roof_height.Text & "層樓或樓高" & txt_floor_height.Text & "公尺</td></tr>"
            showdet = showdet & "<tr><td colspan='3'>" & rooftoollist & show1 & "</td></tr>"
            showdet = showdet & "<tr align='center'><td colspan='4'>輕質屋頂營建作業安全檢查重點及注意事項如下：</td></tr>"
            showdet = showdet & "<tr left><td colspan='4'>1.雇主對於高度2公尺以上之屋頂、開口部分、工作臺等場所作業，勞工有遭受墜落危險之虞者，應於該處設置護欄、護蓋或安全網等防護設備。<br>"
            showdet = showdet & "雇主設置前項設備有困難，或因作業之需要臨時將護欄、護蓋或安全網等防護設備拆除者，應採取使勞工使用安全帶等防止墜落致勞工遭受危險之措施。<br>"
            showdet = showdet & "2.勞工於石綿板、鐵皮板、瓦、木板、茅草、塑膠等材料構築之屋頂從事作業時，事業單位為防止勞工踏穿墜落，應於屋架上設置適當強度，且寬度在30公分以上之踏板或裝設安全護網。<br>"
            showdet = showdet & "3.勞工於高差超過1.5公尺以上之場所作業時，應設置能使勞工安全上下之設備。<br>"
            showdet = showdet & "4.在高度2公尺以上之高處作業，勞工有墜落之虞者，事業單位應使勞工確實使用安全帶、安全帽及其他必要之防護具。</td></tr>"
        End If

        If teartoollist <> "" Then
            showdet = "<tr><td rowspan='4'>施工架組配與拆除：</td>"
            showdet = showdet & "<td colspan='3'>預定組配時間" & txt_tearon_sdate.Text & "至" & txt_tearon_edate.Text & "</td></tr>"
            showdet = showdet & "<tr><td colspan='3'>預定拆除時間" & txt_tearoff_sdate.Text & "至" & txt_tearoff_edate.Text & "</td></tr>"
            showdet = showdet & "<tr><td colspan='3'>施工架組配或拆除高度：" & txt_tear_height.Text & "層樓或高度" & txt_tearfloor_height.Text & "公尺</td></tr>"
            showdet = showdet & "<tr><td colspan='3'>" & teartoollist & show2 & "</td></tr>"
            showdet = showdet & "<tr align='center'><td colspan='4'>施工架組配與拆除作業安全檢查重點及注意事項如下：</td></tr>"
            showdet = showdet & "<tr left><td colspan='4'>1.高度5 公尺以上施工架之構築及拆除，應依結構力學原理妥為設計，置備施工圖說，指派所僱專任工程人員簽章確認強度計算書及施工圖說，並建立按施工圖說施作之查驗機制。<br>"
            showdet = showdet & "2.高度5 公尺以上施工架之組配及拆除作業，應指派施工架組配作業主管於作業現場指揮勞工作業及確認安全衛生設備及措施之有效狀況。<br>"
            showdet = showdet & "3.施工架組立及拆除應設置防止作業勞工墜落之設備，如扶手先行工法或安全母索支柱工法。<br>"
            showdet = showdet & "4.高度2 公尺以上之施工架，工作台應舖滿密接之踏板，踏板間、踏板之工作用板料間之縫隙不得大於3 公分，使無墜落、跌倒之虞。<br>"
            showdet = showdet & "5.施工架上之載重限制應於明顯易見之處明確標示，並規定不得超過其荷重限制及應避免發生不均衡現象。<br>"
            showdet = showdet & "6.雇主對於在高度2公尺以上之高處作業，勞工有墜落之虞者，應使勞工確實使用安全帶、安全帽及其他必要之防護具。<br>"
            showdet = showdet & "前項安全帶之使用，應採用符合國家標準一四二五三規定之背負式安全帶及捲揚式防墜器。<br>"
            showdet = showdet & "7.相關重點事項請參考勞動部職業安全衛生署公布之「施工架作業安全檢查重點及注意事項」。</td></tr>"
        End If

        If cagetoollist <> "" Then
            showdet = "<tr><td rowspan='3'>使用吊籠作業：</td>"
            showdet = showdet & "<td colspan='3'>預定施工時間起" & txt_cage_sdate.Text & "至" & txt_cage_edate.Text & "</td></tr>"
            showdet = showdet & "<tr><td colspan='3'>建築物高度：" & txt_cagefloor_height.Text & "層樓或樓高" & txt_cageroof_height.Text & "公尺</td></tr>"
            showdet = showdet & "<tr><td colspan='3'>" & cagetoollist & show3 & "</td></tr>"
 showdet = showdet & "<tr><td colspan='3'>作業項目:" & rtype2list & "</td></tr>"

            showdet = showdet & "<tr align='center'><td colspan='4'>使用吊籠作業安全注意事項</td></tr>"
            showdet = showdet & "<tr left><td colspan='4'>●機具合格證證明、人員操作執照證明。<br>"
            showdet = showdet & "●背負式安全帶、安全帽、救命母索、防墜器。<br>"
            showdet = showdet & "●作業區設置、隔離與管制。<br>"
            showdet = showdet & "●穩固繫掛點，救命母索分長短，安全帶防墜器用上身，短索搭設機具用，長索作業救命母索用。</td></tr>"
            showdet = showdet & "<tr><td colspan='4'>支撐設備要做好，機具本體檢查好，鋼索巡檢有看好，安全防護有顧好，生命財產一定保。</td></tr>"
        End If


        Dim sBody As String = "" & _
            "<table cellpadding=5 cellspacing=0 bordercolor=#7EBAD1 border=1 bgcolor=#ffffff style='border-collapse:collapse;'" & _
            "width='720' summary='排版表格' align=center class='c12'>" & _
            "<tr bgcolor='#7EBAD1' align='center'><td colspan='4'><p><strong>輕質屋頂與施工架及吊籠作業通報</strong></p></td></tr>" & _
            "<tr>" & _
            "  <td width='25%'>工程名稱：</td>" & _
            "  <td width='30%'>" & txt_com_name.Text & "</td>" & _
            "  <td width='20%'>地址：</td>" & _
            "  <td>臺北市" & ddl_att_area.SelectedItem.Text & ddl_att_road.SelectedItem.Text & txt_att_add1.Text & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td width='20%'>廠商名稱：</td>" & _
            "  <td width='30%'>" & txt_att_name1.Text & "</td>" & _
            "  <td width='20%'>聯絡電話：</td>" & _
            "  <td width='30%'>" & txt_att_tel1.Text & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td width='20%'>現場負責人：</td>" & _
            "  <td >" & txt_att_name2.Text & "</td>" & _
            "  <td width='20%'>行動電話：</td>" & _
            "  <td >" & txt_att_tel2.Text & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td >聯絡Email：</td>" & _
            "  <td colspan='3'>" & txt_roof_mail.Text & "</td>" & _
            "<tr>" & _
            "  <td>通報日期：</td>" & _
            "  <td colspan='3'>"& txt_pet_date.Text & "</td>" & _
            "</tr>" & _
            "<tr>" & _
            "  <td>備註：</td>" & _
            "  <td colspan='3'>"& " 完工請補上傳檔案至 " &  Server.HTMLEncode("https://service2.lio.gov.taipei/LSIO_F/RoofWork/Upload?code=") & Server.HTMLEncode(txt_roof_cod.Text) & "</td>" & _
            "</tr>" & _			
            "" & showdet & "" & _
            "</table>"
			

        If SendMail("臺北市勞動檢查處-申辦服務-輕質屋頂與施工架及吊籠作業通報-送出成功", dtMail, sBody, s_prgcode) Then
            Show_Message("送出成功")
        Else
            Show_Message("送出失敗，請通知系統管理員")
            Exit Sub
        End If
    End Sub


End Class
