Imports System.Data
Imports System.Drawing

Partial Class basic_LSI020A_Default
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI020A"
    Dim s_prgname As String = ws.Get_PrgName(s_prgcode)
    Dim UserLimit1 As String
    Dim bPrint As Boolean = False

    '程式自定區 以下

    Protected Sub setFields()
        '建立資料繫結欄位
        Dim fix_code As New BoundField()
        Dim att_area As New BoundField()
        Dim com_name As New BoundField()
        Dim safe_item As New BoundField()
        Dim work_date_s As New BoundField()
        Dim work_date_e As New BoundField()
        Dim pet_date As New BoundField()


        Dim work_floor As New BoundField()
        Dim work_floor_type As New BoundField()
        Dim fix_type As New BoundField()
        Dim work_money As New BoundField()
        Dim att_cname1 As New BoundField()
        Dim att_name1 As New BoundField()
        Dim att_tel1 As New BoundField()
        Dim att_cname2 As New BoundField()
        Dim att_name2 As New BoundField()
        Dim att_tel2 As New BoundField()
        Dim att_mail As New BoundField()


        '指定資料來源欄位
        fix_code.DataField = "fix_code"
        att_area.DataField = "att_areas"
        com_name.DataField = "com_names"
        safe_item.DataField = "safe_item"
        work_date_s.DataField = "work_date_s"
        work_date_e.DataField = "work_date_e"
        pet_date.DataField = "pet_date"

        work_floor.DataField = "work_floor"
        work_floor_type.DataField = "work_floor_type_c"
        fix_type.DataField = "fix_type_c"
        work_money.DataField = "work_money"
        att_cname1.DataField = "att_cname1"
        att_name1.DataField = "att_name1"
        att_tel1.DataField = "att_tel1"
        att_cname2.DataField = "att_cname2"
        att_name2.DataField = "att_name2"
        att_tel2.DataField = "att_tel2"
        att_mail.DataField = "att_mail"

        '設定欄位標頭名稱
        fix_code.HeaderText = "資料編號"
        att_area.HeaderText = "區域"
        com_name.HeaderText = "學校"
        safe_item.HeaderText = "合約工程名稱"
        work_date_s.HeaderText = "施工期間起"
        work_date_e.HeaderText = "施工期間止"
        pet_date.HeaderText = "通報日期"


        'work_floor.HeaderText = "作業樓層"


        '設定欄位是否自動換行(不換行的再設定)
        'prg_code.ItemStyle.Wrap = False

        '設定排序所對應的欄位
        fix_code.SortExpression = "fix_code"
        att_area.SortExpression = "att_area"
        com_name.SortExpression = "com_name"
        safe_item.SortExpression = "safe_item"
        work_date_s.SortExpression = "work_date_s"
        work_date_e.SortExpression = "work_date_e"
        pet_date.SortExpression = "pet_date"
        

        '欄位寬度(%):要留6%給編輯圖示
        fix_code.ItemStyle.Width = Unit.Percentage(11)
        att_area.ItemStyle.Width = Unit.Percentage(8)
        com_name.ItemStyle.Width = Unit.Percentage(22)
        safe_item.ItemStyle.Width = Unit.Percentage(24)
        work_date_s.ItemStyle.Width = Unit.Percentage(10)
        work_date_e.ItemStyle.Width = Unit.Percentage(10)
        pet_date.ItemStyle.Width = Unit.Percentage(9)

        '設定欄位標題顏色
        fix_code.HeaderStyle.CssClass = "grid1"
        att_area.HeaderStyle.CssClass = "grid1"
        com_name.HeaderStyle.CssClass = "grid1"
        safe_item.HeaderStyle.CssClass = "grid1"
        work_date_s.HeaderStyle.CssClass = "grid1"
        work_date_e.HeaderStyle.CssClass = "grid1"
        pet_date.HeaderStyle.CssClass = "grid1"
        '設定欄位的水平對齊
        fix_code.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        att_area.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        com_name.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        safe_item.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        work_date_s.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        work_date_e.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        pet_date.ItemStyle.HorizontalAlign = HorizontalAlign.Center

        '設定欄位是否顯示(不顯示的再設定)
        'com_name.Visible = False
        fix_code.HtmlEncode = True
        GridView1.Columns.Add(fix_code)
        GridView1.Columns.Add(att_area)
        GridView1.Columns.Add(com_name)
        GridView1.Columns.Add(safe_item)
        GridView1.Columns.Add(work_date_s)
        GridView1.Columns.Add(work_date_e)
        GridView1.Columns.Add(pet_date)
        GridView1.Columns.Add(work_floor)
        GridView1.Columns.Add(work_floor_type)
        GridView1.Columns.Add(fix_type)
        GridView1.Columns.Add(work_money)
        GridView1.Columns.Add(att_cname1)
        GridView1.Columns.Add(att_name1)
        GridView1.Columns.Add(att_tel1)
        GridView1.Columns.Add(att_cname2)
        GridView1.Columns.Add(att_name2)
        GridView1.Columns.Add(att_tel2)
        GridView1.Columns.Add(att_mail)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) Then
            setGridViewStyle()
            setFields()
            setCheckBox()
            setGridViewPage()

            Set_DDL_Option("ZipCode1", "", ddl_att_area, "", "---請選擇---", "")
            Set_DDL_Option(ddl_att_area.SelectedValue, "", ddl_com_name, "", "", "B")

            Set_CKL_Option("ZipCode1", "", ckl_att_area, "", "", "")

            Set_CKL_Option("work_floor", "", ckl_work_floor_type, "", "", "")
            Set_CKL_Option("fix_type1", "", ckl_fix_type, "", "", "")
            ''進階搜尋預設值
            'txt_work_date_s_s.Text = Format(Now, "yyyy/MM/") & "01"
            'txt_work_date_s_e.Text = Format(Now, "yyyy/MM/dd")
            'txt_work_date_e_s.Text = Format(Now, "yyyy/MM/") & "01"
            'txt_work_date_e_e.Text = Format(Now, "yyyy/MM/dd")

            '進階搜尋增加點
            WebSubMenu.setName("GridView1", "SqlDS2", "txtOrder", "txtSubWhere")

            '設定Row的鍵值組成，具有唯一性
            Dim KeyNames() As String = {"fix_code"}
            GridView1.DataKeyNames = KeyNames
            ws.InsUsrRec(s_prgname, "SRC", s_prgcode, Session("usr_code"), Session("usr_ip"), "", "", "")
            txtOrder.Text = " ORDER BY fix_code DESC"
        End If

        '程式變數設定
        UserLimit1 = ws.Get_limit(Session("usr_code"), s_prgcode)
        setLimit(UserLimit1)

        '設定查詢欄位
        SqlDS2.ConnectionString = ws.Get_ConnStr("con_db")
        SqlDS2.SelectCommand = Get_SqlStr(s_prgcode, "ddlField")

        '程式抬頭
        Label1.Text = s_prgname & "(" & s_prgcode & ")"

        '取得預設條件
        If Trim(Request("subwhere") & "") <> "" Then txtSubWhere.Text = Replace(Request("subwhere"), "$", "%")
        '進階搜尋增加點
        If Trim(Request("order") & "") <> "" Then txtOrder.Text = Request("order")

        '設定onclick
        CBL_Field1.Items(0).Attributes.Add("onclick", "Chk_All_1()")
        'IB_import.Attributes("onclick") = "subwindow=window.open('FileUpload.aspx','FileUpload','height=260,width=620,top=200,left=400,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no','');subwindow.focus();return false;"
        img_date1.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_work_date_s_s")
        img_date2.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_work_date_s_e")
        img_date3.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_work_date_e_s")
        img_date4.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_work_date_e_e")

        img_date5.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_ins_date_s")
        img_date6.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_ins_date_e")
        '取得資料
        setGridViewData()
    End Sub

    Protected Sub setGridViewData()
        '取得GridView資料

        '設定SqlDataSource連線及Select命令
        SqlDS1.ConnectionString = ws.Get_ConnStr("con_db")
        '進階搜尋增加點
        If txtSubWhere.Text <> "" Then
            SqlDS1.SelectCommand = Get_SqlStr(s_prgcode, "SELECT", txtSubWhere.Text, txtOrder.Text)
        Else
            SqlDS1.SelectCommand = Get_SqlStr(s_prgcode, "SELECT", "WHERE 1=0")
        End If

        '設定Delete命令
        Dim CtrPara As New ControlParameter("datakeys", "GridView1", "SelectedValue")
        SqlDS1.DeleteParameters.Add(CtrPara)
        SqlDS1.DeleteCommand = Get_SqlStr(s_prgcode, "DELETE")

        '設定GridView資料來源ID
        If GridView1.DataSourceID Is Nothing Then
            GridView1.DataSourceID = SqlDS1.ID
        End If
    End Sub

    '程式自定區 以上





    '程式固定區-不要修改 以下
    Protected Sub setLimit(ByVal Climit As String)
        '依據權限控制物件
        GridView1.Columns(0).Visible = CBool(InStr(Climit, "D") > 0)
        GridView1.Columns(1).Visible = CBool(InStr(Climit, "M") > 0)
        GridView1.Columns(2).Visible = CBool(InStr(Climit, "A") > 0)
        GridView1.Columns(2).Visible = False
        Button1.Enabled = CBool(InStr(Climit, "S") > 0)
        Button4.Enabled = CBool(InStr(Climit, "A") > 0)
        Bt_subMenu.Enabled = CBool(InStr(Climit, "S") > 0)
        IB_export.Enabled = CBool(InStr(Climit, "S") > 0)
    End Sub

    Protected Sub setGridViewPage()
        '設定GridView的分頁樣式
        GridView1.PagerSettings.Mode = PagerButtons.NextPreviousFirstLast
        GridView1.PagerSettings.Visible = False
        'GridView1.PagerSettings.FirstPageImageUrl = "~/Images/First.gif"
        'GridView1.PagerSettings.FirstPageText = "第一頁"
        'GridView1.PagerSettings.PreviousPageImageUrl = "~/Images/Previous.gif"
        'GridView1.PagerSettings.PreviousPageText = "上一頁"
        'GridView1.PagerSettings.NextPageImageUrl = "~/Images/Next.gif"
        'GridView1.PagerSettings.NextPageText = "下一頁"
        'GridView1.PagerSettings.LastPageImageUrl = "~/Images/Last.gif"
        'GridView1.PagerSettings.LastPageText = "最後頁"
    End Sub

    Protected Sub setGridViewStyle()
        GridView1.Attributes.Add("style", "word-break:break-all;word-wrap:break-word")
        '設定GridView外觀樣式
        GridView1.AutoGenerateColumns = False

        '設定GridView屬性
        GridView1.AllowPaging = True   '設定分頁
        GridView1.AllowSorting = True  '設定排序
        GridView1.Font.Size = 12       '設定字型大小
        GridView1.GridLines = GridLines.Both '設定格線
        GridView1.PageSize = CInt(ws.Get_SysPara("pagecount", "20"))
        '非同步Callback模式
        GridView1.EnableSortingAndPagingCallbacks = False
        '分頁位置
        GridView1.PagerSettings.Position = PagerPosition.TopAndBottom
        '分頁對齊
        GridView1.PagerStyle.HorizontalAlign = HorizontalAlign.Justify

        'GridView1.HeaderStyle.BackColor = Color.Tan
        'GridView1.RowStyle.BackColor = Color.LightGoldenrodYellow
        'GridView1.AlternatingRowStyle.BackColor = Color.PaleGoldenrod
        'GridView1.HeaderStyle.ForeColor = Color.Black
        'GridView1.PagerStyle.BackColor = Color.Goldenrod
        '設定選擇列背景顏色
        'GridView1.SelectedRowStyle.BackColor = Color.LightBlue

        GridView1.Style("width") = "100%"
    End Sub

    Protected Sub GridView1_DataBinding(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.DataBinding
        '取得資料
        setGridViewData()
    End Sub

    Protected Sub GridView1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.DataBound
        '取得資料總筆數
        Dim dt_tmp As DataTable = Get_DataSet(s_prgcode, "SqlStr", SqlDS1.SelectCommand).Tables(0)
        If dt_tmp Is Nothing Then
            Label2.Text = "總筆數：0/總頁數:" & GridView1.PageCount.ToString() & "/每頁筆數:" & GridView1.PageSize.ToString()
        Else
            Label2.Text = "總筆數：" & CStr(dt_tmp.Rows.Count) & "/總頁數:" & GridView1.PageCount.ToString() & "/每頁筆數:" & GridView1.PageSize.ToString()

            ''有資料才秀()
            'If CInt(dt_tmp.rows.count) > 0 Then
            '    Dim bottonPagerRow As GridViewRow = GridView1.BottomPagerRow
            '    Dim bottonPagerNo As New Label()
            '    '進階搜尋增加點
            '    If Not bottonPagerRow Is Nothing Then
            '        bottonPagerNo.Text = "目前所在分頁碼 (" & (GridView1.PageIndex + 1) & "/" & GridView1.PageCount.ToString() & ")"
            '        bottonPagerRow.Cells(0).Controls.Add(bottonPagerNo)
            '    End If
            'End If

            '設定分頁訊息
            l_PageMsg.Visible = False
            If GridView1.PageCount > 1 Then
                l_PageMsg.Visible = True
                l_PageMsg.Text = "　目前所在分頁碼 (" & (GridView1.PageIndex + 1) & "/" & GridView1.PageCount.ToString() & ")"
            End If
        End If
        ViewState("LineNo") = 0

        If bPrint = False Then
            For Each Row As GridViewRow In GridView1.Rows
                'If Row.Cells(20).Text >= Format(Now, "yyyy/MM/dd") And Row.Cells(20).Text <= Format(Now.AddDays(3), "yyyy/MM/dd") Then
                'Row.Cells(3).Text &= "<img src='../images/icon/new2.png' />"
                'End If
            Next
        Else
            For Each Row As GridViewRow In GridView1.Rows
                'If Row.Cells(20).Text >= Format(Now, "yyyy/MM/dd") And Row.Cells(20).Text <= Format(Now.AddDays(3), "yyyy/MM/dd") Then
                'Row.Cells(3).Text &= " - NEW"
                ' End If
            Next
        End If

        '設定分頁按鈕是否顯示
        IB_GV1_First1.Visible = False
        IB_GV1_First2.Visible = False
        IB_GV1_Previous1.Visible = False
        IB_GV1_Previous2.Visible = False
        IB_GV1_Next1.Visible = False
        IB_GV1_Next2.Visible = False
        IB_GV1_Last1.Visible = False
        IB_GV1_Last2.Visible = False

        If GridView1.PageCount > 1 Then
            '分頁數量大於一頁
            If GridView1.PageIndex <> 0 Then
                IB_GV1_First1.Visible = True
                IB_GV1_First2.Visible = True
                IB_GV1_Previous1.Visible = True
                IB_GV1_Previous2.Visible = True
            End If

            If GridView1.PageIndex <> GridView1.PageCount - 1 Then
                IB_GV1_Next1.Visible = True
                IB_GV1_Next2.Visible = True
                IB_GV1_Last1.Visible = True
                IB_GV1_Last2.Visible = True
            End If
        End If
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        GridView1.SelectedIndex = e.RowIndex
        Dim txtDataKey As String = GridView1.SelectedDataKey.Value.ToString
        If Not txtDataKey = "" Then
            '設定記錄資料
            Dim sSql As String = Get_SqlStr(s_prgcode, "Detail", txtDataKey)
            Dim sFieldName As String = ws.Get_FieldName(s_prgcode)
            Dim sOldData As String = ws.Get_OldData(sSql, s_prgcode)
            Dim sNewData As String = Get_NewData(s_prgcode, "DEL")

            SqlDS1.Delete()
            ws.InsUsrRec(txtDataKey, "DEL", s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)
        End If
        GridView1.SelectedIndex = -1
    End Sub

    Protected Sub GridView1_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Sorted
        Select Case GridView1.SortDirection
            '設定升降冪外觀樣式
            Case SortDirection.Ascending
                'GridView1.HeaderStyle.BackColor = Color.Tan
                'GridView1.RowStyle.BackColor = Color.LightGoldenrodYellow
                'GridView1.AlternatingRowStyle.BackColor = Color.PaleGoldenrod
                'GridView1.HeaderStyle.ForeColor = Color.Black
                'GridView1.PagerStyle.BackColor = Color.Goldenrod
                'GridView1.SelectedRowStyle.BackColor = Color.LightBlue
                '設定升降冪外觀樣式
            Case SortDirection.Descending
                'GridView1.HeaderStyle.BackColor = Color.MidnightBlue
                'GridView1.RowStyle.BackColor = Color.LightBlue
                'GridView1.AlternatingRowStyle.BackColor = Color.Lavender
                'GridView1.HeaderStyle.ForeColor = Color.White
                'GridView1.PagerStyle.BackColor = Color.LightPink
                'GridView1.SelectedRowStyle.BackColor = Color.LightPink
        End Select
    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        If (e.Row.RowType = DataControlRowType.Header) Then
            Dim myheader As TableCell
            For Each myheader In e.Row.Cells '對每一格 
                If (myheader.HasControls()) Then
                    Dim LinkButton As LinkButton = myheader.Controls(0)
                    If LinkButton.CommandArgument = GridView1.SortExpression Then
                        '否為為排序欄位(myheader.Controls(0))
                        If (GridView1.SortDirection = SortDirection.Descending) Then '依排序方向加入箭號
                            myheader.Controls.Add(New LiteralControl("▼"))
                        Else
                            myheader.Controls.Add(New LiteralControl("▲"))
                        End If
                    ElseIf GridView1.SortExpression = "" Then
                        '進階搜尋增加點
                        If InStr(txtOrder.Text, LinkButton.CommandArgument & " ASC") > 0 Then
                            myheader.Controls.Add(New LiteralControl("▲"))
                        ElseIf InStr(txtOrder.Text, LinkButton.CommandArgument & " DESC") > 0 Then
                            myheader.Controls.Add(New LiteralControl("▼"))
                        End If
                    End If
                End If
            Next
        End If
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.Header Or e.Row.RowType = DataControlRowType.DataRow Then
            ' e.Row.Cells(7).Visible = False
            ' e.Row.Cells(8).Visible = False
            e.Row.Cells(10).Visible = False
            e.Row.Cells(11).Visible = False
            e.Row.Cells(12).Visible = False
            e.Row.Cells(13).Visible = False
            e.Row.Cells(14).Visible = False
            e.Row.Cells(15).Visible = False
            e.Row.Cells(16).Visible = False
            e.Row.Cells(17).Visible = False
            e.Row.Cells(18).Visible = False
            e.Row.Cells(19).Visible = False
            e.Row.Cells(20).Visible = False

            '匯出Excel時顯示設定
            If bPrint = True Then
                '隱藏功能列
                If GridView1.Columns(0).Visible = True Then e.Row.Cells(0).Visible = False
                If GridView1.Columns(1).Visible = True Then e.Row.Cells(1).Visible = False
                'If GridView1.Columns(2).Visible = True Then e.Row.Cells(2).Visible = False

                '選擇全部
                If ckb_all.Checked Then
                    e.Row.Cells(10).Visible = True
                    e.Row.Cells(11).Visible = True
                    e.Row.Cells(12).Visible = True
                    e.Row.Cells(13).Visible = True
                    e.Row.Cells(14).Visible = True
                    e.Row.Cells(15).Visible = True
                    e.Row.Cells(16).Visible = True
                    e.Row.Cells(17).Visible = True
                    e.Row.Cells(18).Visible = True
                    e.Row.Cells(19).Visible = True
                    e.Row.Cells(20).Visible = True
					e.Row.Cells(11).Text = Replace(e.Row.Cells(11).Text, "|", " & ")
					e.Row.Cells(12).Text = Replace(e.Row.Cells(12).Text, "|", " & ")
                    '設定欄位標頭名稱
                    If e.Row.RowType = DataControlRowType.Header Then
                        e.Row.Cells(10).Text = "作業樓層"
                        e.Row.Cells(11).Text = "包括"
                        e.Row.Cells(12).Text = "種類"
                        e.Row.Cells(13).Text = "承攬金額"
                        e.Row.Cells(14).Text = "主辦單位"
                        e.Row.Cells(15).Text = "主辦單位-聯絡人"
                        e.Row.Cells(16).Text = "主辦單位-聯絡電話"
                        e.Row.Cells(17).Text = "承覽廠商"
                        e.Row.Cells(18).Text = "承覽廠商-工地負責人"
                        e.Row.Cells(19).Text = "承覽廠商-手機"
                        e.Row.Cells(20).Text = "通報回覆"
                    End If
                Else
                    '隱藏未選擇欄位
                    For i As Integer = 1 To CBL_Field1.Items.Count - 1
                        If CBL_Field1.Items(i).Selected = False Then
                            e.Row.Cells(CBL_Field1.Items(i).Value).Visible = False
                        End If
						
                    Next
                End If
            End If
        End If

        Select Case e.Row.RowType
            Case DataControlRowType.Header
                e.Row.BackColor = Color.FromArgb(153, 0, 0)
                e.Row.ForeColor = Color.White
            Case DataControlRowType.DataRow
                '建立奇數資料列與偶數資料列的onmouseover及onmouseout的顏色變換
                If (CType(ViewState("LineNo"), Int16) = 0) Then
                    'e.Row.BackColor = Color.FromArgb(255, 251, 214)
                    e.Row.Attributes.Add("onmouseout", "this.style.backgroundColor='" & gridview_row_color("10") & "';this.style.color='" & gridview_row_color("11") & "'")
                    e.Row.Attributes.Add("onmouseover", "this.style.backgroundColor='" & gridview_row_color("30") & "';this.style.color='" & gridview_row_color("31") & "'")
                    ViewState("LineNo") = 1
                Else
                    'e.Row.BackColor = Color.White

                    e.Row.Attributes.Add("onmouseout", "this.style.backgroundColor='" & gridview_row_color("20") & "';this.style.color='" & gridview_row_color("11") & "'")
                    e.Row.Attributes.Add("onmouseover", "this.style.backgroundColor='" & gridview_row_color("30") & "';this.style.color='" & gridview_row_color("31") & "'")

                    ViewState("LineNo") = 0
                End If
        End Select

        Dim is_ok As Boolean
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            is_ok = Chk_UsrLimit(Session("usr_code"), e.Row.Cells(4).Text, Session("grp_code"))
            e.Row.FindControl("ImageButton2").Visible = is_ok
            e.Row.FindControl("ImageButton3").Visible = is_ok
        End If

        'ViewState("LineNo") = 1
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '資料搜尋-單一條件
        If TxtFind.Text <> "" Then
            txtSubWhere.Text = " WHERE " & ddlField.SelectedValue & " LIKE N'%" & RepSql(TxtFind.Text) & "%'"
        Else
            txtSubWhere.Text = " WHERE 1=1"
        End If
        setGridViewData()
        GridView1.DataBind()
    End Sub

  

    Protected Sub Button4_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles Button4.Click
        txtPgmType.Text = "ADD"
        Server.Transfer("edit.aspx")
    End Sub

    Protected Sub IB_export_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_export.Click
        If Panel1.Visible = False Then
            Panel1.Visible = True
        Else
            Panel1.Visible = False
        End If
    End Sub

    Protected Sub bt_excel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bt_excel.Click
        bPrint = True '開始輸出Excel

        Response.Clear()
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")
        Response.AddHeader("content-disposition", "attachment;filename=Export.xls")
        Response.ContentType = "application/vnd.xls"
        Dim sw As System.IO.StringWriter = New System.IO.StringWriter()
        Dim htw As System.Web.UI.HtmlTextWriter = New HtmlTextWriter(sw)

        '關閉換頁跟排序
        GridView1.AllowSorting = False
        GridView1.AllowPaging = False

        '移去不要的欄位
        'GridView1.Columns.RemoveAt(GridView1.Columns.Count - 1)
        GridView1.DataBind()

        '建立假HtmlForm避免以下錯誤
        'Control 'GridView1' of type 'GridView' must be placed inside 
        'a form tag with runat=server. 
        '另一種做法是override VerifyRenderingInServerForm後不做任何事
        '這樣就可以直接GridView1.RenderControl(htw)

        Dim hf As HtmlForm = New HtmlForm()
        Dim NewTable As New Table
        Dim NewRow As New TableRow
        Dim NewCell As New TableCell
        'NewCell.Text = "綜合搜尋"
        NewCell.Attributes.Add("colspan", 1)
        NewCell.Attributes.Add("style", "text-align:center")
        NewRow.Cells.Add(NewCell)
        NewTable.Rows.Add(NewRow)

        Controls.Add(hf)
        hf.Controls.Add(GridView1)
        hf.Controls.Add(NewTable)
        hf.RenderControl(htw)

        Response.Write(sw.ToString())
        Response.End()
    End Sub

    Protected Sub setCheckBox()
        '設置匯出欄位選項
        CBL_Field1.Items.Add(New ListItem("全選", ""))

        For i As Integer = 0 To GridView1.Columns.Count - 1
            If GridView1.Columns(i).HeaderText = "" Then Continue For
            CBL_Field1.Items.Add(New ListItem(GridView1.Columns(i).HeaderText, i))
        Next
    End Sub

    Protected Sub GridView1_SelectedIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSelectEventArgs) Handles GridView1.SelectedIndexChanging
        GridView1.SelectedIndex = e.NewSelectedIndex
        If e.NewSelectedIndex >= 0 Then
            Dim txtDataKey As String = GridView1.SelectedDataKey.Value.ToString
            txtPrimary.Text = txtDataKey
        End If

        If txtPgmType.Text = "ADD" Or txtPgmType.Text = "MDY" Or txtPgmType.Text = "COPY" Then
            Server.Transfer("edit.aspx")
        End If
    End Sub

    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        txtPgmType.Text = e.CommandArgument
    End Sub

    Protected Sub Bt_subMenu_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles Bt_subMenu.Click
        '進階搜尋增加點
        WebSubMenu.show()
    End Sub

    Protected Sub IB_GV1_First1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_First1.Click
        GridView1.PageIndex = 0
        GridView1.DataBind()
    End Sub

    Protected Sub IB_GV1_Previous1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Previous1.Click
        GridView1.PageIndex -= 1
        GridView1.DataBind()
    End Sub

    Protected Sub IB_GV1_Next1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Next1.Click
        GridView1.PageIndex += 1
        GridView1.DataBind()
    End Sub

    Protected Sub IB_GV1_Last1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Last1.Click
        GridView1.PageIndex = GridView1.PageCount - 1
        GridView1.DataBind()
    End Sub

    Protected Sub IB_GV1_First2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_First2.Click
        GridView1.PageIndex = 0
        GridView1.DataBind()
    End Sub

    Protected Sub IB_GV1_Previous2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Previous2.Click
        GridView1.PageIndex -= 1
        GridView1.DataBind()
    End Sub

    Protected Sub IB_GV1_Next2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Next2.Click
        GridView1.PageIndex += 1
        GridView1.DataBind()
    End Sub

    Protected Sub IB_GV1_Last2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Last2.Click
        GridView1.PageIndex = GridView1.PageCount - 1
        GridView1.DataBind()
    End Sub

    '程式固定區-不要修改 以上

    Protected Sub Page_LoadComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LoadComplete
        If (Not IsPostBack) Then
            Button1_Click(sender, e)
        End If
    End Sub

    Protected Sub Bt_ExSearch_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles Bt_ExSearch.Click
        If Panel_ExSearch.Visible = False Then
            Panel_ExSearch.Visible = True
        Else
            Panel_ExSearch.Visible = False
        End If
    End Sub

    Protected Sub Bt_ExSearch__ok_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Bt_ExSearch_ok.Click
        '資料搜尋-進階搜尋2
        '廠商名稱
        txtSubWhere.Text = " WHERE 1=1 "

        '區域包括
        Dim bAreaToolChecked As Boolean = False
        For i As Integer = 0 To ckl_att_area.Items.Count - 1
            If ckl_att_area.Items(i).Selected = True Then
                If bAreaToolChecked = False Then
                    txtSubWhere.Text &= " AND ("
                    bAreaToolChecked = True
                Else
                    txtSubWhere.Text &= " OR "
                End If
                txtSubWhere.Text &= "att_area" & " = N'" & RepSql(ckl_att_area.Items(i).Value) & "'"
            End If
        Next
        If bAreaToolChecked = True Then txtSubWhere.Text &= ")"

        '學校
        If ddl_att_area.SelectedValue <> "" Then
            txtSubWhere.Text &= " AND " & "att_area" & " = N'" & RepSql(ddl_att_area.SelectedValue) & "'"
        End If

        If ddl_com_name.SelectedValue <> "" Then
            txtSubWhere.Text &= " AND " & "com_name" & " = N'" & RepSql(ddl_com_name.SelectedValue) & "'"
        End If

        '施工期間-起
        If RepSql(txt_work_date_s_s.Text) <> "" And RepSql(txt_work_date_s_e.Text) <> "" Then
            txtSubWhere.Text &= " AND " & "work_date_s" & " BETWEEN N'" & RepSql(txt_work_date_s_s.Text) & "' AND '" & RepSql(txt_work_date_s_e.Text) & "'"
        ElseIf RepSql(txt_work_date_s_s.Text) <> "" Then
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('請輸入 [開工日期-起迄] 的結束日期條件');</script>"))
            Exit Sub
        ElseIf RepSql(txt_work_date_s_e.Text) <> "" Then
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('請輸入 [開工日期-起迄] 的開始日期條件');</script>"))
            Exit Sub
        End If
        '施工期間-止
        If RepSql(txt_work_date_e_s.Text) <> "" And RepSql(txt_work_date_e_e.Text) <> "" Then
            txtSubWhere.Text &= " AND " & "work_date_e" & " BETWEEN N'" & RepSql(txt_work_date_e_s.Text) & "' AND '" & RepSql(txt_work_date_e_e.Text) & "'"
        ElseIf RepSql(txt_work_date_e_s.Text) <> "" Then
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('請輸入 [完工日期-起迄] 的結束日期條件');</script>"))
            Exit Sub
        ElseIf RepSql(txt_work_date_e_e.Text) <> "" Then
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('請輸入 [完工日期0起迄] 的開始日期條件');</script>"))
            Exit Sub
        End If

        '通報日期起迄
        If RepSql(txt_ins_date_s.Text) <> "" And RepSql(txt_ins_date_e.Text) <> "" Then
            txtSubWhere.Text &= " AND " & "pet_date" & " BETWEEN N'" & RepSql(txt_ins_date_s.Text) & "' AND '" & RepSql(txt_ins_date_e.Text) & "'"
        ElseIf RepSql(txt_ins_date_s.Text) <> "" Then
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('請輸入 [ 通報日期-起迄] 的結束日期條件');</script>"))
            Exit Sub
        ElseIf RepSql(txt_ins_date_e.Text) <> "" Then
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('請輸入 [ 通報日期-起迄] 的開始日期條件');</script>"))
            Exit Sub
        End If
        '承攬金額超過
        If RepSql(txt_work_money.Text) <> "" Then
            txtSubWhere.Text &= " AND " & "work_money >=" & RepSql(txt_work_money.Text)
        End If
        '作業樓層包括
        Dim bUseToolChecked As Boolean = False
        For i As Integer = 0 To ckl_work_floor_type.Items.Count - 1
            If ckl_work_floor_type.Items(i).Selected = True Then
                If bUseToolChecked = False Then
                    txtSubWhere.Text &= " AND ("
                    bUseToolChecked = True
                Else
                    txtSubWhere.Text &= " OR "
                End If
                txtSubWhere.Text &= "work_floor_type" & " LIKE N'%" & RepSql(ckl_work_floor_type.Items(i).Value) & "%'"
            End If
        Next
        If bUseToolChecked = True Then txtSubWhere.Text &= ")"

        '作業方式
        Dim bSaveToolChecked As Boolean = False
        For i As Integer = 0 To ckl_fix_type.Items.Count - 1
            If ckl_fix_type.Items(i).Selected = True Then
                If bSaveToolChecked = False Then
                    txtSubWhere.Text &= " AND ("
                    bSaveToolChecked = True
                Else
                    txtSubWhere.Text &= " OR "
                End If
                txtSubWhere.Text &= "fix_type" & " LIKE N'%" & RepSql(ckl_fix_type.Items(i).Value) & "%'"
            End If
        Next
        If bSaveToolChecked = True Then txtSubWhere.Text &= ")"

        setGridViewData()
        GridView1.DataBind()
    End Sub

    Protected Sub Bt_ExSearch_no_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Bt_ExSearch_no.Click
        Panel_ExSearch.Visible = False
    End Sub
    Protected Sub ddl_att_area_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddl_att_area.SelectedIndexChanged
        '    txtSubWhere.Text = 0
        ddl_com_name.Items.Clear()
        Set_DDL_Option(ddl_att_area.SelectedValue, "", ddl_com_name, "", "", "B")
    End Sub
  
End Class
