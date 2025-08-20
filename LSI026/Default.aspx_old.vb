Imports System.Data
Imports System.Drawing

Partial Class basic_LSI026_Default
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI026"
    Dim s_prgname As String = ws.Get_PrgName(s_prgcode)
    Dim UserLimit1 As String
    Dim bPrint As Boolean = False

    '程式自定區 以下

    Protected Sub setFields()
        '建立資料繫結欄位
        Dim pc_code As New BoundField()
		Dim eng_name As New BoundField()
        Dim eng_code As New BoundField()
        Dim area_name As New BoundField()
        Dim road_name As New BoundField()
        Dim eng_add As New BoundField()
        Dim com_name As New BoundField()
		Dim pc_date As New BoundField()
		Dim pc_date_s As New BoundField()
        Dim pc_date_e As New BoundField()
        Dim att_tel1 As New BoundField()
        Dim att_name1 As New BoundField()
        Dim att_mail As New BoundField()
        Dim att_cell_phone As New BoundField()      
        Dim eng_floor As New BoundField()
        Dim pc_kind1_c As New BoundField()
		Dim pc_kind2_c As New BoundField()

        '指定資料來源欄位
        pc_code.DataField = "pc_code"
		eng_name.DataField = "eng_name"
		eng_code.DataField = "eng_code"
        area_name.DataField = "area_name"
        road_name.DataField = "road_name"
        eng_add.DataField = "eng_add"
        com_name.DataField = "com_name"
		pc_date.DataField = "pc_date"
		pc_date_s.DataField = "pc_date_s"
        pc_date_e.DataField = "pc_date_e"
        att_tel1.DataField = "att_tel1"
        att_name1.DataField = "att_name1"
        att_mail.DataField = "att_mail"
        att_cell_phone.DataField = "att_cell_phone"       
        eng_floor.DataField = "eng_floor"
        pc_kind1_c.DataField = "pc_kind1_c"
		pc_kind2_c.DataField = "pc_kind2_c"

        '設定欄位標頭名稱
        pc_code.HeaderText = "資料編號"
		eng_name.HeaderText = "工程名稱"
        eng_code.HeaderText = "建號"
        area_name.HeaderText = "工程地址-區"
        road_name.HeaderText = "工程地址-路段"
        eng_add.HeaderText = "工程地址"
		com_name.HeaderText = "通報單位"
		pc_date.HeaderText = "通報日期"
		pc_date_s.HeaderText = "預定開工日期"
        pc_date_e.HeaderText = "預定完工日期"
        att_tel1.HeaderText = "聯絡電話"
        att_name1.HeaderText = "聯絡人"
        att_mail.HeaderText = "電子信箱"
        att_cell_phone.HeaderText = "手機"
        eng_floor.HeaderText = "作業區域及樓層"
        pc_kind1_c.HeaderText = "丁類危險性工作場所主要危害作業"
		pc_kind2_c.HeaderText = "機械作業"


        '設定欄位是否自動換行(不換行的再設定)
        'prg_code.ItemStyle.Wrap = False

        '設定排序所對應的欄位
        pc_code.SortExpression = "pc_code"
		eng_name.SortExpression = "eng_name"
        eng_code.SortExpression = "eng_code"
        area_name.SortExpression = "area_name"
        road_name.SortExpression = "road_name"
        eng_add.SortExpression = "eng_add"
        com_name.SortExpression = "com_name"
		pc_date.SortExpression = "pc_date"
        pc_date_s.SortExpression = "pc_date_s"
        pc_date_e.SortExpression = "pc_date_e"
        'att_tel1.SortExpression = "att_tel1"
        'att_name1.SortExpression = "att_name1"
        'att_mail.SortExpression = "att_mail"
        'att_cell_phone.SortExpression = "att_cell_phone"
        'eng_floor.SortExpression = "eng_floor"
        'pc_kind1_c.SortExpression = "pc_kind1_c"
		'pc_kind2_c.SortExpression = "pc_kind2_c"

        '欄位寬度(%):要留6%給編輯圖示
        pc_code.ItemStyle.Width = Unit.Percentage(11)
		eng_name.ItemStyle.Width = Unit.Percentage(9)
        eng_code.ItemStyle.Width = Unit.Percentage(7)
        area_name.ItemStyle.Width = Unit.Percentage(8)
        road_name.ItemStyle.Width = Unit.Percentage(8)
        eng_add.ItemStyle.Width = Unit.Percentage(8)
        com_name.ItemStyle.Width = Unit.Percentage(9)
		pc_date.ItemStyle.Width = Unit.Percentage(10)
		pc_date_s.ItemStyle.Width = Unit.Percentage(12)
        pc_date_e.ItemStyle.Width = Unit.Percentage(12)
        att_tel1.ItemStyle.Width = Unit.Percentage(10)
        att_name1.ItemStyle.Width = Unit.Percentage(10)
        att_mail.ItemStyle.Width = Unit.Percentage(10)
        att_cell_phone.ItemStyle.Width = Unit.Percentage(10)
        eng_floor.ItemStyle.Width = Unit.Percentage(10)
        pc_kind1_c.ItemStyle.Width = Unit.Percentage(10)
		pc_kind2_c.ItemStyle.Width = Unit.Percentage(10)

        '設定欄位標題顏色
        pc_code.HeaderStyle.CssClass = "grid1"
		eng_name.HeaderStyle.CssClass = "grid1"
        eng_code.HeaderStyle.CssClass = "grid1"
        area_name.HeaderStyle.CssClass = "grid1"
        road_name.HeaderStyle.CssClass = "grid1"
        eng_add.HeaderStyle.CssClass = "grid1"
        com_name.HeaderStyle.CssClass = "grid1"
		pc_date.HeaderStyle.CssClass = "grid1"
		pc_date_s.HeaderStyle.CssClass = "grid1"
        pc_date_e.HeaderStyle.CssClass = "grid1"
        att_tel1.HeaderStyle.CssClass = "grid1"
        att_name1.HeaderStyle.CssClass = "grid1"
        att_mail.HeaderStyle.CssClass = "grid1"
        att_cell_phone.HeaderStyle.CssClass = "grid1"
        eng_floor.HeaderStyle.CssClass = "grid1"
        pc_kind1_c.HeaderStyle.CssClass = "grid1"
		pc_kind2_c.HeaderStyle.CssClass = "grid1"


        '設定欄位的水平對齊
        pc_code.ItemStyle.HorizontalAlign = HorizontalAlign.Center
		eng_name.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        eng_code.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        area_name.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        road_name.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        eng_add.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        com_name.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        pc_date.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        pc_date_s.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        pc_date_e.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        att_tel1.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        att_name1.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        att_mail.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        att_cell_phone.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        eng_floor.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        pc_kind1_c.ItemStyle.HorizontalAlign = HorizontalAlign.Left
		pc_kind2_c.ItemStyle.HorizontalAlign = HorizontalAlign.Left

        '設定欄位是否顯示(不顯示的再設定)
        'com_name.Visible = False
        pc_code.HtmlEncode = True
        GridView1.Columns.Add(pc_code)
		GridView1.Columns.Add(eng_name)
        GridView1.Columns.Add(eng_code)
        GridView1.Columns.Add(area_name)
        GridView1.Columns.Add(road_name)
        GridView1.Columns.Add(eng_add)
        GridView1.Columns.Add(com_name)
		GridView1.Columns.Add(pc_date)
		GridView1.Columns.Add(pc_date_s)
        GridView1.Columns.Add(pc_date_e)
        GridView1.Columns.Add(att_tel1)
        GridView1.Columns.Add(att_name1)
        GridView1.Columns.Add(att_mail)
        GridView1.Columns.Add(att_cell_phone)
        GridView1.Columns.Add(eng_floor)
        GridView1.Columns.Add(pc_kind1_c)
		GridView1.Columns.Add(pc_kind2_c)

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) Then
            setGridViewStyle()
            setFields()
            setCheckBox()
            setGridViewPage()

            Set_DDL_Option("ZipCode", "", ddl_eng_area, "", "---請選擇---", "B")
            Set_DDL_Option(ddl_eng_area.SelectedValue, "", ddl_eng_road, "", "---請選擇---", "B")
            Set_CKL_Option("pc_kind1", "", ckl_pc_kind1, "", "", "B")
			Set_CKL_Option("pc_kind2", "", ckl_pc_kind2, "", "", "B")
            'Set_CKL_Option("tear_tool1", "", ckl_tear_tool1, "", "", "")
            'Set_CKL_Option("cage_tool1", "", ckl_cage_tool1, "", "", "")
            ''進階搜尋預設值
            'txt_work_date_s_s.Text = Format(Now, "yyyy/MM/") & "01"
            'txt_work_date_s_e.Text = Format(Now, "yyyy/MM/dd")
            'txt_work_date_e_s.Text = Format(Now, "yyyy/MM/") & "01"
            'txt_work_date_e_e.Text = Format(Now, "yyyy/MM/dd")

            '進階搜尋增加點
            WebSubMenu.setName("GridView1", "SqlDS2", "txtOrder", "txtSubWhere")

            '設定Row的鍵值組成，具有唯一性
            Dim KeyNames() As String = {"pc_code"}
            GridView1.DataKeyNames = KeyNames
            ws.InsUsrRec(s_prgname, "SRC", s_prgcode, Session("usr_code"), Session("usr_ip"), "", "", "")
            txtOrder.Text = " ORDER BY pc_code DESC"
            'Response.Write(txtOrder.Text)
            'Response.End()
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
        img_date1.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_spc_date")
        img_date2.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_epc_date")
        img_date3.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_spc_date_s")
        img_date4.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_epc_date_s")

        img_date5.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_spc_date_e")
        img_date6.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_epc_date_e")
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
        'Response.Write(SqlDS1.SelectCommand)
        'Response.End()
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
                '    Row.Cells(3).Text &= "<img src='/images/icon/new2.png' />"
                'End If
            Next
        Else
            For Each Row As GridViewRow In GridView1.Rows
                'If Row.Cells(20).Text >= Format(Now, "yyyy/MM/dd") And Row.Cells(20).Text <= Format(Now.AddDays(3), "yyyy/MM/dd") Then
                '    Row.Cells(3).Text &= " - NEW"
                'End If
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
            'e.Row.Cells(7).Visible = False
            'e.Row.Cells(8).Visible = False
            'e.Row.Cells(9).Visible = False
            'e.Row.Cells(10).Visible = False

            e.Row.Cells(13).Visible = False
            e.Row.Cells(14).Visible = False
            e.Row.Cells(15).Visible = False
            e.Row.Cells(16).Visible = False
            e.Row.Cells(17).Visible = False
            e.Row.Cells(18).Visible = False
            e.Row.Cells(19).Visible = False

            'e.Row.Cells(20).Visible = False
            'e.Row.Cells(21).Visible = False
            'e.Row.Cells(23).Visible = False
            'e.Row.Cells(24).Visible = False
            'e.Row.Cells(26).Visible = False
            'e.Row.Cells(27).Visible = False
            'e.Row.Cells(28).Visible = False
            'e.Row.Cells(29).Visible = False
            'e.Row.Cells(30).Visible = False
            'e.Row.Cells(31).Visible = False
            'e.Row.Cells(32).Visible = False

            '匯出Excel時顯示設定
            If bPrint = True Then
                '隱藏功能列
                If GridView1.Columns(0).Visible = True Then e.Row.Cells(0).Visible = False
                If GridView1.Columns(1).Visible = True Then e.Row.Cells(1).Visible = False
                If GridView1.Columns(2).Visible = True Then e.Row.Cells(2).Visible = False

                '選擇全部
                If ckb_all.Checked Then
                    e.Row.Cells(3).Visible = True
                    e.Row.Cells(4).Visible = True
                    e.Row.Cells(5).Visible = True
                    e.Row.Cells(6).Visible = True
                    e.Row.Cells(7).Visible = True
                    e.Row.Cells(8).Visible = True
                    e.Row.Cells(9).Visible = True
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

                    '設定欄位標頭名稱
                    If e.Row.RowType = DataControlRowType.Header Then
                        e.Row.Cells(3).Text = "資料編號"
                        e.Row.Cells(4).Text = "工程名稱"
						e.Row.Cells(5).Text = "建號"
                        e.Row.Cells(6).Text = "工程地址-區"
                        e.Row.Cells(7).Text = "工程地址-路段"
                        e.Row.Cells(8).Text = "工程地址"
                        e.Row.Cells(9).Text = "通報單位"
                        e.Row.Cells(10).Text = "通報日期"
                        e.Row.Cells(11).Text = "開工日期"
                        e.Row.Cells(12).Text = "完工日期"
                        e.Row.Cells(13).Text = "聯絡電話"
                        e.Row.Cells(14).Text = "聯絡人"                      
                        e.Row.Cells(15).Text = "電子信箱"
                        e.Row.Cells(16).Text = "手機"
                        e.Row.Cells(17).Text = "作業區域及樓層"
                        e.Row.Cells(18).Text = "丁類危險性工作場所主要危害作業"
						e.Row.Cells(19).Text = "機械作業"

                    End If
                Else
                    '隱藏未選擇欄位
                    For i As Integer = 1 To CBL_Field1.Items.Count - 1
                        If CBL_Field1.Items(i).Selected = False Then
                            e.Row.Cells(CBL_Field1.Items(i).Value).Visible = False
                        End If
                    Next

                    '顯示選擇欄位
                    For i As Integer = 1 To CBL_Field1.Items.Count - 1
                        If CBL_Field1.Items(i).Selected = True Then
                            e.Row.Cells(CBL_Field1.Items(i).Value).Visible = True
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
				
				Dim nSpilt1 = Split(Trim(e.Row.Cells(18).Text & ""), "|")
				Dim nSpilt2 = Split(Trim(e.Row.Cells(19).Text & ""), "|")
				If  bPrint  =  True  Then
					For i As Integer = 0 To ckl_pc_kind1.Items.Count - 1
					    For j As Integer = 0 To nSpilt1.Length -1
						   If nSpilt1(j) = ckl_pc_kind1.Items(i).Value Then nSpilt1(j) = ckl_pc_kind1.Items(i).Text
					    Next
                    Next
					e.Row.Cells(18).Text = ""
		            For i As Integer = 0 To nSpilt1.Length -1
					    If e.Row.Cells(18).Text <> "" Then e.Row.Cells(18).Text &= "|"
						e.Row.Cells(18).Text &= nSpilt1(i)
					Next
		            For i As Integer = 0 To ckl_pc_kind2.Items.Count - 1
                        For j As Integer = 0 To nSpilt2.Length -1
						   If nSpilt2(j) = ckl_pc_kind2.Items(i).Value Then nSpilt2(j) = ckl_pc_kind2.Items(i).Text
					    Next
                    Next
                    e.Row.Cells(19).Text = ""
		            For i As Integer = 0 To nSpilt2.Length -1
					    If e.Row.Cells(19).Text <> "" Then e.Row.Cells(19).Text &= "|"
						e.Row.Cells(19).Text &= nSpilt2(i)
					Next					
                End If
        End Select

        Dim is_ok As Boolean
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            is_ok = Chk_UsrLimit(Session("usr_code"), e.Row.Cells(4).Text, Session("grp_code"))
            e.Row.FindControl("ImageButton2").Visible = is_ok
            e.Row.FindControl("ImageButton3").Visible = is_ok
        End If

        '控制截止日期的資料修改權限

        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                Dim sEndDate As String = ws.Get_SysPara("LSI026_END_DATE", "")
                If Trim(e.Row.Cells(10).Text & "") < sEndDate Then
                    'CType(e.Row.Cells(0).Controls(1), ImageButton).Visible = False
                    e.Row.FindControl("ImageButton2").Visible = False
                    e.Row.FindControl("ImageButton3").Visible = False
                End If
        End Select



        'ViewState("LineNo") = 1
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '資料搜尋-單一條件
        If TxtFind.Text <> "" Then
            txtSubWhere.Text = " WHERE " & ddlField.SelectedValue & " LIKE N'%" & RepSql(TxtFind.Text) & "%'"
        Else
            txtSubWhere.Text = " WHERE pc_date LIKE N'%" & Year(Now) & "%'"
        End If
        'Response.Write(txtSubWhere.Text)
        'Response.End()
        setGridViewData()
        GridView1.DataBind()
    End Sub

    Protected Sub Bt_3Days_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Bt_3Days.Click
        '資料搜尋-單一條件
        If TxtFind.Text <> "" Then
            txtSubWhere.Text = " WHERE " & ddlField.SelectedValue & " LIKE N'%" & RepSql(TxtFind.Text) & "%'" & _
                               " AND work_date_e BETWEEN '" & Format(Now, "yyyy/MM/dd") & "' AND '" & Format(Now.AddDays(3), "yyyy/MM/dd") & "'"
        Else
            txtSubWhere.Text = " WHERE work_date_e BETWEEN '" & Format(Now, "yyyy/MM/dd") & "' AND '" & Format(Now.AddDays(3), "yyyy/MM/dd") & "'"
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
        '工程名稱
        txtSubWhere.Text = " WHERE " & "eng_name" & " LIKE N'%" & RepSql(txt_eng_name.Text) & "%'"
        '工程地址-區
        If ddl_eng_area.SelectedValue <> "" Then
            txtSubWhere.Text &= " AND " & "eng_area" & " = N'" & RepSql(ddl_eng_area.SelectedValue) & "'"
        End If

        '工程地址-路段
        If ddl_eng_road.SelectedValue <> "" Then
            txtSubWhere.Text &= " AND " & "eng_road" & " = N'" & RepSql(ddl_eng_road.SelectedValue) & "'"
        End If

        '通報單位
        If RepSql(txt_com_name.Text) <> "" Then
            txtSubWhere.Text &= " AND " & "com_name" & " LIKE N'%" & RepSql(txt_com_name.Text) & "%'"
        End If

        '通報日期
        If RepSql(txt_spc_date.Text) <> "" And RepSql(txt_epc_date.Text) <> "" Then
            txtSubWhere.Text &= " AND " & "pc_date" & " BETWEEN N'" & RepSql(txt_spc_date.Text) & "' AND '" & RepSql(txt_epc_date.Text) & "'"
        ElseIf RepSql(txt_spc_date.Text) <> "" Then
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('請輸入 [通報日期] 的結束日期條件');</script>"))
            Exit Sub
        ElseIf RepSql(txt_epc_date.Text) <> "" Then
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('請輸入 [通報日期] 的開始日期條件');</script>"))
            Exit Sub
        End If
        '預計開工日
        If RepSql(txt_spc_date_s.Text) <> "" And RepSql(txt_epc_date_s.Text) <> "" Then
            txtSubWhere.Text &= " AND " & "pc_date_s" & " BETWEEN N'" & RepSql(txt_spc_date_s.Text) & "' AND '" & RepSql(txt_epc_date_s.Text) & "'"
        ElseIf RepSql(txt_spc_date_s.Text) <> "" Then
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('請輸入 [預計開工日] 的結束日期條件');</script>"))
            Exit Sub
        ElseIf RepSql(txt_epc_date_s.Text) <> "" Then
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('請輸入 [預計開工日] 的開始日期條件');</script>"))
            Exit Sub
        End If

        '預計完工日
        If RepSql(txt_spc_date_e.Text) <> "" And RepSql(txt_epc_date_e.Text) <> "" Then
            txtSubWhere.Text &= " AND " & "pc_date_e" & " BETWEEN N'" & RepSql(txt_spc_date_e.Text) & "' AND '" & RepSql(txt_epc_date_e.Text) & "'"
        ElseIf RepSql(txt_spc_date_e.Text) <> "" Then
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('請輸入 [預計完工日] 的結束日期條件');</script>"))
            Exit Sub
        ElseIf RepSql(txt_epc_date_e.Text) <> "" Then
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('請輸入 [預計完工日] 的開始日期條件');</script>"))
            Exit Sub
        End If

        '種類
        Dim bPcKindChecked As Boolean = False
        For i As Integer = 0 To ckl_pc_kind1.Items.Count - 1
            If ckl_pc_kind1.Items(i).Selected = True Then
                If bPcKindChecked = False Then
                    txtSubWhere.Text &= " AND ("
                    bPcKindChecked = True
                Else
                    txtSubWhere.Text &= " OR "
                End If
                txtSubWhere.Text &= "LSI026.pc_kind1" & " LIKE N'%" & RepSql(ckl_pc_kind1.Items(i).Value) & "%'"
            End If
        Next
		
		For i As Integer = 0 To ckl_pc_kind2.Items.Count - 1
            If ckl_pc_kind2.Items(i).Selected = True Then
                If bPcKindChecked = False Then
                    txtSubWhere.Text &= " AND ("
                    bPcKindChecked = True
                Else
                    txtSubWhere.Text &= " OR "
                End If
                txtSubWhere.Text &= "LSI026.pc_kind2" & " LIKE N'%" & RepSql(ckl_pc_kind2.Items(i).Value) & "%'"
            End If
        Next
		
        If bPcKindChecked = True Then txtSubWhere.Text &= ")"       

        setGridViewData()
        GridView1.DataBind()
    End Sub

    Protected Sub Bt_ExSearch_no_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Bt_ExSearch_no.Click
        Panel_ExSearch.Visible = False
    End Sub

    Protected Sub ddl_eng_area_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddl_eng_area.SelectedIndexChanged
        '    txtSubWhere.Text = 0
        ddl_eng_road.Items.Clear()
        Set_DDL_Option(ddl_eng_area.SelectedValue, "", ddl_eng_road, "", "---請選擇---", "B")
    End Sub
    Protected Sub ImageButton1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton1.Click
        'Server.Transfer("guidance.aspx")
    End Sub
End Class
