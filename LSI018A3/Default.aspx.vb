Imports System.Data
Imports System.Drawing

Partial Class basic_LSI018A_Default
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI018A"
    Dim s_prgname As String = ws.Get_PrgName(s_prgcode)
    Dim UserLimit1 As String

    '程式自定區 以下

    Protected Sub setFields()
        '建立資料繫結欄位
        Dim apl_code As New BoundField()
        Dim apl_type As New BoundField()
        Dim apl_date As New BoundField()
        Dim apl_time As New BoundField()
        Dim app_type As New BoundField()
        Dim usr_code As New BoundField()

        '指定資料來源欄位
        apl_code.DataField = "apl_code"
        apl_type.DataField = "apl_type"
        apl_date.DataField = "apl_date"
        apl_time.DataField = "apl_time"
        app_type.DataField = "app_type"
        usr_code.DataField = "usr_code"

        '設定欄位標頭名稱
        apl_type.HeaderText = "操作類別"
        apl_date.HeaderText = "操作日期"
        apl_time.HeaderText = "操作時間"
        app_type.HeaderText = "APP種類"
        usr_code.HeaderText = "使用者ID"

        '設定欄位是否自動換行
        'apl_type.ItemStyle.Wrap = False
        'apl_date.ItemStyle.Wrap = False
        'is_use.ItemStyle.Wrap = False

        '設定排序所對應的欄位
        apl_type.SortExpression = "apl_type"
        apl_date.SortExpression = "apl_date"
        apl_time.SortExpression = "apl_time"
        app_type.SortExpression = "app_type"
        usr_code.SortExpression = "usr_code"

        '欄位寬度(%):要留6%給編輯圖示
        apl_type.ItemStyle.Width = Unit.Percentage(15)
        apl_date.ItemStyle.Width = Unit.Percentage(15)
        apl_time.ItemStyle.Width = Unit.Percentage(20)
        app_type.ItemStyle.Width = Unit.Percentage(20)
        usr_code.ItemStyle.Width = Unit.Percentage(20)

        '設定欄位標題顏色
        apl_type.HeaderStyle.CssClass = "grid1"
        apl_date.HeaderStyle.CssClass = "grid1"
        apl_time.HeaderStyle.CssClass = "grid1"
        app_type.HeaderStyle.CssClass = "grid1"
        usr_code.HeaderStyle.CssClass = "grid1"

        '設定欄位的水平對齊
        apl_type.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        apl_date.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        apl_time.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        app_type.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        usr_code.ItemStyle.HorizontalAlign = HorizontalAlign.Left


        '設定欄位是否顯示(不顯示的再設定)
        apl_code.Visible = False

        'GridView1.Columns.Add(apl_code)
        GridView1.Columns.Add(apl_type)
        GridView1.Columns.Add(apl_date)
        GridView1.Columns.Add(apl_time)
        GridView1.Columns.Add(app_type)
        GridView1.Columns.Add(usr_code)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) Then
            setGridViewStyle()
            setFields()
            setGridViewPage()

            '設定Row的鍵值組成，具有唯一性
            Dim KeyNames() As String = {"apl_code"}
            GridView1.DataKeyNames = KeyNames
        End If

        '程式變數設定
        UserLimit1 = ws.Get_limit(Session("usr_code"), s_prgcode)
        Call setLimit(UserLimit1)

        '設定查詢欄位
        SqlDS2.ConnectionString = ws.Get_ConnStr("con_db")
        SqlDS2.SelectCommand = Get_SqlStr(s_prgcode, "ddlField")

        '程式抬頭
        Label1.Text = s_prgname & "(" & s_prgcode & ")"

        '取得預設條件
        If Trim(Request("subwhere") & "") <> "" Then txtSubWhere.Text = Request("subwhere")

        '取得資料
        setGridViewData()
    End Sub
    Protected Sub setGridViewData()
        '取得GridView資料

        '設定SqlDataSource連線及Select命令
        SqlDS1.ConnectionString = ws.Get_ConnStr("con_db")
        '進階搜尋增加點
        If txtSubWhere.Text <> "" Then
            SqlDS1.SelectCommand = Get_SqlStr(s_prgcode, "SELECT", txtSubWhere.Text)
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
        GridView1.Columns(1).Visible = CBool(InStr(Climit, "S") > 0)
        GridView1.Columns(2).Visible = CBool(InStr(Climit, "A") > 0)
        GridView1.Columns(2).Visible = False
        Button1.Enabled = CBool(InStr(Climit, "S") > 0)
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
        GridView1.PageSize = CInt(Get_SysPara("pagecount", "20"))
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

            ''有資料才秀
            'If CInt(dt_tmp.rows.count) > 0 Then
            '    Dim bottonPagerRow As GridViewRow = GridView1.BottomPagerRow
            '    Dim bottonPagerNo As New Label()
            '    If Not bottonPagerRow.Cells(0) Is Nothing Then
            '        bottonPagerNo.Text = "目前所在分頁碼（" & (GridView1.PageIndex + 1) & "/" & GridView1.PageCount.ToString() & ")"
            '        bottonPagerRow.Cells(0).Controls.Add(bottonPagerNo)
            '    End If
            'End If

            '設定分頁訊息
            l_PageMsg.Visible = False
            If GridView1.PageCount > 1 Then
                l_PageMsg.Visible = True
                l_PageMsg.Text = "　目前所在分頁碼 (" & (GridView1.PageIndex + 1) & "/" & GridView1.PageCount.ToString() & ")"
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
        End If
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        GridView1.SelectedIndex = e.RowIndex
        Dim txtDataKey As String = GridView1.SelectedDataKey.Value.ToString
        If Not txtDataKey = "" Then
            SqlDS1.Delete()
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

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.Header Or e.Row.RowType = DataControlRowType.DataRow Then
            e.Row.Cells(1).Visible = False
            'e.Row.Cells(11).Visible = False
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

                '加入onclick事件
                'If e.Row.Cells(11).Text <> "&nbsp;" Then
                '    e.Row.Cells(1).Attributes.Add("onclick", "event.returnValue=false;RecordDetail('" & e.Row.Cells(3).Text & "')")
                'Else
                '    e.Row.Cells(1).Text = ""
                'End If
        End Select
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '資料搜尋-單一條件
        If TxtFind.Text <> "" Then
            'txtSubWhere.Text = " WHERE (usr_type='ADD' OR usr_type='DEL' OR usr_type='MDY' OR usr_type='COPY' OR usr_type='SRC' or usr_type='Mail' or usr_type='UPF') AND " & ddlField.SelectedValue & " LIKE '%" & RepSql(TxtFind.Text) & "%'"
            txtSubWhere.Text = " WHERE 1=1 " & " AND " & ddlField.SelectedValue & " LIKE '%" & RepSql(TxtFind.Text) & "%'"
        Else
            'txtSubWhere.Text = " WHERE (usr_type='ADD' OR usr_type='DEL' OR usr_type='MDY' OR usr_type='COPY' OR usr_type='SRC' or usr_type='Mail' or usr_type='UPF')"
            txtSubWhere.Text = " WHERE 1=1 "
        End If
        setGridViewData()
        GridView1.DataBind()
    End Sub

    Protected Sub GridView1_SelectedIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSelectEventArgs) Handles GridView1.SelectedIndexChanging
        GridView1.SelectedIndex = e.NewSelectedIndex
        If e.NewSelectedIndex >= 0 Then
            Dim txtDataKey As String = GridView1.SelectedDataKey.Value.ToString
            txtPrimary.Text = txtDataKey
        End If
    End Sub

    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        txtPgmType.Text = e.CommandArgument
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

End Class

