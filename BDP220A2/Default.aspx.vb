Imports System.Drawing
Imports System.Web.Configuration

Partial Class management_BDP220A_Default
    Inherits PageBase

    Dim s_prgcode As String
    Dim s_prgname As String
    Dim UserLimit1 As String

    '程式自定區 以下

    Protected Sub setFields()
        '建立資料繫結欄位
        Dim bdp220 As New BoundField()
        Dim prg_code As New BoundField()
        Dim field_code As New BoundField()
        Dim field_name As New BoundField()
        Dim data_source As New BoundField()
        Dim data_type As New BoundField()
        Dim scr_no As New BoundField()
        Dim is_use As New BoundField()

        '指定資料來源欄位
        bdp220.DataField = "bdp220"
        prg_code.DataField = "prg_code"
        field_code.DataField = "field_code"
        field_name.DataField = "field_name"
        data_source.DataField = "data_source"
        data_type.DataField = "data_type"
        scr_no.DataField = "scr_no"
        is_use.DataField = "is_use"

        '設定欄位標頭名稱
        prg_code.HeaderText = "程式編號"
        field_code.HeaderText = "欄位代碼"
        field_name.HeaderText = "欄位名稱"
        data_source.HeaderText = "資料來源"
        data_type.HeaderText = "資料類型"
        scr_no.HeaderText = "順序"
        is_use.HeaderText = "是否使用"

        '設定欄位是否自動換行
        'prg_code.ItemStyle.Wrap = False
        'field_code.ItemStyle.Wrap = False
        'is_use.ItemStyle.Wrap = False

        '設定排序所對應的欄位
        prg_code.SortExpression = "prg_code"
        field_code.SortExpression = "field_code"
        field_name.SortExpression = "field_name"
        data_source.SortExpression = "data_source"
        data_type.SortExpression = "data_type"
        scr_no.SortExpression = "scr_no"
        is_use.SortExpression = "is_use"

        '欄位寬度(%):要留6%給編輯圖示
        prg_code.ItemStyle.Width = Unit.Percentage(15)
        field_code.ItemStyle.Width = Unit.Percentage(15)
        field_name.ItemStyle.Width = Unit.Percentage(20)
        data_source.ItemStyle.Width = Unit.Percentage(20)
        'data_type.ItemStyle.Width = Unit.Percentage(20)
        scr_no.ItemStyle.Width = Unit.Percentage(8)
        is_use.ItemStyle.Width = Unit.Percentage(8)

        '設定欄位標題顏色
        prg_code.HeaderStyle.CssClass = "grid1"
        field_code.HeaderStyle.CssClass = "grid1"
        field_name.HeaderStyle.CssClass = "grid1"
        data_source.HeaderStyle.CssClass = "grid1"
        data_type.HeaderStyle.CssClass = "grid1"
        scr_no.HeaderStyle.CssClass = "grid1"
        is_use.HeaderStyle.CssClass = "grid1"

        '設定欄位的水平對齊
        prg_code.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        field_code.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        field_name.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        data_source.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        data_type.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        scr_no.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        is_use.ItemStyle.HorizontalAlign = HorizontalAlign.Center

        '設定欄位是否顯示(不顯示的再設定)
        bdp220.Visible = False

        GridView1.Columns.Add(bdp220)
        GridView1.Columns.Add(prg_code)
        GridView1.Columns.Add(field_code)
        GridView1.Columns.Add(field_name)
        GridView1.Columns.Add(data_source)
        GridView1.Columns.Add(scr_no)
        GridView1.Columns.Add(is_use)

    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        '程式變數設定
        s_prgcode = "BDP220A"
        s_prgname = Get_PrgName(s_prgcode)
        UserLimit1 = Get_limit(Session("usr_code"), s_prgcode)
        Call setLimit(UserLimit1)

        If (Not IsPostBack) Then
            setGridViewStyle()
            setFields()
            setGridViewPage()

            '設定Row的鍵值組成，具有唯一性
            Dim KeyNames() As String = {"bdp220"}
            GridView1.DataKeyNames = KeyNames
            Me.InsUsrRec(s_prgname, "SRC", s_prgcode)
        End If

        ''建立連線物件
        'Con_db = Me.ConnectDB()

        '設定查詢欄位
        SqlDS2.ConnectionString = WebConfigurationManager.ConnectionStrings("con_db").ConnectionString
        SqlDS2.SelectCommand = "select * from BDP140 where prg_code='" & s_prgcode & "' order by scr_no"

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
        SqlDS1.ConnectionString = WebConfigurationManager.ConnectionStrings("con_db").ConnectionString
        If txtSubWhere.Text <> "" Then
            Dim Tmp_str As String

            Tmp_str = txtSubWhere.Text
            If InStr(Left(Tmp_str, 6), "where") > 0 Then
                SqlDS1.SelectCommand = "select * from BDP220 " & txtSubWhere.Text & " order by prg_code,scr_no"
            Else
                SqlDS1.SelectCommand = "select * from BDP220 where 1=1 " & txtSubWhere.Text & " order by prg_code,scr_no"
            End If
        Else
            SqlDS1.SelectCommand = "select * from BDP220 where 1=0"
        End If

        '設定Delete命令
        Dim CtrPara As New ControlParameter("datakeys", "GridView1", "SelectedValue")
        SqlDS1.DeleteParameters.Add(CtrPara)
        SqlDS1.DeleteCommand = "delete from BDP220 where bdp220=@datakeys"

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
        'GridView1.Columns(2).Visible = False
        Button1.Enabled = CBool(InStr(Climit, "S") > 0)
        Button4.Enabled = CBool(InStr(Climit, "A") > 0)
    End Sub
    Protected Sub setGridViewPage()
        '設定GridView的分頁樣式
        GridView1.PagerSettings.Mode = PagerButtons.NextPreviousFirstLast
        GridView1.PagerSettings.FirstPageImageUrl = "~/Images/First.gif"
        GridView1.PagerSettings.FirstPageText = "第一頁"
        GridView1.PagerSettings.PreviousPageImageUrl = "~/Images/Previous.gif"
        GridView1.PagerSettings.PreviousPageText = "上一頁"
        GridView1.PagerSettings.NextPageImageUrl = "~/Images/Next.gif"
        GridView1.PagerSettings.NextPageText = "下一頁"
        GridView1.PagerSettings.LastPageImageUrl = "~/Images/Last.gif"
        GridView1.PagerSettings.LastPageText = "最後頁"
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
        Dim dt_tmp As Object
        dt_tmp = Get_DataTable(SqlDS1.SelectCommand)
        If dt_tmp Is Nothing Then
            Label2.Text = "總筆數：0/總頁數:" & GridView1.PageCount.ToString() & "/每頁筆數:" & GridView1.PageSize.ToString()
        Else
            Label2.Text = "總筆數：" & CStr(dt_tmp.rows.count) & "/總頁數:" & GridView1.PageCount.ToString() & "/每頁筆數:" & GridView1.PageSize.ToString()

            '有資料才秀
            If CInt(dt_tmp.rows.count) > 0 Then
                Dim bottonPagerRow As GridViewRow = GridView1.BottomPagerRow
                Dim bottonPagerNo As New Label()
                If Not bottonPagerRow.Cells(0) Is Nothing Then
                    bottonPagerNo.Text = "目前所在分頁碼（" & (GridView1.PageIndex + 1) & "/" & GridView1.PageCount.ToString() & ")"
                    bottonPagerRow.Cells(0).Controls.Add(bottonPagerNo)
                End If
            End If
        End If
    End Sub

    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        GridView1.SelectedIndex = e.RowIndex
        Dim txtDataKey As String = GridView1.SelectedDataKey.Value.ToString
        If Not txtDataKey = "" Then
            '設定記錄資料
            Dim Sql_str As String = "select * from BDP220 where bdp220='" & txtDataKey & "'"
            Dim sFieldName As String = Get_FieldName(s_prgcode)
            Dim sOldData As String = Get_OldData(Sql_str, s_prgcode)
            Dim sNewData As String = Get_NewData(s_prgcode, "DEL")

            SqlDS1.Delete()
            InsUsrRec(txtDataKey, "DEL", s_prgcode, sFieldName, sOldData, sNewData)
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
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles Button1.Click
        '資料搜尋-單一條件
        If Trim(Replace(TxtFind.Text, "'", "") & "") <> "" Then
            txtSubWhere.Text = " where " & ddlField.SelectedValue & " like '%" & TxtFind.Text & "%'"
        Else
            txtSubWhere.Text = " where 1=1"
        End If
        setGridViewData()
    End Sub
    Protected Sub Button4_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles Button4.Click
        txtPgmType.Text = "ADD"
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

    '程式固定區-不要修改 以上

End Class
