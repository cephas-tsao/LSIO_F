Imports System.Drawing
Imports System.Web.Configuration

Partial Class management_BDP003_A_Default
    Inherits PageBase

    Dim s_prgcode As String
    Dim s_prgname As String
    Dim UserLimit1 As String

    '程式自定區 以下

    Protected Sub setFields()
        '建立資料繫結欄位
        'Dim prg_code As New BoundField()
        'Dim prg_name As New BoundField()
        'Dim prg_type As New BoundField()
        'Dim is_use As New BoundField()
        Dim usr_code As New BoundField()
        Dim prg_code As New BoundField()
        Dim usr_ip As New BoundField()
        Dim usr_date As New BoundField()
        Dim usr_time As New BoundField()
        Dim usr_type As New BoundField()
        Dim cmemo As New BoundField()
        Dim bdp200 As New BoundField()

        '指定資料來源欄位
        'prg_code.DataField = "prg_code"
        'prg_name.DataField = "prg_name"
        'prg_type.DataField = "prg_type"
        'is_use.DataField = "is_use"
        bdp200.DataField = "bdp200"
        usr_code.DataField = "usr_code"
        prg_code.DataField = "prg_code"
        usr_ip.DataField = "usr_ip"
        usr_date.DataField = "usr_date"
        usr_time.DataField = "usr_time"
        usr_type.DataField = "usr_type"
        cmemo.DataField = "cmemo"

        '設定欄位標頭名稱
        'prg_code.HeaderText = "程式編號"
        'prg_name.HeaderText = "程式名稱"
        'prg_type.HeaderText = "程式類型"
        'is_use.HeaderText = "是否使用"
        bdp200.HeaderText = "序號"
        usr_code.HeaderText = "使用者ID"
        prg_code.HeaderText = "程式ID"
        usr_ip.HeaderText = "使用者IP"
        usr_date.HeaderText = "使用日期"
        prg_code.HeaderText = "程式編號"
        usr_time.HeaderText = "使用時間"
        usr_type.HeaderText = "類型"
        cmemo.HeaderText = "內容"

        '設定欄位是否自動換行(不換行的再設定)
        'prg_code.ItemStyle.Wrap = False

        '設定排序所對應的欄位
        bdp200.SortExpression = "bdp200"
        usr_code.SortExpression = "usr_code"
        prg_code.SortExpression = "prg_code"
        usr_ip.SortExpression = "usr_ip"
        usr_date.SortExpression = "usr_date"
        prg_code.SortExpression = "prg_code"
        usr_time.SortExpression = "usr_time"
        usr_type.SortExpression = "usr_type"
        cmemo.SortExpression = "cmemo"

        '欄位寬度(%):要留6%給編輯圖示
        bdp200.ItemStyle.Width = Unit.Percentage(10)
        usr_code.ItemStyle.Width = Unit.Percentage(10)
        prg_code.ItemStyle.Width = Unit.Percentage(10)
        usr_ip.ItemStyle.Width = Unit.Percentage(10)
        usr_date.ItemStyle.Width = Unit.Percentage(10)
        prg_code.ItemStyle.Width = Unit.Percentage(10)
        usr_time.ItemStyle.Width = Unit.Percentage(10)
        usr_type.ItemStyle.Width = Unit.Percentage(10)
        cmemo.ItemStyle.Width = Unit.Percentage(14)

        '設定欄位的水平對齊
        'prg_code.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        'prg_name.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        'prg_type.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        'is_use.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        bdp200.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        usr_code.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        prg_code.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        usr_ip.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        usr_date.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        prg_code.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        usr_time.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        usr_type.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        cmemo.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        

        '設定欄位是否顯示(不顯示的再設定)
        'dep_code.Visible = False
        bdp200.Visible = False
        'GridView1.Columns.Add(prg_code)
        'GridView1.Columns.Add(prg_name)
        'GridView1.Columns.Add(prg_type)
        'GridView1.Columns.Add(is_use)
        GridView1.Columns.Add(bdp200)
        GridView1.Columns.Add(usr_code)
        GridView1.Columns.Add(prg_code)
        GridView1.Columns.Add(usr_ip)
        GridView1.Columns.Add(usr_date)
        GridView1.Columns.Add(usr_time)
        GridView1.Columns.Add(usr_type)
        GridView1.Columns.Add(prg_code)
        GridView1.Columns.Add(cmemo)
        
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) Then
            setGridViewStyle()
            setFields()
            setGridViewPage()

            '設定Row的鍵值組成，具有唯一性
            Dim KeyNames() As String = {"bdp200"}
            GridView1.DataKeyNames = KeyNames
        End If
        Button4.Visible = False
        ''建立連線物件
        'Con_db = Me.ConnectDB()

        '程式變數設定
        s_prgcode = "BDP020A"
        s_prgname = Get_PrgName(s_prgcode)
        UserLimit1 = Get_limit(Session("usr_code"), s_prgcode)
        Call setLimit(UserLimit1)

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
                SqlDS1.SelectCommand = "select * from BDP200 " & txtSubWhere.Text
            Else
                SqlDS1.SelectCommand = "select * from BDP200 where 1=1 " & txtSubWhere.Text
            End If
        Else
            SqlDS1.SelectCommand = "select * from BDP200 where 1=0"
        End If

        '設定Delete命令
        Dim CtrPara As New ControlParameter("datakeys", "GridView1", "SelectedValue")
        SqlDS1.DeleteParameters.Add(CtrPara)
        SqlDS1.DeleteCommand = "delete from BDP200 where bdp200=@datakeys"

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
        Select Case e.Row.RowType
            Case DataControlRowType.Header
                e.Row.BackColor = Color.FromArgb(153, 0, 0)
                e.Row.ForeColor = Color.White
            Case DataControlRowType.DataRow
                '建立奇數資料列與偶數資料列的onmouseover及onmouseout的顏色變換
                If (CType(ViewState("LineNo"), Int16) = 0) Then
                    'e.Row.BackColor = Color.FromArgb(255, 251, 214)
                    e.Row.Attributes.Add("onmouseout", "this.style.backgroundColor='#FFFBD6';this.style.color='black'")
                    e.Row.Attributes.Add("onmouseover", "this.style.backgroundColor='#C0C0FF';this.style.color='#ffffff'")

                    ViewState("LineNo") = 1
                Else
                    'e.Row.BackColor = Color.White
                    e.Row.Attributes.Add("onmouseout", "this.style.backgroundColor='#FFFFFF';this.style.color='black'")
                    e.Row.Attributes.Add("onmouseover", "this.style.backgroundColor='#C0C0FF';this.style.color='#ffffff'")

                    ViewState("LineNo") = 0
                End If
        End Select
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles Button1.Click
        '資料搜尋-單一條件
        If Trim(Replace(TxtFind.Text, "'", "") & "") <> "" Then
            txtSubWhere.Text = " where (usr_type='ADD' OR usr_type='DEL' OR usr_type='MDY') and " & ddlField.SelectedValue & " like '%" & TxtFind.Text & "%'"
            'txtSubWhere.Text = " where " & ddlField.SelectedValue & " like '%" & TxtFind.Text & "%'"
        Else
            txtSubWhere.Text = " where (usr_type='ADD' OR usr_type='DEL' OR usr_type='MDY')"
            'txtSubWhere.Text = " where 1=1"
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

