Imports System.Data
Imports System.Drawing

Partial Class public_addForm
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "addForm"
    '程式自定區 以下

    Protected Sub setFields()
        '建立資料繫結欄位
        Dim prg_code As New BoundField()
        Dim prg_name As New BoundField()
        Dim prg_type As New BoundField()
        Dim is_use As New BoundField()

        Dim Bound0 As New BoundField()
        Dim Bound1 As New BoundField()
        Dim Bound2 As New BoundField()
        Dim Bound3 As New BoundField()
        Dim Bound4 As New BoundField()
        Dim Bound5 As New BoundField()
        Dim Bound6 As New BoundField()
        Dim Bound7 As New BoundField()

        Dim dt_tmp1 As DataTable = Get_DataSet(s_prgcode, "GetField", Request("add_code")).Tables(0)

        For Ilop As Integer = 0 To dt_tmp1.Rows.Count - 1
            Select Case Ilop
                Case 0
                    Bound0.DataField = Trim(dt_tmp1.Rows(Ilop).Item("field_code") & "")
                    Bound0.HeaderText = Trim(dt_tmp1.Rows(Ilop).Item("field_name") & "")
                    Bound0.SortExpression = Trim(dt_tmp1.Rows(Ilop).Item("field_code") & "")
                Case 1
                    Bound1.DataField = Trim(dt_tmp1.Rows(Ilop).Item("field_code") & "")
                    Bound1.HeaderText = Trim(dt_tmp1.Rows(Ilop).Item("field_name") & "")
                    Bound1.SortExpression = Trim(dt_tmp1.Rows(Ilop).Item("field_code") & "")
                Case 2
                    Bound2.DataField = Trim(dt_tmp1.Rows(Ilop).Item("field_code") & "")
                    Bound2.HeaderText = Trim(dt_tmp1.Rows(Ilop).Item("field_name") & "")
                    Bound2.SortExpression = Trim(dt_tmp1.Rows(Ilop).Item("field_code") & "")
                Case 3
                    Bound3.DataField = Trim(dt_tmp1.Rows(Ilop).Item("field_code") & "")
                    Bound3.HeaderText = Trim(dt_tmp1.Rows(Ilop).Item("field_name") & "")
                    Bound3.SortExpression = Trim(dt_tmp1.Rows(Ilop).Item("field_code") & "")
                Case 4
                    Bound4.DataField = Trim(dt_tmp1.Rows(Ilop).Item("field_code") & "")
                    Bound4.HeaderText = Trim(dt_tmp1.Rows(Ilop).Item("field_name") & "")
                    Bound4.SortExpression = Trim(dt_tmp1.Rows(Ilop).Item("field_code") & "")
                Case 5
                    Bound5.DataField = Trim(dt_tmp1.Rows(Ilop).Item("field_code") & "")
                    Bound5.HeaderText = Trim(dt_tmp1.Rows(Ilop).Item("field_name") & "")
                    Bound5.SortExpression = Trim(dt_tmp1.Rows(Ilop).Item("field_code") & "")
                Case 6
                    Bound6.DataField = Trim(dt_tmp1.Rows(Ilop).Item("field_code") & "")
                    Bound6.HeaderText = Trim(dt_tmp1.Rows(Ilop).Item("field_name") & "")
                    Bound6.SortExpression = Trim(dt_tmp1.Rows(Ilop).Item("field_code") & "")
                Case 7
                    Bound7.DataField = Trim(dt_tmp1.Rows(Ilop).Item("field_code") & "")
                    Bound7.HeaderText = Trim(dt_tmp1.Rows(Ilop).Item("field_name") & "")
                    Bound7.SortExpression = Trim(dt_tmp1.Rows(Ilop).Item("field_code") & "")
            End Select
        Next

        If Bound0.HeaderText <> "" Then GridView1.Columns.Add(Bound0)
        If Bound1.HeaderText <> "" Then GridView1.Columns.Add(Bound1)
        If Bound2.HeaderText <> "" Then GridView1.Columns.Add(Bound2)
        If Bound3.HeaderText <> "" Then GridView1.Columns.Add(Bound3)
        If Bound4.HeaderText <> "" Then GridView1.Columns.Add(Bound4)
        If Bound5.HeaderText <> "" Then GridView1.Columns.Add(Bound5)
        If Bound6.HeaderText <> "" Then GridView1.Columns.Add(Bound6)
        If Bound7.HeaderText <> "" Then GridView1.Columns.Add(Bound7)
    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) Then
            setGridViewStyle()
            setFields()
            setGridViewPage()
        End If

        '取得變數資料
        AddCode.Text = Request("add_code")
        FieldList.Text = Request("fieldlist")
        TextBoxList.Text = Request("textboxlist")

        'If Request.Browser.Browser = "Firefox" Then
        TextBoxList.Text = Replace(TextBoxList.Text, "$", "_")
        'End If

        '設定查詢欄位
        SqlDS2.ConnectionString = ws.Get_ConnStr("con_db")
        SqlDS2.SelectCommand = Get_SqlStr(s_prgcode, "", Request("add_code"))

        If Trim(Request("subwhere") & "") <> "" Then txtSubWhere2.Text = Replace(Request("subwhere"), "@", "'")

        '取得資料
        'If Not IsPostBack Then setGridViewData()
        setGridViewData()
    End Sub
    Protected Sub setGridViewData()
        '取得GridView資料
        Dim order_str As String = ""
        Dim dt_tmp1 As DataTable = Get_DataSet(s_prgcode, "GetGridView", Request("add_code")).Tables(0)

        If String.IsNullOrEmpty(dt_tmp1.rows(0).item("order_str")) = False Then order_str = " ORDER BY " & dt_tmp1.rows(0).item("order_str")

        '設定SqlDataSource連線及Select命令
        SqlDS1.ConnectionString = ws.Get_ConnStr("con_db")


        If txtSubWhere.Text <> "" Then
            If txtSubWhere2.Text = "" Then
                If InStr(Trim(dt_tmp1.rows(0).item("sel_str") & ""), "WHERE") > 0 Then
                    'Response.Write("A")
                    SqlDS1.SelectCommand = Trim(dt_tmp1.rows(0).item("sel_str") & "") & Replace(txtSubWhere.Text, " WHERE ", " AND ") & order_str
                Else
                    'Response.Write("B")
                    SqlDS1.SelectCommand = Trim(dt_tmp1.rows(0).item("sel_str") & "") & txtSubWhere.Text & order_str
                End If
            Else
                If InStr(Trim(dt_tmp1.rows(0).item("sel_str") & ""), "WHERE") > 0 Then
                    SqlDS1.SelectCommand = Trim(dt_tmp1.rows(0).item("sel_str") & "") & Replace(txtSubWhere.Text, " WHERE ", " AND ") & " AND " & txtSubWhere2.Text & order_str
                Else
                    SqlDS1.SelectCommand = Trim(dt_tmp1.rows(0).item("sel_str") & "") & txtSubWhere.Text & " AND " & txtSubWhere2.Text & order_str
                End If
            End If
        Else
            If txtSubWhere2.Text = "" Then
                SqlDS1.SelectCommand = Trim(dt_tmp1.rows(0).item("sel_str") & "") & order_str
            Else
                SqlDS1.SelectCommand = Trim(dt_tmp1.rows(0).item("sel_str") & "") & " WHERE " & txtSubWhere2.Text & order_str
            End If
        End If

        'Response.End()

        '設定GridView資料來源ID
        If GridView1.DataSourceID Is Nothing Then
            GridView1.DataSourceID = SqlDS1.ID
        End If
    End Sub

    '程式自定區 以上





    '程式固定區-不要修改 以下

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
        '設定GridView外觀樣式
        GridView1.AutoGenerateColumns = False

        '設定GridView屬性
        GridView1.AllowPaging = True   '設定分頁
        GridView1.AllowSorting = True  '設定排序
        GridView1.Font.Size = 12       '設定字型大小
        GridView1.GridLines = GridLines.Both '設定格線
        GridView1.PageSize = 15
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
            Label2.Text = "總筆數：0/總頁數:" & GridView1.PageCount.ToString()
        Else
            Label2.Text = "總筆數：" & CStr(dt_tmp.rows.count) & "/總頁數:" & GridView1.PageCount.ToString()

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

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '資料搜尋-單一條件
        If Trim(Replace(TxtFind.Text, "'", "") & "") <> "" Then
            txtSubWhere.Text = " WHERE " & ddlField.SelectedValue & " LIKE '%" & TxtFind.Text & "%'"
        Else
            txtSubWhere.Text = " WHERE 1=1"
        End If
        setGridViewData()
    End Sub

    Protected Sub GridView1_SelectedIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSelectEventArgs) Handles GridView1.SelectedIndexChanging
        Dim Para1 As String()
        Dim para2 As String()
        Dim sScript As String

        Para1 = Split(FieldList.Text, "|")
        para2 = Split(TextBoxList.Text, "|")
        sScript = ""
        GridView1.SelectedIndex = e.NewSelectedIndex

        For ilop As Integer = LBound(Para1) To UBound(Para1)
            If GridView1.SelectedRow.Cells(Para1(ilop)).Text = "&nbsp;" Then GridView1.SelectedRow.Cells(Para1(ilop)).Text = ""
            sScript = sScript & "opener.document.getElementById('" & para2(ilop) & "').value = '" & GridView1.SelectedRow.Cells(Para1(ilop)).Text & "';"
        Next

        If sScript <> "" Then
            sScript = sScript & "window.close();"
            Response.Write("<script language='javascript'>" & sScript & "</script>")
            Response.End()
        End If

    End Sub

    '程式固定區-不要修改 以上

End Class


