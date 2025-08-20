Imports System.Data
Imports System.Drawing
Imports System.IO
Imports System.Web.Helpers

Partial Class basic_LSI026_Com_check
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "Self_com"
    Dim s_prgname As String = ws.Get_PrgName(s_prgcode)
    Dim UserLimit1 As String
    Dim bPrint As Boolean = False
    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
        'checks database connection string and handles error if there is one
        anticfrs.InnerHtml = AntiForgery.GetHtml().ToString()
    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) Then
            'setGridViewStyle()
            'setFields()
            'setCheckBox()
            setGridViewPage()
            setGridViewData()

            '進階搜尋增加點
            WebSubMenu.setName("GridView1", "SqlDS2", "txtOrder", "txtSubWhere")

            '設定Row的鍵值組成，具有唯一性
            Dim KeyNames() As String = {"cid"}
            GridView1.DataKeyNames = KeyNames
            'ws.InsUsrRec(s_prgname, "SRC", s_prgcode, Session("usr_code"), Session("usr_ip"), "", "", "")
            txtOrder.Text = " ORDER BY com_id ASC"
            'Response.Write(txtOrder.Text)
            'Response.End()
        End If

        '程式變數設定
        'UserLimit1 = ws.Get_limit(Session("usr_code"), s_prgcode)
        'setLimit(UserLimit1)

        '設定查詢欄位
        'SqlDS2.ConnectionString = ws.Get_ConnStr("con_db")
        'SqlDS2.SelectCommand = Get_SqlStr(s_prgcode, "ddlField")

        '程式抬頭(要在BDP030資料表內新增title)
        Label1.Text = s_prgname & "(" & s_prgcode & ")"

        '取得預設條件
        If Trim(Request("subwhere") & "") <> "" Then txtSubWhere.Text = HttpContext.Current.Server.HtmlEncode(Replace(Request("subwhere"), "$", "%"))
        '進階搜尋增加點
        If Trim(Request("order") & "") <> "" Then txtOrder.Text = HttpContext.Current.Server.HtmlEncode(Request("order"))

        '設定onclick
        'CBL_Field1.Items(0).Attributes.Add("onclick", "Chk_All_1()")

        IB_import.Attributes("onclick") = "subwindow=window.open('FileUpload.aspx','FileUpload','height=260,width=620,top=200,left=400,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no','');subwindow.focus();return false;"
        'img_date1.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_spc_date")
        'img_date2.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_epc_date")
        'img_date3.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_spc_date_s")
        'img_date4.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_epc_date_s")

        'img_date5.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_spc_date_e")
        'img_date6.Attributes("onclick") = ws.Set_Calendar("../../public/calendar.aspx", "txt_epc_date_e")
        '取得資料
        'setGridViewData()
        'keepdata()
    End Sub
    Protected Sub keepdata()
        Dim dt As DataTable
        If GridView1.DataSource Is Nothing Then
            dt = ViewState("dt")
            GridView1.DataSource = dt
        Else
            dt = GridView1.DataSource
        End If
        For i = 0 To GridView1.Rows.Count - 1
            Dim textbox1 As System.Web.UI.WebControls.TextBox
            Dim textbox2 As System.Web.UI.WebControls.TextBox
            Dim textbox3 As System.Web.UI.WebControls.TextBox
            Dim textbox4 As System.Web.UI.WebControls.TextBox
            Dim tex_cid As System.Web.UI.WebControls.Label
            textbox1 = DirectCast(Me.GridView1.Rows(i).Cells(0).FindControl("sea1"), System.Web.UI.WebControls.TextBox)
            textbox2 = DirectCast(Me.GridView1.Rows(i).Cells(1).FindControl("sea2"), System.Web.UI.WebControls.TextBox)
            textbox3 = DirectCast(Me.GridView1.Rows(i).Cells(2).FindControl("sea3"), System.Web.UI.WebControls.TextBox)
            textbox4 = DirectCast(Me.GridView1.Rows(i).Cells(3).FindControl("sea4"), System.Web.UI.WebControls.TextBox)
            tex_cid = DirectCast(Me.GridView1.Rows(i).Cells(4).FindControl("cid"), System.Web.UI.WebControls.Label)
            Dim t1 As String = HttpUtility.HtmlEncode(textbox1.Text)
            Dim t2 As String = HttpUtility.HtmlEncode(textbox2.Text)
            Dim t3 As String = HttpUtility.HtmlEncode(textbox3.Text)
            Dim t4 As String = HttpUtility.HtmlEncode(textbox4.Text)
            Dim t_cid As String = tex_cid.Text
            For j = 0 To dt.Rows.Count - 1
                Dim CID = dt.Rows(j)("cid")
                If t_cid = CID Then
                    dt.Rows(j)("sea1") = t1
                    dt.Rows(j)("sea2") = t2
                    dt.Rows(j)("sea3") = t3
                    dt.Rows(j)("sea4") = t4
                End If
            Next
            

        Next

        ViewState("dt") = dt
    End Sub

    '程式自定區 以下

    Protected Sub setFields()
        '建立資料繫結欄位
        Dim com_id As New BoundField()
        'Dim Uniform_num As New BoundField()
        Dim com_name As New TemplateField()
        Dim sea1 As New BoundField()
        Dim sea2 As New BoundField()
        Dim sea3 As New BoundField()
        Dim sea4 As New BoundField()

        com_id.DataField = "com_id"
        'Uniform_num.DataField = "Uniform_num"
        'com_name.ItemTemplate = "com_name"
        'tmpFd.ItemTemplate = New TemplateField(DataControlRowType.DataRow, 欄位id, 欄位內容, 額外參數(可自定))
        ' lname.ItemTemplate = new AddTemplateToGridView(ListItemType.Item, "au_lname");
        'com_name.HeaderText = "com_name"
        sea1.DataField = "sea1"
        sea2.DataField = "sea2"
        sea3.DataField = "sea3"
        sea4.DataField = "sea4"


        '設定欄位標頭名稱
        com_id.HeaderText = "資料編號"
        'Uniform_num.HeaderText = "統編"
        'com_name.HeaderText = "事業單位名稱"
        sea1.HeaderText = "第一季"
        sea2.HeaderText = "第二季"
        sea3.HeaderText = "第三季"
        sea4.HeaderText = "第四季"



        '設定欄位是否自動換行(不換行的再設定)
        'prg_code.ItemStyle.Wrap = False

        '設定排序所對應的欄位
        com_id.SortExpression = "com_id"
        'Uniform_num.SortExpression = "Uniform_num"
        ' com_name.SortExpression = "com_name"
        'sea1.SortExpression = "chk_name"
        'sea2.SortExpression = "unit"
        'sea3.SortExpression = "start_year"
        'sea4.SortExpression = "end_year"


        '欄位寬度(%):要留6%給編輯圖示
        com_id.ItemStyle.Width = Unit.Percentage(5)
        'Uniform_num.ItemStyle.Width = Unit.Percentage(10)
        com_name.ItemStyle.Width = Unit.Percentage(20)
        sea1.ItemStyle.Width = Unit.Percentage(10)
        sea2.ItemStyle.Width = Unit.Percentage(10)
        sea3.ItemStyle.Width = Unit.Percentage(10)
        sea4.ItemStyle.Width = Unit.Percentage(10)


        '設定欄位標題顏色
        com_id.HeaderStyle.CssClass = "grid1"
        'Uniform_num.HeaderStyle.CssClass = "grid1"
        com_name.HeaderStyle.CssClass = "grid1"
        sea1.HeaderStyle.CssClass = "grid1"
        sea2.HeaderStyle.CssClass = "grid1"
        sea3.HeaderStyle.CssClass = "grid1"
        sea4.HeaderStyle.CssClass = "grid1"


        '設定欄位的水平對齊
        com_id.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        'Uniform_num.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        com_name.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        sea1.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        sea2.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        sea3.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        sea4.ItemStyle.HorizontalAlign = HorizontalAlign.Left



        '設定欄位是否顯示(不顯示的再設定)
        com_name.Visible = False
        com_id.HtmlEncode = True
        GridView1.Columns.Add(com_id)
        'GridView1.Columns.Add(Uniform_num)
        GridView1.Columns.Add(com_name)
        GridView1.Columns.Add(sea1)
        GridView1.Columns.Add(sea2)
        GridView1.Columns.Add(sea3)
        GridView1.Columns.Add(sea4)


    End Sub
    ' ====== 可查詢欄位（依實際 SQL 調整；你頁面目前顯示這些）======
    Private Shared ReadOnly AllowedQueryFields As String() = {
    "com_name", "sea1", "sea2", "sea3", "sea4", "cid"
}

    ' ====== 可排序欄位（依 GridView SortExpression/SQL 調整）======
    Private Shared ReadOnly AllowedOrderFields As String() = {
    "com_name", "sea1", "sea2", "sea3", "sea4"
}
    Private Shared ReadOnly AllowedDirections As String() = {"ASC", "DESC"}

    ' 年度白名單（依你 DropDownList 選項）
    Private Function SafeYear(ByVal v As String) As String
        Select Case v
            Case "110", "111", "112" : Return v
            Case Else : Return "111"
        End Select
    End Function

    ' 類別白名單（對應 ddlField 的兩個選項）
    Private Function SafeField(ByVal v As String) As String
        If String.IsNullOrEmpty(v) Then Return "一般科"
        If v = "一般科" OrElse v = "勞條科" Then Return v
        Return "一般科"
    End Function

    ' 清理搜尋值（僅 XSS 風險降低；SQL 安全請在 Get_DataSet 內參數化）
    Private Function SanitizeSearchValue(ByVal v As String) As String
        If v Is Nothing Then Return String.Empty
        Dim s As String = v.Replace("%", "$")
        s = Regex.Replace(s, "[\u0000-\u001F]", "")
        If s.Length > 400 Then s = s.Substring(0, 400)
        Return s
    End Function

    ' 將 txtOrder.Text 正規化為「欄位 空白 方向」，不合法則回預設
    Private Function NormalizeOrder(ByVal input As String) As String
        Dim defaultOrder As String = "com_name ASC"
        If String.IsNullOrWhiteSpace(input) Then Return defaultOrder

        Dim m As Match = Regex.Match(input, "^\s*([A-Za-z_]\w*)\s*(ASC|DESC)?\s*$", RegexOptions.IgnoreCase)
        If Not m.Success Then Return defaultOrder

        Dim col As String = m.Groups(1).Value
        Dim dir As String = If(m.Groups(2).Success, m.Groups(2).Value.ToUpperInvariant(), "ASC")

        Dim colOk As Boolean = False
        For Each s As String In AllowedOrderFields
            If String.Equals(s, col, StringComparison.OrdinalIgnoreCase) Then
                colOk = True : Exit For
            End If
        Next
        If Not colOk Then Return defaultOrder

        If dir <> "ASC" AndAlso dir <> "DESC" Then dir = "ASC"
        Return col & " " & dir
    End Function

    ' 將資料逐格 HtmlEncode（.NET 3.5 無 <%#: %>，TemplateField 用這個最穩）
    Private Function HtmlEncodeDataTable(ByVal src As DataTable) As DataTable
        Dim dst As DataTable = src.Clone() ' 保留原欄位結構/型別
        For Each row As DataRow In src.Rows
            Dim newRow As DataRow = dst.NewRow()
            For c As Integer = 0 To src.Columns.Count - 1
                Dim val As Object = row(c)
                If val IsNot Nothing AndAlso TypeOf val Is String Then
                    newRow(c) = HttpUtility.HtmlEncode(DirectCast(val, String))
                Else
                    newRow(c) = val
                End If
            Next
            dst.Rows.Add(newRow)
        Next
        Return dst
    End Function


    Protected Sub setGridViewData()
        '取得GridView資料

        ' 1) 年度/類別欄位：僅允許清單內（避免亂值）
        Dim yearVal As String = SafeYear(DDLYear.SelectedValue)
        Dim fieldVal As String = SafeField(ddlField.SelectedValue)

        ' 2) 查詢條件：清理控制字元 + 長度限制 + 你原本的 %->$
        Dim searchVal As String = SanitizeSearchValue(If(txtSubWhere.Text, String.Empty))

        ' 3) 排序：只允許白名單欄位 + ASC/DESC
        Dim orderBy As String = NormalizeOrder(If(txtOrder.Text, String.Empty))

        ' 4) 查詢

        Dim ds As DataSet = Get_DataSet(s_prgcode, fieldVal, searchVal, orderBy)
        'Dim ds As Data.DataSet = Get_DataSet(s_prgcode, ddlField.SelectedValue, txtSubWhere.Text, txtOrder.Text)
        Dim dt As DataTable
        dt = ds.Tables(0)
        Dim ds1 As Data.DataSet = Get_DataSet("LSI026", ddlField.SelectedValue, txtSubWhere.Text, txtOrder.Text)
        Dim dt1 = ds1.Tables(0)
        If dt.Rows.Count > 0 Then
            dt.Columns.Add("sea1")
            dt.Columns.Add("sea2")
            dt.Columns.Add("sea3")
            dt.Columns.Add("sea4")
            If dt1.Rows.count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    Dim com_name As String = dt.Rows(i)("com_name")
                    For j = 0 To dt1.Rows.Count - 1
                        Dim com_name1 As String = dt1.Rows(j)("com_name")
                        If com_name = com_name1 Then
                            Dim season As String = dt1.Rows(j)("pc_season")
                            If season = "1" Then
                                dt.Rows(i)("sea1") = "是"
                            ElseIf dt.Rows(i)("pc_season") = "2" Then
                                dt.Rows(i)("sea2") = "是"
                            ElseIf dt.Rows(i)("pc_season") = "3" Then
                                dt.Rows(i)("sea3") = "是"
                            ElseIf dt.Rows(i)("pc_season") = "4" Then
                                dt.Rows(i)("sea4") = "是"
                            End If
                        End If
                    Next
                Next
            End If
            'For i = 0 To dt.Rows.Count - 1
            '    '对单条记录操作
            '    If IsDBNull(dt.Rows(i)("pc_season")) = False Then
            '        Dim season As String = dt.Rows(i)("pc_season")
            '        If season = "1" Then
            '            dt.Rows(i)("sea1") = "是"
            '        ElseIf dt.Rows(i)("pc_season") = "2" Then
            '            dt.Rows(i)("sea2") = "是"
            '        ElseIf dt.Rows(i)("pc_season") = "3" Then
            '            dt.Rows(i)("sea3") = "是"
            '        ElseIf dt.Rows(i)("pc_season") = "4" Then
            '            dt.Rows(i)("sea4") = "是"
            '        End If
            '    End If
            'Next


        End If

        'Else
        'SqlDS1.SelectCommand = Get_SqlStr(s_prgcode, "Selfcheck", "")
        'End If
        'GridView1.DataSource = dt


        'Response.Write(SqlDS1.SelectCommand)
        'Response.End()
        '設定Delete命令
        'Dim CtrPara As New ControlParameter("datakeys", "GridView1", "SelectedValue")
        'SqlDS1.DeleteParameters.Add(CtrPara)
        'SqlDS1.DeleteCommand = Get_SqlStr(s_prgcode, "DELETE")
        Try
            GridView1.DataSource = dt
            ViewState("dt") = dt
            GridView1.DataBind()
        Catch
            Throw
        End Try
        '設定GridView資料來源ID
        'If GridView1.DataSourceID Is Nothing Then
        '    GridView1.DataSourceID = SqlDS1.ID
        'End If
    End Sub
    'Protected Sub GridView1_DataBinding(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.DataBinding
    '    '取得資料
    '    setGridViewData()
    'End Sub
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
    Protected Sub setGridViewPage()
        '設定GridView的分頁樣式
        GridView1.PagerSettings.Mode = PagerButtons.NextPreviousFirstLast
        GridView1.PagerSettings.Visible = False

    End Sub
    Protected Sub IB_GV1_First1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_First1.Click
        'AntiForgery.Validate()
        GridView1.PageIndex = 0
        keepdata()
     
        GridView1.DataBind()
    End Sub

    Protected Sub IB_GV1_Previous1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Previous1.Click
        'AntiForgery.Validate()
        GridView1.PageIndex -= 1
        keepdata()
      
        GridView1.DataBind()
    End Sub

    Protected Sub IB_GV1_Next1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Next1.Click
        'AntiForgery.Validate()
        GridView1.PageIndex += 1
        keepdata()

        GridView1.DataBind()
    End Sub

    Protected Sub IB_GV1_Last1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Last1.Click
        'AntiForgery.Validate()
        GridView1.PageIndex = GridView1.PageCount - 1
        keepdata()
      
        GridView1.DataBind()
    End Sub

    Protected Sub IB_GV1_First2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_First2.Click
        'AntiForgery.Validate()
        GridView1.PageIndex = 0
        keepdata()
      
        GridView1.DataBind()
    End Sub

    Protected Sub IB_GV1_Previous2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Previous2.Click
        AntiForgery.Validate()
        GridView1.PageIndex -= 1
        keepdata()
      
        GridView1.DataBind()
    End Sub

    Protected Sub IB_GV1_Next2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Next2.Click
        AntiForgery.Validate()
        GridView1.PageIndex += 1
        keepdata()

        GridView1.DataBind()
    End Sub

    Protected Sub IB_GV1_Last2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Last2.Click
        AntiForgery.Validate()
        GridView1.PageIndex = GridView1.PageCount - 1
        keepdata()

        GridView1.DataBind()
    End Sub

    Protected Sub Bt_srch_Click(sender As Object, e As EventArgs)
        AntiForgery.Validate()
        txtSubWhere.Text = HttpUtility.UrlEncode(DDLYear.SelectedValue)
        setGridViewData()
        GridView1.DataBind()
    End Sub

    Protected Sub IB_export_Click(sender As Object, e As ImageClickEventArgs)
        AntiForgery.Validate()
        keepdata()
        Dim filename As String = Date.Now()
        Dim dt As DataTable
        If GridView1.DataSource Is Nothing Then
            dt = ViewState("dt")
            GridView1.DataSource = dt
        Else
            dt = GridView1.DataSource
        End If
        If GridView1.Rows.Count > 0 Then
            '處理 GridView 有分頁、用 DataSourceID 
            GridView1.AllowPaging = False
            GridView1.DataBind()
            Response.ClearContent()
            Response.AddHeader("content-disposition", "attachment; filename=" & filename & ".xls")
            Response.ContentType = "application/vnd.ms-excel"
            Dim sw As New StringWriter()
            Dim htw As New HtmlTextWriter(sw)
            GridView1.RenderControl(htw)
            Response.Write(sw.ToString())
            Response.[End]()

        End If
    End Sub
   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        '處理'GridView' 的控制項 'GridView' 必須置於有 runat=server 的表單標記之中
    End Sub
End Class
