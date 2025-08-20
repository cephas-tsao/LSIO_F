Imports System.Data
Imports System.Drawing
Imports System.Web.Helpers
Partial Class basic_LSI026_Company
    Inherits PageBase
    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "Self_com"
    Dim s_prgname As String = ws.Get_PrgName(s_prgcode)
    Dim UserLimit1 As String
    Dim bPrint As Boolean = False
    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
        'checks database connection string and handles error if there is one
        'anticfrs.InnerHtml = AntiForgery.GetHtml().ToString()
    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) Then
            'setGridViewStyle()
            setFields()
            'setCheckBox()
            setGridViewPage()


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
        SqlDS2.ConnectionString = ws.Get_ConnStr("con_db")
        SqlDS2.SelectCommand = Get_SqlStr(s_prgcode, "ddlField")

        '程式抬頭(要在BDP030資料表內新增title)
        Label1.Text = "匯入公司"'s_prgname & "(" & s_prgcode & ")"

        '取得預設條件
        If Trim(Request("subwhere") & "") <> "" Then txtSubWhere.Text = Replace(Request("subwhere"), "$", "%")
        '進階搜尋增加點
        If Trim(Request("order") & "") <> "" Then txtOrder.Text = Request("order")

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
        setGridViewData()
    End Sub

    '程式自定區 以下

    Protected Sub setFields()
        '建立資料繫結欄位
        Dim com_id As New BoundField()
        Dim Uniform_num As New BoundField()
        Dim com_name As New BoundField()
        Dim chk_name As New BoundField()
        Dim unit_name As New BoundField()
        Dim start_year As New BoundField()
        Dim end_year As New BoundField()

        com_id.DataField = "com_id"
        Uniform_num.DataField = "Uniform_num"
        com_name.DataField = "com_name"
        chk_name.DataField = "chk_name"
        unit_name.DataField = "unit"
        start_year.DataField = "start_year"
        end_year.DataField = "end_year"


        '設定欄位標頭名稱
        com_id.HeaderText = "資料編號"
        Uniform_num.HeaderText = "統編"
        com_name.HeaderText = "事業單位名稱"
        chk_name.HeaderText = "檢查人"
        unit_name.HeaderText = "匯入單位"
        start_year.HeaderText = "起始年度"
        end_year.HeaderText = "結束年度"



        '設定欄位是否自動換行(不換行的再設定)
        'prg_code.ItemStyle.Wrap = False

        '設定排序所對應的欄位
        com_id.SortExpression = "com_id"
        Uniform_num.SortExpression = "Uniform_num"
        com_name.SortExpression = "com_name"
        chk_name.SortExpression = "chk_name"
        unit_name.SortExpression = "unit"
        start_year.SortExpression = "start_year"
        end_year.SortExpression = "end_year"


        '欄位寬度(%):要留6%給編輯圖示
        com_id.ItemStyle.Width = Unit.Percentage(5)
        Uniform_num.ItemStyle.Width = Unit.Percentage(10)
        com_name.ItemStyle.Width = Unit.Percentage(20)
        chk_name.ItemStyle.Width = Unit.Percentage(10)
        unit_name.ItemStyle.Width = Unit.Percentage(10)
        start_year.ItemStyle.Width = Unit.Percentage(10)
        end_year.ItemStyle.Width = Unit.Percentage(10)


        '設定欄位標題顏色
        com_id.HeaderStyle.CssClass = "grid1"
        Uniform_num.HeaderStyle.CssClass = "grid1"
        com_name.HeaderStyle.CssClass = "grid1"
        chk_name.HeaderStyle.CssClass = "grid1"
        unit_name.HeaderStyle.CssClass = "grid1"
        start_year.HeaderStyle.CssClass = "grid1"
        end_year.HeaderStyle.CssClass = "grid1"


        '設定欄位的水平對齊
        com_id.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        Uniform_num.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        com_name.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        chk_name.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        unit_name.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        start_year.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        end_year.ItemStyle.HorizontalAlign = HorizontalAlign.Left



        '設定欄位是否顯示(不顯示的再設定)
        'com_name.Visible = False
        com_id.HtmlEncode = True
        GridView1.Columns.Add(com_id)
        GridView1.Columns.Add(Uniform_num)
        GridView1.Columns.Add(com_name)
        GridView1.Columns.Add(chk_name)
        GridView1.Columns.Add(unit_name)
        GridView1.Columns.Add(start_year)
        GridView1.Columns.Add(end_year)


    End Sub
    Protected Sub setGridViewData()
        '取得GridView資料

        '設定SqlDataSource連線及Select命令
        SqlDS1.ConnectionString = ws.Get_ConnStr("con_db")
        '進階搜尋增加點
        If txtSubWhere.Text <> "" Or txtOrder.Text <> "" Then
            SqlDS1.SelectCommand = Get_SqlStr(s_prgcode, "SELECT", txtSubWhere.Text, txtOrder.Text)
        Else
            SqlDS1.SelectCommand = Get_SqlStr(s_prgcode, "SELECT", "")
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
    Protected Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridView1.RowDeleting
        GridView1.SelectedIndex = e.RowIndex
        Dim txtDataKey As String = GridView1.SelectedDataKey.Value.ToString
        If Not txtDataKey = "" Then
            '設定記錄資料
            Dim sSql As String = Get_SqlStr(s_prgcode, "Detail", txtDataKey)
            Dim sFieldName As String = ws.Get_FieldName(s_prgcode)
            Dim sOldData As String = ws.Get_OldData(sSql, s_prgcode)
            Dim sNewData As String = Get_NewData(s_prgcode, "DELETE")

            SqlDS1.Delete()
            ws.InsUsrRec(txtDataKey, "DEL", s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)
        End If
        GridView1.SelectedIndex = -1
    End Sub
    Protected Sub IB_GV1_First1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_First1.Click
        'AntiForgery.Validate()
        GridView1.PageIndex = 0
        GridView1.DataBind()
    End Sub

    Protected Sub IB_GV1_Previous1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Previous1.Click
        'AntiForgery.Validate()
        GridView1.PageIndex -= 1
        GridView1.DataBind()
    End Sub

    Protected Sub IB_GV1_Next1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Next1.Click
        'AntiForgery.Validate()
        GridView1.PageIndex += 1
        GridView1.DataBind()
    End Sub

    Protected Sub IB_GV1_Last1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Last1.Click
        'AntiForgery.Validate()
        GridView1.PageIndex = GridView1.PageCount - 1
        GridView1.DataBind()
    End Sub

    Protected Sub IB_GV1_First2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_First2.Click
        'AntiForgery.Validate()
        GridView1.PageIndex = 0
        GridView1.DataBind()
    End Sub

    Protected Sub IB_GV1_Previous2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Previous2.Click
        'AntiForgery.Validate()
        GridView1.PageIndex -= 1
        GridView1.DataBind()
    End Sub

    Protected Sub IB_GV1_Next2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Next2.Click
        'AntiForgery.Validate()
        GridView1.PageIndex += 1
        GridView1.DataBind()
    End Sub

    Protected Sub IB_GV1_Last2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Last2.Click
        'AntiForgery.Validate()
        GridView1.PageIndex = GridView1.PageCount - 1
        GridView1.DataBind()
    End Sub

    Protected Sub Bt_srch_Click(sender As Object, e As EventArgs)
        'AntiForgery.Validate()
        txtSubWhere.Text = " where unit='" & ddlField.SelectedValue & "'"
        GridView1.DataBind()
    End Sub
End Class
