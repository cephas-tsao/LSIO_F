Imports System.Data
Imports System.Drawing
Imports System.Web.Helpers
Partial Class basic_LSI026_Supload
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI026"
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

            'Set_DDL_Option("ZipCode", "", ddl_eng_area, "", "---請選擇---", "B")
            'Set_DDL_Option(ddl_eng_area.SelectedValue, "", ddl_eng_road, "", "---請選擇---", "B")
            'Set_CKL_Option("pc_kind1", "", ckl_pc_kind1, "", "", "B")
            'Set_CKL_Option("pc_kind2", "", ckl_pc_kind2, "", "", "B")
            'Set_CKL_Option("tear_tool1", "", ckl_tear_tool1, "", "", "")
            'Set_CKL_Option("cage_tool1", "", ckl_cage_tool1, "", "", "")


            '進階搜尋增加點
            'WebSubMenu.setName("GridView1", "SqlDS2", "txtOrder", "txtSubWhere")

            '設定Row的鍵值組成，具有唯一性
            Dim KeyNames() As String = {"tid"}
            GridView1.DataKeyNames = KeyNames
            ws.InsUsrRec(s_prgname, "SRC", s_prgcode, Session("usr_code"), Session("usr_ip"), "", "", "")
            txtOrder.Text = " ORDER BY pc_code DESC"
            'Response.Write(txtOrder.Text)
            'Response.End()
            Dim sqlstr As String = Get_SqlStr("supload", "SELGroup")
            Dim ds As DataSet
            Dim dt As DataTable
            ds = Get_DataSet("LSI026_SUPERVISE_ITEM", "ITEM")
            dt = ds.Tables(0)
            For Ilop = 0 To dt.Rows.Count - 1
                Dim gid As String = Convert.ToString(dt.Rows(Ilop).Item("id"))
                group_name.Items.Add(New ListItem(Trim(dt.Rows(Ilop).Item("group_name")), gid))
            Next
        End If

        '程式變數設定
        UserLimit1 = ws.Get_limit(Session("usr_code"), s_prgcode)
        'setLimit(UserLimit1)

        '設定查詢欄位
        SqlDS2.ConnectionString = ws.Get_ConnStr("con_db")
        SqlDS2.SelectCommand = Get_SqlStr("supload", "SELGroup")
       

        '程式抬頭
        'Label1.Text = s_prgname & "(" & s_prgcode & ")"

        '取得預設條件
        If Trim(Request("subwhere") & "") <> "" Then txtSubWhere.Text = Replace(Request("subwhere"), "$", "%")
        '進階搜尋增加點
        If Trim(Request("order") & "") <> "" Then txtOrder.Text = Request("order")

        '設定onclick
        'CBL_Field1.Items(0).Attributes.Add("onclick", "Chk_All_1()")

        'IB_import.Attributes("onclick") = "subwindow=window.open('FileUpload.aspx','FileUpload','height=260,width=620,top=200,left=400,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no','');subwindow.focus();return false;"

        '取得資料
        setGridViewData()
    End Sub
    Protected Sub setFields()
        '建立資料繫結欄位
        Dim tid As New BoundField()
        Dim group_name As New BoundField()
        Dim content As New BoundField()
        Dim is_enabled As New BoundField()
        'Dim road_name As New BoundField()

        '指定資料來源欄位
        tid.DataField = "tid"
        'group_name.DataField = "group_name"
        'content.DataField = "content"
        'is_enabled.DataField = "is_enabled"
        'road_name.DataField = "road_name"

        '設定欄位標頭名稱
        tid.HeaderText = "資料編號"
        group_name.HeaderText = "項目"
        content.HeaderText = "督導事項"
        is_enabled.HeaderText = "是否使用"
        'road_name.HeaderText = "地址-路段"



        '設定欄位是否自動換行(不換行的再設定)
        'prg_code.ItemStyle.Wrap = False


        '欄位寬度(%):要留6%給編輯圖示
        tid.ItemStyle.Width = Unit.Percentage(5)
        'group_name.ItemStyle.Width = Unit.Percentage(20)
        'content.ItemStyle.Width = Unit.Percentage(30)
        'is_enabled.ItemStyle.Width = Unit.Percentage(8)
        ''road_name.ItemStyle.Width = Unit.Percentage(8)


        '設定欄位標題顏色
        tid.HeaderStyle.CssClass = "grid1"
        'group_name.HeaderStyle.CssClass = "grid1"
        'content.HeaderStyle.CssClass = "grid1"
        'is_enabled.HeaderStyle.CssClass = "grid1"
        'road_name.HeaderStyle.CssClass = "grid1"


        '設定欄位的水平對齊
        tid.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        'group_name.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        'content.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        'is_enabled.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        'road_name.ItemStyle.HorizontalAlign = HorizontalAlign.Left



        '設定欄位是否顯示(不顯示的再設定)
        'com_name.Visible = False
        tid.HtmlEncode = True
        GridView1.Columns.Add(tid)
        'GridView1.Columns.Add(group_name)
        'GridView1.Columns.Add(content)
        'GridView1.Columns.Add(is_enabled)
        'GridView1.Columns.Add(road_name)
        group_name.ReadOnly = True
        tid.Visible = False


    End Sub
    Protected Sub setGridViewData()
        '取得GridView資料

        '設定SqlDataSource連線及Select命令
        SqlDS1.ConnectionString = ws.Get_ConnStr("con_db")
        '進階搜尋增加點
        If txtSubWhere.Text <> "" Then
            SqlDS1.SelectCommand = Get_SqlStr("supload", "SELECT", txtSubWhere.Text, txtOrder.Text)
        Else
            SqlDS1.SelectCommand = Get_SqlStr("supload", "SELECT", "")
        End If


        '設定GridView資料來源ID
        If GridView1.DataSourceID Is Nothing Then
            GridView1.DataSourceID = SqlDS1.ID
        End If
    End Sub
    Protected Sub GridView1_DataBound(ByVal sender As Object, ByVal e As GridViewRowEventArgs)
        'If e.Row.RowType = DataControlRowType.DataRow AndAlso e.Row.RowState = DataControlRowState.Edit Then

        'Else
        'Dim is_enable As Label = CType(e.Row.FindControl("is_enabled"), Label)
        ''Dim is_enable As Label = TryCast(e.Row.FindControl("is_enabled"), Label)
        'If is_enable Is Nothing = False Then
        '    If is_enable.Text = "True" Then
        '        is_enable.Text = "是"
        '    Else
        '        is_enable.Text = "否"
        '    End If
        'End If

        If e.Row.RowState = DataControlRowState.Edit Then
            Dim tb As DropDownList = TryCast(e.Row.FindControl("is_enabled"), DropDownList)
            Dim item As Textbox = TryCast(e.Row.FindControl("content"), Textbox)
        Else
            Dim is_enable As Label = TryCast(e.Row.FindControl("is_enabled"), Label)
            Dim item As label = TryCast(e.Row.FindControl("content"), Label)


            If is_enable Is Nothing = False Then
                If is_enable.Text = "True" Then
                    is_enable.Text = "是"
                Else
                    is_enable.Text = "否"
                End If
            End If
           
        End If
        '取得資料總筆數
        Dim dt_tmp As DataTable = Get_DataSet(s_prgcode, "SqlStr", SqlDS1.SelectCommand).Tables(0)
        If dt_tmp Is Nothing Then
            Label2.Text = "總筆數：0/總頁數:" & GridView1.PageCount.ToString() & "/每頁筆數:" & GridView1.PageSize.ToString()
        Else
            Label2.Text = "總筆數：" & CStr(dt_tmp.Rows.Count) & "/總頁數:" & GridView1.PageCount.ToString() & "/每頁筆數:" & GridView1.PageSize.ToString()

        End If

        '設定分頁訊息
        l_PageMsg.Visible = False
        If GridView1.PageCount > 1 Then
            l_PageMsg.Visible = True
            l_PageMsg.Text = "　目前所在分頁碼 (" & (GridView1.PageIndex + 1) & "/" & GridView1.PageCount.ToString() & ")"
        End If
        'End If
        ViewState("LineNo") = 0

        'If e.Row.RowType = DataControlRowType.DataRow AndAlso e.Row.RowState <> DataControlRowState.Edit Then
        'If CType(e.Row.FindControl("is_enabled") Is Nothing = True Then

        '    End If
        'End If
        'End If


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
    'Protected Sub GridView1_RowEditing(sender As Object, e As GridViewEditEventArgs)
    '    Dim tb As DropDownList = TryCast(GridView1.Rows(e.NewEditIndex).FindControl("is_enabled"), DropDownList)
    '    'GridView1.Rows[e.RowIndex].FindControl("子控制項ID")
    '    'If tb.SelectedValue = "True" Then
    '    '    tb.Text = "是"
    '    'Else
    '    '    tb.Text = "否"
    '    'End If
    'End Sub
    Protected Sub OnRowCancellingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs)
        Me.GridView1.EditIndex = -1
        GridView1.DataBind()
    End Sub

    Protected Sub GridView1_RowUpdating(sender As Object, e As GridViewUpdateEventArgs)
        Dim row As GridViewRow = GridView1.Rows(e.RowIndex)
        Dim Id = GridView1.DataKeys(e.RowIndex).Values(0)
        ' Dim name As String = TryCast(row.Cells(3).Controls(0), DropDownList).SelectedValue

        Dim tb As DropDownList = TryCast(row.Cells(3).FindControl("is_enabled"), DropDownList)
        Dim content As Textbox = TryCast(row.Cells(2).FindControl("content"), Textbox)

        'DataKeyNames = Id
        SqlDS1.UpdateCommand = Get_SqlStr("supload", "UPDATE", tb.SelectedValue, content.text, Id)
        GridView1.DataBind()
    End Sub
    Protected Sub setGridViewPage()
        '設定GridView的分頁樣式
        GridView1.PagerSettings.Mode = PagerButtons.NextPreviousFirstLast
        GridView1.PagerSettings.Visible = False

    End Sub
    Protected Sub IB_GV1_First1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_First1.Click
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
        
    End Sub

    Protected Sub drpSites_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Protected Sub Button4_Click(sender As Object, e As ImageClickEventArgs)
        'AntiForgery.Validate()
        Panel1.Visible = True
        Panel2.Visible = False
    End Sub

    Protected Sub save_Click(sender As Object, e As EventArgs)
        'AntiForgery.Validate()
        Dim newItem As String = txt_Item_name.Text
        Dim Item As String = group_name.SelectedValue
        Dim Itemname As String = group_name.SelectedItem.Text
        SaveEditData("LSI026_SUPERVISE_ITEM", "ADD", Item, "", newItem, 1)
        txt_Item_name.Text = ""
        Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('存檔完成');</script>"))
    End Sub

    Protected Sub exit_Click(sender As Object, e As EventArgs)
        'AntiForgery.Validate()
        Panel1.Visible = False
    End Sub

   
    Protected Sub AddGroup_Click(sender As Object, e As ImageClickEventArgs)
        'AntiForgery.Validate()
        Panel1.Visible = False
        Panel2.Visible = True
    End Sub

    Protected Sub SaveGroup_Click(sender As Object, e As EventArgs)
        'AntiForgery.Validate()
        Dim Gitem = TextGroup.Text
        Dim sGUID As String
        sGUID = System.Guid.NewGuid.ToString()
        SaveEditData("LSI026_SUPERVISE_ITEM", "AddGroup", sGUID, Gitem, 1)
        TextGroup.Text = ""
        Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('存檔完成');</script>"))
    End Sub

    Protected Sub eixt_group_Click(sender As Object, e As EventArgs)
        'AntiForgery.Validate()
        Panel2.Visible = False
    End Sub
End Class
