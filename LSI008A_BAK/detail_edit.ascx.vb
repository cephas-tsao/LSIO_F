Imports System.Data
Imports System.Drawing

Partial Class basic_LSI008A_detail_edit
    Inherits System.Web.UI.UserControl

    Dim p1 As New PageBase
    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI008A_detail"
    Dim UserLimit1 As String
    Dim bPrint As Boolean = False

    '程式自定區 以下

    Protected Sub setFields()
        '建立資料繫結欄位
        Dim serial_no As New BoundField()
        'Dim mrd_code As New BoundField()
        Dim usr_code As New BoundField()
        Dim usr_name As New BoundField()
        Dim usr_dep As New BoundField()
        Dim usr_mail As New BoundField()

        '指定資料來源欄位
        serial_no.DataField = "serial_no"
        'mrd_code.DataField = "mrd_code"
        usr_code.DataField = "usr_code"
        usr_name.DataField = "usr_name"
        usr_dep.DataField = "usr_dep"
        usr_mail.DataField = "usr_mail"


        '設定欄位標頭名稱
        'mrd_code.HeaderText = "預約編號"
        usr_code.HeaderText = "使用者編號"
        usr_name.HeaderText = "使用者名稱"
        usr_dep.HeaderText = "科室名稱"
        usr_mail.HeaderText = "連絡mail"

        '設定欄位是否自動換行
        'mrd_code.ItemStyle.Wrap = False

        '設定排序所對應的欄位
        'mrd_code.SortExpression = "mrd_code"
        usr_code.SortExpression = "usr_code"
        usr_name.SortExpression = "usr_name"
        usr_dep.SortExpression = "usr_dep"
        usr_mail.SortExpression = "usr_mail"


        '欄位寬度(%):要留6%給編輯圖示
        'mrd_code.ItemStyle.Width = Unit.Percentage(50)
        usr_code.ItemStyle.Width = Unit.Percentage(15)
        usr_name.ItemStyle.Width = Unit.Percentage(15)
        usr_dep.ItemStyle.Width = Unit.Percentage(20)
        usr_mail.ItemStyle.Width = Unit.Percentage(35)

        '設定欄位顏色
        'mrd_code.HeaderStyle.CssClass = "grid1"
        usr_code.HeaderStyle.CssClass = "grid1"
        usr_name.HeaderStyle.CssClass = "grid1"
        usr_dep.HeaderStyle.CssClass = "grid1"
        usr_mail.HeaderStyle.CssClass = "grid1"

        '設定欄位的水平對齊
        'mrd_code.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        usr_code.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        usr_name.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        usr_dep.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        usr_mail.ItemStyle.HorizontalAlign = HorizontalAlign.Left

        '設定欄位是否顯示(不顯示的再設定)
        serial_no.Visible = False

        GridView1.Columns.Add(serial_no)
        'GridView1.Columns.Add(mrd_code)
        GridView1.Columns.Add(usr_name)
        GridView1.Columns.Add(usr_code)
        GridView1.Columns.Add(usr_dep)
        GridView1.Columns.Add(usr_mail)
    End Sub

    Protected Sub bt_Save_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bt_Save.Click
        '資料存檔
        'If Session("UsrListDo") = "N" Then  '避免reload重覆執行
        '    Session("UsrListDo") = "Y"
        '    If txtPgmType.Text <> "ADD" Then
        '        GridView1.SelectedIndex = -1
        '        l_mode.Text = ""
        '        Panel1.Visible = False
        '    Else
        '        initEdit()
        '    End If
        '    Exit Sub
        'End If

        '檢查資料正確性
        If Chk_data() = False Then Exit Sub

        '設定記錄資料
        Dim sSql As String = p1.Get_SqlStr(s_prgcode, "Detail", txtPrimary.Text)
        Dim sFieldName As String = ws.Get_FieldName(s_prgcode)
        Dim sOldData As String = ws.Get_OldData(sSql, s_prgcode)
        Dim sNewData As String = Me.Get_NewData(s_prgcode)

        If p1.SaveEditData(s_prgcode, txtPgmType.Text, txt_serial_no.Text, txt_mrd_code.Text, txt_usr_code.Text, txt_usr_name.Text, _
                           txt_usr_mail.Text, txt_usr_dep.Text) Then
            ws.InsUsrRec(txtPrimary.Text, txtPgmType.Text, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)
            Show_Message("出席人員存檔成功!")
            If txtPgmType.Text <> "ADD" Then
                GridView1.SelectedIndex = -1
                l_mode.Text = ""
                Panel1.Visible = False
            Else
                initEdit()
            End If
        Else
            '存檔出錯就秀出錯誤訊息
            Show_Message("儲存失敗!")
        End If
        'Session("UsrListDo") = "Y"
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) Then
            setGridViewStyle()
            setFields()
            setGridViewPage()

            '進階搜尋增加點 
            IB_usr_code.Attributes("onclick") = ws.Set_AddForm("../../public/addForm.aspx", "USR-2", "1|2|3|4", "detail_edit$txt_usr_code|detail_edit$txt_usr_name|detail_edit$txt_usr_dep|detail_edit$txt_usr_mail", "")
            '設定Row的鍵值組成，具有唯一性
            Dim KeyNames() As String = {"serial_no"}
            GridView1.DataKeyNames = KeyNames

            '設定唯讀欄位
            txt_usr_code.Attributes.Add("ReadOnly", "ReadOnly")
        End If

        '程式變數設定
        UserLimit1 = ws.Get_limit(Session("usr_code"), s_prgcode)
        setLimit(UserLimit1)

        '程式抬頭
        Label1.Text = "場地預約出席人員"

        '清空訊息
        l_ErrMsg.Text = ""

        '設定編輯固定欄位
        txt_mrd_code.Text = txtKeyName.Text

        '取得資料
        setGridViewData()
    End Sub

    Protected Sub setEditData()
        Select Case txtPgmType.Text
            Case "MDY", "COPY"
                '非新增模式則預設取得資料
                Dim Dt_tmp As DataTable = p1.Get_DataSet(s_prgcode, "Detail", txtPrimary.Text).Tables(0)
                If Dt_tmp.Rows.Count > 0 Then
                    txt_serial_no.Text = Dt_tmp.Rows(0).Item("serial_no").ToString.Trim
                    'txt_mrd_code.Text = Dt_tmp.Rows(0).Item("mrd_code").ToString.Trim
                    txt_usr_code.Text = Dt_tmp.Rows(0).Item("usr_code").ToString.Trim
                    txt_usr_name.Text = Dt_tmp.Rows(0).Item("usr_name").ToString.Trim
                    txt_usr_dep.Text = Dt_tmp.Rows(0).Item("usr_dep").ToString.Trim
                    txt_usr_mail.Text = Dt_tmp.Rows(0).Item("usr_mail").ToString.Trim
                End If
            Case "ADD"
                '新增模式則清空資料
                initEdit()
        End Select
    End Sub

    Protected Sub initEdit()
        '初始化編輯區域
        txt_serial_no.Text = ""
        txt_usr_code.Text = ""
        txt_usr_name.Text = ""
        txt_usr_dep.Text = ""
        txt_usr_mail.Text = ""
        txtPrimary.Text = ""
    End Sub

    Protected Sub setGridViewData()
        '取得GridView資料

        '設定SqlDataSource連線及Select命令
        SqlDS1.ConnectionString = ws.Get_ConnStr("con_db")
        '進階搜尋增加點
        SqlDS1.SelectCommand = p1.Get_SqlStr(s_prgcode, "SELECT", txtKeyName.Text)

        '設定Delete命令
        Dim CtrPara As New ControlParameter("datakeys", "GridView1", "SelectedValue")
        SqlDS1.DeleteParameters.Add(CtrPara)
        SqlDS1.DeleteCommand = p1.Get_SqlStr(s_prgcode, "DELETE")

        '設定GridView資料來源ID
        If GridView1.DataSourceID Is Nothing Then
            GridView1.DataSourceID = SqlDS1.ID
        End If
    End Sub

    Protected Function Chk_data() As Boolean
        'Server端資料檢查 自訂函數
        Chk_data = True

        ''檢查重覆資料
        'If txtPgmType.Text = "ADD" Or txtPgmType.Text = "MDY" Then
        '    If p1.Chk_RelData(s_prgcode, "", txt_tsh_code.Text, txt_tn_code.Text) = False Then
        '        Chk_data = False
        '        Show_Message("有重覆編號資料，請確認")
        '    End If
        'End If

        '檢查資料正確性
        If txt_usr_mail.Text <> "" And p1.Chk_Email(txt_usr_mail.Text) = False Then
            Chk_data = False
            Show_Message("連絡mail格式錯誤，請確認")
        End If

        '檢查必填欄位是否空白
        If txt_usr_name.Text = "" Then
            Chk_data = False
            Show_Message("必填欄位不得為空白，請確認")
        End If
    End Function

    '程式自定區 以上





    '程式固定區-不要修改 以下
    Protected Sub setLimit(ByVal Climit As String)
        '依據權限控制物件
        GridView1.Columns(0).Visible = CBool(InStr(Climit, "D") > 0)
        GridView1.Columns(1).Visible = CBool(InStr(Climit, "M") > 0)
        GridView1.Columns(2).Visible = CBool(InStr(Climit, "A") > 0)
        GridView1.Columns(2).Visible = False
        Button4.Enabled = CBool(InStr(Climit, "A") > 0)
    End Sub

    Public Sub setKeyName(ByVal sKeyName As String)
        txtKeyName.Text = sKeyName
        bt_usr_list.Attributes("onclick") = "subwindow=window.open('UsrList.aspx?MrdCode=" & txtKeyName.Text & "','UsrList','height=550,width=600,top=100,left=200,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no','');subwindow.focus();return false;"
        bt_InFile.Attributes("onclick") = "subwindow=window.open('FileUpload.aspx?MrdCode=" & txtKeyName.Text & "','FileUpload','height=260,width=620,top=200,left=400,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no','');subwindow.focus();return false;"
        setGridViewData()
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
        Dim dt_tmp As DataTable = p1.Get_DataSet(s_prgcode, "SqlStr", SqlDS1.SelectCommand).Tables(0)
        If dt_tmp Is Nothing Then
            Label2.Text = "總筆數：0/總頁數:" & GridView1.PageCount.ToString() & "/每頁筆數:" & GridView1.PageSize.ToString()
        Else
            Label2.Text = "總筆數：" & CStr(dt_tmp.Rows.Count) & "/總頁數:" & GridView1.PageCount.ToString() & "/每頁筆數:" & GridView1.PageSize.ToString()

            ''有資料才秀
            'If CInt(dt_tmp.rows.count) > 0 Then
            '    Dim bottonPagerRow As GridViewRow = GridView1.BottomPagerRow
            '    Dim bottonPagerNo As New Label()
            '    '修改點
            '    If Not bottonPagerRow Is Nothing Then
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
        End If
        ViewState("LineNo") = 0

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
            Dim sSql As String = p1.Get_SqlStr(s_prgcode, "Detail", txtDataKey)
            Dim sFieldName As String = ws.Get_FieldName(s_prgcode)
            Dim sOldData As String = ws.Get_OldData(sSql, s_prgcode)
            Dim sNewData As String = p1.Get_NewData(s_prgcode, "DEL")

            SqlDS1.Delete()
            ws.InsUsrRec(txtDataKey, "DEL", s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)
        End If
        GridView1.SelectedIndex = -1
        l_mode.Text = ""
        Panel1.Visible = False
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
            '匯出Excel時顯示設定
            If bPrint = True Then
                '隱藏功能列
                If GridView1.Columns(0).Visible = True Then e.Row.Cells(0).Visible = False
                If GridView1.Columns(1).Visible = True Then e.Row.Cells(1).Visible = False
                If GridView1.Columns(2).Visible = True Then e.Row.Cells(2).Visible = False
            End If
        End If
        Select Case e.Row.RowType
            Case DataControlRowType.Header
                e.Row.BackColor = Color.FromArgb(153, 0, 0)
                e.Row.ForeColor = Color.White
            Case DataControlRowType.DataRow
                If p1.Chk_Email(e.Row.Cells(7).Text) = False Then e.Row.Cells(7).Text &= "[無效信箱]"

                '建立奇數資料列與偶數資料列的onmouseover及onmouseout的顏色變換
                If GridView1.SelectedIndex <> ViewState("LineNo") Then
                    If (CType(ViewState("LineNo"), Int16) Mod 2 = 0) Then
                        e.Row.Attributes.Add("onmouseout", "this.style.backgroundColor='" & p1.gridview_row_color("10") & "';this.style.color='" & p1.gridview_row_color("11") & "'")
                        e.Row.Attributes.Add("onmouseover", "this.style.backgroundColor='" & p1.gridview_row_color("30") & "';this.style.color='" & p1.gridview_row_color("31") & "'")
                    Else
                        e.Row.Attributes.Add("onmouseout", "this.style.backgroundColor='" & p1.gridview_row_color("20") & "';this.style.color='" & p1.gridview_row_color("11") & "'")
                        e.Row.Attributes.Add("onmouseover", "this.style.backgroundColor='" & p1.gridview_row_color("30") & "';this.style.color='" & p1.gridview_row_color("31") & "'")

                    End If
                End If
                ViewState("LineNo") += 1
        End Select
    End Sub

    Protected Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
        txtPgmType.Text = "ADD"
        l_mode.Text = "(新增模式)"
        Panel1.Visible = True
        GridView1.SelectedIndex = -1
        setEditData()
        'Session("UsrListDo") = "A"
    End Sub

    Protected Sub Edit_mode()
        txtPgmType.Text = "MDY"
        l_mode.Text = "(修改模式)"
        Panel1.Visible = True
    End Sub

    Protected Sub bt_Cancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bt_Cancel.Click
        l_mode.Text = ""
        Panel1.Visible = False
        GridView1.SelectedIndex = -1
        'Session("UsrListDo") = "C"
    End Sub

    Protected Sub GridView1_SelectedIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSelectEventArgs) Handles GridView1.SelectedIndexChanging
        GridView1.SelectedIndex = e.NewSelectedIndex
        If e.NewSelectedIndex >= 0 Then
            Dim txtDataKey As String = GridView1.SelectedDataKey.Value.ToString
            txtPrimary.Text = txtDataKey
            setEditData()
            'Session("UsrListDo") = "E"
        End If
    End Sub

    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        txtPgmType.Text = e.CommandArgument
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
                        'ElseIf GridView1.SortExpression = "" Then
                        '    '進階搜尋增加點
                        '    If InStr(txtOrder.Text, LinkButton.CommandArgument & " ASC") > 0 Then
                        '        myheader.Controls.Add(New LiteralControl("▲"))
                        '    ElseIf InStr(txtOrder.Text, LinkButton.CommandArgument & " DESC") > 0 Then
                        '        myheader.Controls.Add(New LiteralControl("▼"))
                        '    End If
                    End If
                End If
            Next
        End If
    End Sub

    Protected Function Get_NewData(ByVal i_prgcode As String) As String
        Get_NewData = ""
        Dim Dt_Field = ws.Get_FieldSet(i_prgcode).Tables(0)

        For i As Integer = 0 To Dt_Field.Rows.Count - 1
            If i <> 0 Then Get_NewData &= "_|@"

            Select Case Dt_Field.Rows(i).Item("data_type")
                Case "TextBox"
                    Get_NewData &= CType(Me.FindControl(Dt_Field.Rows(i).Item("data_source")), TextBox).Text
                Case "DDL"
                    Get_NewData &= CType(Me.FindControl(Dt_Field.Rows(i).Item("data_source")), DropDownList).SelectedValue
            End Select
        Next
    End Function

    Protected Sub Show_Message(ByVal sMsg As String)
        Dim ErrMsg_Type As String = ws.Get_SysPara("ErrMsg_Type", "0")

        If ErrMsg_Type = "0" Or ErrMsg_Type = "2" Then
            '彈出式對話方塊
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('" & sMsg & "');</script>"))
        End If

        If ErrMsg_Type = "1" Or ErrMsg_Type = "2" Then
            '下方紅字訊息
            l_ErrMsg.Text &= sMsg & "<br />"
        End If
    End Sub

    Protected Sub IB_GV1_First1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_First1.Click
        GridView1.PageIndex = 0
        GridView1.DataBind()
        'Session("UsrListDo") = "M"
    End Sub

    Protected Sub IB_GV1_Previous1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Previous1.Click
        GridView1.PageIndex -= 1
        GridView1.DataBind()
        'Session("UsrListDo") = "M"
    End Sub

    Protected Sub IB_GV1_Next1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Next1.Click
        GridView1.PageIndex += 1
        GridView1.DataBind()
        'Session("UsrListDo") = "M"
    End Sub

    Protected Sub IB_GV1_Last1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Last1.Click
        GridView1.PageIndex = GridView1.PageCount - 1
        GridView1.DataBind()
        'Session("UsrListDo") = "M"
    End Sub

    Protected Sub IB_GV1_First2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_First2.Click
        GridView1.PageIndex = 0
        GridView1.DataBind()
        'Session("UsrListDo") = "M"
    End Sub

    Protected Sub IB_GV1_Previous2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Previous2.Click
        GridView1.PageIndex -= 1
        GridView1.DataBind()
        'Session("UsrListDo") = "M"
    End Sub

    Protected Sub IB_GV1_Next2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Next2.Click
        GridView1.PageIndex += 1
        GridView1.DataBind()
        'Session("UsrListDo") = "M"
    End Sub

    Protected Sub IB_GV1_Last2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles IB_GV1_Last2.Click
        GridView1.PageIndex = GridView1.PageCount - 1
        GridView1.DataBind()
        'Session("UsrListDo") = "M"
    End Sub


    '程式固定區-不要修改 以上

    Protected Sub IB_export_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles IB_export.Click
        'Response.Clear()
        'Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")
        'Response.AddHeader("content-disposition", "attachment;filename=Export.xls")
        'Response.ContentType = "application/vnd.xls"
        'Dim sw As System.IO.StringWriter = New System.IO.StringWriter()
        'Dim htw As System.Web.UI.HtmlTextWriter = New HtmlTextWriter(sw)

        'Dim hf As HtmlForm = New HtmlForm()
        'Dim NewTable As New Table
        'Dim NewRow As New TableRow
        'Dim NewCell1 As New TableCell
        'Dim NewCell2 As New TableCell
        'Dim NewCell3 As New TableCell
        'Dim NewCell4 As New TableCell
        'NewCell1.Text = "使用者名稱"
        'NewCell2.Text = "使用者名稱"
        'NewCell3.Text = "使用者名稱"
        'NewCell4.Text = "使用者名稱"
        'NewCell1.Attributes.Add("style", "text-align:center")
        'NewRow.Cells.Add(NewCell1)
        'NewTable.Rows.Add(NewRow)

        'For i As Integer = 0 To GridView1.Rows.Count - 1
        '    Dim TmpRow As New TableRow
        '    Dim TmpCell As New TableCell

        '    TmpCell.Text = GridView1.Rows(i).Cells(4).Text
        '    TmpCell.Attributes.Add("style", "text-align:center")
        '    TmpRow.Cells.Add(TmpCell)
        '    NewTable.Rows.Add(TmpRow)
        'Next

        'Controls.Add(hf)
        'hf.Controls.Add(NewTable)
        'hf.RenderControl(htw)

        'Response.Write(sw.ToString())
        'Response.End()

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
End Class

