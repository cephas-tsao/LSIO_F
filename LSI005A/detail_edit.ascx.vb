Imports System.Data
Imports System.Drawing

Partial Class basic_LSI004A_detail_edit
    Inherits System.Web.UI.UserControl

    Dim p1 As New PageBase
    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI005A_detail"
    Dim UserLimit1 As String

    '程式自定區 以下

    Protected Sub setFields()
        '建立資料繫結欄位
        Dim serial_no As New BoundField()
        Dim sch_code As New BoundField()
        Dim mail_list As New BoundField()
        Dim cmemo As New BoundField()

        '指定資料來源欄位
        serial_no.DataField = "serial_no"
        sch_code.DataField = "sch_code"
        mail_list.DataField = "mail_list"
        cmemo.DataField = "cmemo"


        '設定欄位標頭名稱
        sch_code.HeaderText = "排程編號"
        mail_list.HeaderText = "Mail位址"
        cmemo.HeaderText = "備註"

        '設定欄位是否自動換行
        'sch_code.ItemStyle.Wrap = False

        '設定排序所對應的欄位
        sch_code.SortExpression = "sch_code"
        mail_list.SortExpression = "mail_list"
        cmemo.SortExpression = "cmemo"


        '欄位寬度(%):要留6%給編輯圖示
        sch_code.ItemStyle.Width = Unit.Percentage(12)
        mail_list.ItemStyle.Width = Unit.Percentage(60)
        cmemo.ItemStyle.Width = Unit.Percentage(30)

        '設定欄位顏色
        sch_code.HeaderStyle.CssClass = "grid1"
        mail_list.HeaderStyle.CssClass = "grid1"
        cmemo.HeaderStyle.CssClass = "grid1"

        '設定欄位的水平對齊
        sch_code.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        mail_list.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        cmemo.ItemStyle.HorizontalAlign = HorizontalAlign.Center

        '設定欄位是否顯示(不顯示的再設定)
        serial_no.Visible = False

        'GridView1.Columns.Add(serial_no)
        'GridView1.Columns.Add(sch_code)
        GridView1.Columns.Add(mail_list)
        GridView1.Columns.Add(cmemo)
    End Sub

    Protected Sub bt_Save_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bt_Save.Click
        '資料存檔
        '檢查資料正確性
        If Chk_data() = False Then Exit Sub

        '設定記錄資料
        Dim sSql As String = p1.Get_SqlStr(s_prgcode, "Detail", txtPrimary.Text)
        Dim sFieldName As String = ws.Get_FieldName(s_prgcode)
        Dim sOldData As String = ws.Get_OldData(sSql, s_prgcode)
        Dim sNewData As String = Me.Get_NewData(s_prgcode)

        If p1.SaveEditData(s_prgcode, txtPgmType.Text, txtPrimary.Text, txt_sch_code.Text, txt_mail_list.Text, txt_cmemo.Text) Then
            ws.InsUsrRec(txtPrimary.Text, txtPgmType.Text, s_prgcode, Session("usr_code"), Session("usr_ip"), sFieldName, sOldData, sNewData)

            '建立電子報
            'Set_Paper_Html()

            Show_Message("存檔成功")
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
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) Then
            setGridViewStyle()
            setFields()
            setGridViewPage()

            '進階搜尋增加點 
            'WebSubMenu.setName("GridView1", "SqlDS2", "txtOrder", "txtSubWhere")
            'p1.Set_RBL_Option("news_type", "01", rbl_news_type, "", "", "")

            '設定Row的鍵值組成，具有唯一性
            Dim KeyNames() As String = {"serial_no"}
            GridView1.DataKeyNames = KeyNames

            '設定隱藏欄位
            'txt_mail_list.Attributes.Add("style", "display:none")

            '設定唯讀欄位
            txt_serial_no.Attributes.Add("ReadOnly", "ReadOnly")
        End If

        '程式變數設定
        UserLimit1 = ws.Get_limit(Session("usr_code"), s_prgcode)
        setLimit(UserLimit1)

        ''設定查詢欄位
        'SqlDS2.ConnectionString = ws.Get_ConnStr("con_db")
        'SqlDS2.SelectCommand = p1.Get_SqlStr(s_prgcode, "ddlField")

        '程式抬頭
        Label1.Text = "郵件補送清單"

        '清空訊息
        l_ErrMsg.Text = ""

        '設定編輯固定欄位
        txt_sch_code.Text = txtKeyName.Text
        'txt_tn_code.Enabled = False

        Button5.Attributes.Add("onclick", "window.open ('default2.aspx?sch_code=" & txt_sch_code.Text & "', '選項編輯畫面', 'height=800, width=740, toolbar=no, menubar=no, scrollbars=yes, resizable=yes, location=no, status=no');")


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
                    txt_mail_list.Text = Dt_tmp.Rows(0).Item("mail_list").ToString.Trim
                    txt_cmemo.Text = Dt_tmp.Rows(0).Item("cmemo").ToString.Trim
                End If
            Case "ADD"
                '新增模式則清空資料
                initEdit()
        End Select
    End Sub

    Protected Sub initEdit()
        '初始化編輯區域
        txtPrimary.Text = ""
        txt_serial_no.Text = ""
        txt_mail_list.Text = ""
        txt_cmemo.Text = ""
    End Sub

    Protected Sub setGridViewData()
        ''取得GridView資料

        '設定SqlDataSource連線及Select命令
        SqlDS1.ConnectionString = ws.Get_ConnStr("con_db")
        '進階搜尋增加點
        SqlDS1.SelectCommand = p1.Get_SqlStr(s_prgcode, "SELECT", txtKeyName.Text, txtSubWhere.Text)

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
        'If txtPgmType.Text = "ADD" Then
        '    If p1.Chk_RelData(s_prgcode, "", txt_sch_code.Text, txt_serial_no.Text) = False Then
        '        Chk_data = False
        '        Show_Message("有重覆編號資料，請確認")
        '    End If
        'End If

        '檢查必填欄位是否空白
        If txt_mail_list.Text = "" Then
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
        Button1.Enabled = CBool(InStr(Climit, "S") > 0)
        Button4.Enabled = CBool(InStr(Climit, "A") > 0)
    End Sub

    Public Sub setKeyName(ByVal sKeyName As String)
        txtKeyName.Text = sKeyName
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
        Select Case e.Row.RowType
            Case DataControlRowType.Header
                e.Row.BackColor = Color.FromArgb(153, 0, 0)
                e.Row.ForeColor = Color.White
            Case DataControlRowType.DataRow
                If p1.Chk_Email(e.Row.Cells(3).Text) = False Then e.Row.Cells(3).Text &= "[無效信箱]"

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

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles Button1.Click
        '資料搜尋-單一條件
        If TxtFind.Text <> "" Then
            txtSubWhere.Text = " WHERE " & ddlField.SelectedValue & " LIKE N'%" & p1.RepSql(TxtFind.Text) & "%'"
        Else
            txtSubWhere.Text = ""
        End If
        setGridViewData()
    End Sub

    Protected Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
        txtPgmType.Text = "ADD"
        l_mode.Text = "(新增模式)"
        Panel1.Visible = True
        GridView1.SelectedIndex = -1
        setEditData()
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
    End Sub

    Protected Sub GridView1_SelectedIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSelectEventArgs) Handles GridView1.SelectedIndexChanging
        GridView1.SelectedIndex = e.NewSelectedIndex
        If e.NewSelectedIndex >= 0 Then
            Dim txtDataKey As String = GridView1.SelectedDataKey.Value.ToString
            txtPrimary.Text = txtDataKey
            setEditData()
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

    Public Sub Set_Paper_Html()
        '電子報主檔
        Dim sWebLogo As String = p1.Get_SysPara("web") & "/FileUpload/File_LSI004A/"
        Dim sWebPic As String = p1.Get_SysPara("web") & "/FileUpload/File_LSI016A/"
        Dim dtPaper As DataTable = p1.Get_DataSet(s_prgcode, "Paper", txt_sch_code.Text).Tables(0)
        '若無資料則離開
        If dtPaper.Rows.Count = 0 Then Exit Sub
        '取得主檔資料
        Dim sPaperLogo1 As String = dtPaper.Rows(0).Item("paper_logo") & ""
        Dim sPaperBottom1 As String = dtPaper.Rows(0).Item("paper_bottom1") & ""
        Dim sPaperBottom2 As String = dtPaper.Rows(0).Item("paper_bottom2") & ""

        'CSS設定
        Dim sHtml As String = "<style type='text/css'>" & _
                "body{margin-left: 0px;margin-top:15px;margin-right:0px;margin-bottom:0px;}" & _
                ".banner{font-family:'標楷體';font-size:large;color:#0042AD;}" & _
                ".menu{font-size:medium;font-family:'微軟正黑體';padding-bottom:15px;}" & _
                ".menu-t{font-size: larger;font-family: '微軟正黑體';color:orange;line-height:30px;font-weight:bold}" & _
                ".source{font-size:medium;color:#000099;text-align: center;}" & _
                ".paragraph{font-family:'微軟正黑體';font-size:medium;font-weight:400;line-height:normal;}" & _
                ".paragraph-red{font-family:'微軟正黑體';font-size:medium;font-weight:400;color:red}" & _
                ".caption{font-family: '微軟正黑體';color:#0099CC;font-weight:bold}" & _
                ".sub-title1{font-family:'微軟正黑體';font-size:larger;font-weight:600;color:maroon;text-align:center;}" & _
                ".sub-title2{font-family:'微軟正黑體';font-size:larger;font-weight:600;color:red;text-align:center;}" & _
                ".title1{font-family:微軟正黑體;font-size:x-large;color:navy;font-weight:600;line-height:normal;text-decoration:underline}" & _
                ".title2{font-family:微軟正黑體;color:red;font-weight:bold;font-size:larger}" & _
                ".title3{font-family:微軟正黑體;color:blue;font-weight:bold;}" & _
                ".title4{font-family:微軟正黑體;color:olive;font-weight:bold;}" & _
                ".title5{font-family:微軟正黑體;color:#336699;font-weight:bold;}" & _
                ".title6{font-family:微軟正黑體;color:#FF3300;font-weight:bold;}" & _
                ".title7{font-family:微軟正黑體;font-size:large;color:orange;font-weight:600;line-height:normal;}" & _
                ".link1{color:blue}" & _
                ".link2{color:olive}" & _
                ".link3{color:black}" & _
                ".link4{color:yellow;text-decoration:none}" & _
                ".botton{background-color:#99CCFF;border:5px blue inset;padding:5px;color:maroon;font-weight:bold;}" & _
                "table{border-collapse:collapse;border:0px;border-spacing:0px 0px;empty-cells:hide;padding:0px}" & _
                ".style1{border-width:0px;width: 300px;height:60px}" & _
                ".style6{color:blue;font-family:微軟正黑體;font-weight:bold;}" & _
                ".style17{font-family:微軟正黑體;color:black;;font-size:larger}" & _
                ".style19{font-family: 'Times New Roman';}" & _
                ".style20{background-color:blue;border:5px blue inset;padding:5px;color:yellow;font-weight:bold;font-family:新細明體;}" & _
                ".style22{font-family: 微軟正黑體;color:green;font-weight:bold;font-size:larger;}" & _
                ".style26{color:maroon;}" & _
                ".style29{border:1px solid #00FFFF;}" & _
                ".style36{font-family: 微軟正黑體;font-weight: bold;}" & _
                ".style37{color: #FF0000;}" & _
                ".style38{margin-bottom:0px;}" & _
                ".hr-style{background: #990000;padding: 0px;height: 74px;background: url(/images/b.gif);}" & _
                "</style>"

        '抬頭圖式
        sHtml &= "<table cellpadding=0 cellspacing=0 border=0  style='position:relative; margin-left:auto; margin-right:auto; width:990px;' summary='排版表格'>" & _
                "<tr align=center><td><img src='" & sWebLogo & sPaperLogo1 & "' alt='勞動檢查處刊頭橫幅' width='750' height='150'></td></tr>"

        '電子報明細-最新消息
        Dim dtPaperDetail01 As DataTable = p1.Get_DataSet(s_prgcode, "PaperDetail01", txt_sch_code.Text).Tables(0)
        If dtPaperDetail01.Rows.Count > 0 Then
            sHtml &= "<tr><td><p class='title1'>最新訊息</p></td></tr>"

            For i As Integer = 0 To dtPaperDetail01.Rows.Count - 1
                Dim sNewsTtitle As String = dtPaperDetail01.Rows(i).Item("news_title") & ""
                Dim sNewsContect As String = dtPaperDetail01.Rows(i).Item("news_contect") & ""
                Dim sNewsPic1 As String = dtPaperDetail01.Rows(i).Item("news_pic1") & ""

                sHtml &= "<tr><td><br>"
                '新聞標題
                If sNewsContect <> "" Then
                    sHtml &= "<p class='title2'>" & sNewsTtitle & "</p>"
                End If

                If sNewsPic1 <> "" Then
                    '照片
                    sHtml &= "<table cellspacing=0 style= 'width:276px;height:180px;' cellpadding=3 align=left cellspacing=0><tr><td>" & _
                            "<img alt='照片' src='" & sWebPic & sNewsPic1 & "' height=216 width=313></td></tr></table>"
                End If
                '內文
                sHtml &= "<p>" & Replace(sNewsContect, "", "&nbsp;") & "</p></td></tr>"
                '分割線
                sHtml &= "<tr><td><hr size=1 noshade style='border:1px dashed #cccccc;'><td></tr>"
            Next
        End If

        '電子報明細-職災省思
        Dim dtPaperDetail06 As DataTable = p1.Get_DataSet(s_prgcode, "PaperDetail06", txt_sch_code.Text).Tables(0)
        If dtPaperDetail06.Rows.Count > 0 Then
            sHtml &= "<tr><td><p class='title1'>職災省思</p></td></tr>"

            For i As Integer = 0 To dtPaperDetail06.Rows.Count - 1
                Dim sNewsTtitle As String = dtPaperDetail06.Rows(i).Item("news_title") & ""
                Dim sNewsContect As String = dtPaperDetail06.Rows(i).Item("news_contect") & ""
                Dim sNewsPic1 As String = dtPaperDetail06.Rows(i).Item("news_pic1") & ""

                sHtml &= "<tr><td><br>"
                '新聞標題
                If sNewsContect <> "" Then
                    sHtml &= "<p class='title2'>" & sNewsTtitle & "</p>"
                End If

                If sNewsPic1 <> "" Then
                    '照片
                    sHtml &= "<table cellspacing=0 style= 'width:153px;height:116px;' cellpadding=3 align=left cellspacing=0><tr><td>" & _
                            "<img alt='照片' src='" & sWebPic & sNewsPic1 & "' height=102 width=153></td></tr></table>"
                End If
                '內文
                sHtml &= "<p>" & Replace(sNewsContect, "", "&nbsp;") & "</p></td></tr>"
                sHtml &= "<tr><td><hr size=1 noshade style='border:1px dashed #cccccc;'><td></tr>"
            Next
        End If

        '電子報明細-交流園地
        Dim dtPaperDetail02 As DataTable = p1.Get_DataSet(s_prgcode, "PaperDetail02", txt_sch_code.Text).Tables(0)
        If dtPaperDetail02.Rows.Count > 0 Then
            sHtml &= "<tr><td><p class='title1'>交流園地</p></td></tr>"

            For i As Integer = 0 To dtPaperDetail02.Rows.Count - 1
                Dim sNewsTtitle As String = dtPaperDetail02.Rows(i).Item("news_title") & ""
                Dim sNewsContect As String = dtPaperDetail02.Rows(i).Item("news_contect") & ""
                Dim sNewsPic1 As String = dtPaperDetail02.Rows(i).Item("news_pic1") & ""

                sHtml &= "<tr><td><br>"
                '新聞標題
                If sNewsContect <> "" Then
                    sHtml &= "<p class='style17'><span class='style26'>" & sNewsTtitle & "</span></p>"
                End If

                If sNewsPic1 <> "" Then
                    '照片
                    sHtml &= "<table cellspacing=0 style= 'width:155px;height:102px;' cellpadding=3 align=left cellspacing=0><tr><td>" & _
                            "<img alt='照片' src='" & sWebPic & sNewsPic1 & "' height=102 width=153></td></tr></table>"
                End If
                '內文
                sHtml &= "<p>" & Replace(sNewsContect, "", "&nbsp;") & "</p></td></tr>"
                sHtml &= "<tr><td><hr size=1 noshade style='border:1px dashed #cccccc;'><td></tr>"
            Next
        End If

        '電子報明細-法規新訊
        Dim dtPaperDetail03 As DataTable = p1.Get_DataSet(s_prgcode, "PaperDetail03", txt_sch_code.Text).Tables(0)
        If dtPaperDetail03.Rows.Count > 0 Then
            sHtml &= "<tr><td><p class='title1'>法規新訊</p></td></tr>"

            For i As Integer = 0 To dtPaperDetail03.Rows.Count - 1
                Dim sNewsTtitle As String = dtPaperDetail03.Rows(i).Item("news_title") & ""
                Dim sNewsContect As String = dtPaperDetail03.Rows(i).Item("news_contect") & ""
                Dim sNewsPic1 As String = dtPaperDetail03.Rows(i).Item("news_pic1") & ""

                sHtml &= "<tr><td><br>"
                '新聞標題
                If sNewsContect <> "" Then
                    sHtml &= "<p class='style22'>" & sNewsTtitle & "</p>"
                End If

                If sNewsPic1 <> "" Then
                    '照片
                    sHtml &= "<table cellspacing=0 style= 'width:153px;height:102px;' cellpadding=3 align=left cellspacing=0><tr><td>" & _
                            "<img alt='照片' src='" & sWebPic & sNewsPic1 & "' height=102 width=153></td></tr></table>"
                End If
                '內文
                sHtml &= "<p>" & Replace(sNewsContect, "", "&nbsp;") & "</p></td></tr>"
                sHtml &= "<tr><td><hr size=1 noshade style='border:1px dashed #cccccc;'><td></tr>"
            Next
        End If

        '電子報明細-訊息報導
        Dim dtPaperDetail04 As DataTable = p1.Get_DataSet(s_prgcode, "PaperDetail04", txt_sch_code.Text).Tables(0)
        If dtPaperDetail04.Rows.Count > 0 Then
            sHtml &= "<tr><td><p class='title1'>訊息報導</p></td></tr>"

            For i As Integer = 0 To dtPaperDetail04.Rows.Count - 1
                Dim sNewsTtitle As String = dtPaperDetail04.Rows(i).Item("news_title") & ""
                Dim sNewsContect As String = dtPaperDetail04.Rows(i).Item("news_contect") & ""
                Dim sNewsPic1 As String = dtPaperDetail04.Rows(i).Item("news_pic1") & ""

                sHtml &= "<tr><td><br>"
                '新聞標題
                If sNewsContect <> "" Then
                    sHtml &= "<p class='style22'>" & sNewsTtitle & "</p>"
                End If

                If sNewsPic1 <> "" Then
                    '照片
                    sHtml &= "<table cellspacing=0 style= 'width:153px;height:102px;' cellpadding=3 align=left cellspacing=0><tr><td>" & _
                            "<img alt='照片' src='" & sWebPic & sNewsPic1 & "' height=102 width=153></td></tr></table>"
                End If
                '內文
                sHtml &= "<p>" & Replace(sNewsContect, "", "&nbsp;") & "</p></td></tr>"
                sHtml &= "<tr style='height:1px'><td><hr size=1 noshade style='border:1px dashed #cccccc;'><td></tr>"
            Next
        End If

        '電子報明細-充電小站
        Dim dtPaperDetail05 As DataTable = p1.Get_DataSet(s_prgcode, "PaperDetail05", txt_sch_code.Text).Tables(0)
        If dtPaperDetail05.Rows.Count > 0 Then
            sHtml &= "<tr><td><p class='title1'>充電小站</p><p></p></td></tr>"

            For i As Integer = 0 To dtPaperDetail05.Rows.Count - 1
                Dim sNewsTtitle As String = dtPaperDetail05.Rows(i).Item("news_title") & ""
                Dim sNewsContect As String = dtPaperDetail05.Rows(i).Item("news_contect") & ""
                Dim sNewsPic1 As String = dtPaperDetail05.Rows(i).Item("news_pic1") & ""

                sHtml &= "<tr><td>"
                '新聞標題
                'If sNewsContect <> "" Then
                '    sHtml &= "<p class='style22'>" & sNewsTtitle & "</p>"
                'End If

                If sNewsPic1 <> "" Then
                    '照片
                    sHtml &= "<table cellspacing=0 style= 'width:102;height:102px;' cellpadding=3 align=left cellspacing=0><tr><td >" & _
                            "<img alt='照片' src='" & sWebPic & sNewsPic1 & "' height=102 width=153></td></tr></table>"
                End If
                '內文
                sHtml &= "<p>" & Replace(sNewsContect, "", "&nbsp;") & "</p></td></tr>"
                sHtml &= "<tr style='height:1px'><td><hr size=1 noshade style='border:1px dashed #cccccc;'><td></tr>"
            Next
        End If

        '電子報底部資訊
        sHtml &= "<tr align=center><td><br>" & sPaperBottom1 & "</td></tr><tr align=center><td>" & sPaperBottom2 & "</td></tr></table>"
        p1.SaveEditData(s_prgcode, "PaperHtml", txt_sch_code.Text, sHtml)
    End Sub


    '程式固定區-不要修改 以上
End Class

