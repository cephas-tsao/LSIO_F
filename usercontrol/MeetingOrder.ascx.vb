Imports System.Drawing
Partial Class usercontrol_MeetingOrder
    Inherits System.Web.UI.UserControl

    Dim p1 As New PageBase
    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI008A"

    Public Sub set_yyyymm(ByVal yyyymm As String, ByVal room_code As String)
        txt_yyyymm.Text = yyyymm
        txt_room_code.Text = room_code
    End Sub

    Public Sub GridViewDataBind()
        GridView1.DataBind()
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) Then
            '設定Row的鍵值組成，具有唯一性
            Dim KeyNames() As String = {"mrd_code"}
            GridView1.DataKeyNames = KeyNames
        End If
        '取得資料
        setGridViewData()
    End Sub

    Protected Sub setGridViewData()
        '取得GridView資料

        '設定SqlDataSource連線及Select命令
        SqlDS1.ConnectionString = ws.Get_ConnStr("con_db")
        SqlDS1.SelectCommand = p1.Get_SqlStr(s_prgcode, "SELECT", txt_yyyymm.Text, txt_room_code.Text)

        '設定Delete命令
        Dim CtrPara As New ControlParameter("datakeys", "GridView1", "SelectedValue")
        SqlDS1.DeleteParameters.Add(CtrPara)
        SqlDS1.DeleteCommand = p1.Get_SqlStr(s_prgcode, "DELETE")

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
        '合併欄位
        Dim i0 As Integer
        Dim tmpStr1 As String

        For i As Integer = 0 To GridView1.Rows.Count - 1
            i0 = i
            tmpStr1 = GridView1.Rows(i).Cells(9).Text

            For j As Integer = i + 1 To GridView1.Rows.Count - 1
                If tmpStr1 <> GridView1.Rows(j).Cells(9).Text Then Exit For
                i += 1
            Next

            If i <> i0 Then
                Dim strRowSpan As String = (i - i0 + 1).ToString()
                GridView1.Rows(i0).Cells(9).Attributes.Add("rowspan", strRowSpan)

                For x As Integer = i0 + 1 To i
                    GridView1.Rows(x).Cells.RemoveAt(9)
                Next
            End If
        Next
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        If e.Row.RowType = DataControlRowType.Header Or e.Row.RowType = DataControlRowType.DataRow Then
            e.Row.Cells(0).Visible = False  'mrd_code
            e.Row.Cells(1).Visible = False  '起始時間
            e.Row.Cells(2).Visible = False  '中止時間
            e.Row.Cells(3).Visible = False  '日期
            e.Row.Cells(4).Visible = False  '借用人
            e.Row.Cells(5).Visible = False  '會議室名稱
            e.Row.Cells(6).Visible = False  '星期
            e.Row.Cells(7).Visible = False  '會議名稱
            e.Row.Cells(8).Visible = False  '主持人
        End If

        Select Case e.Row.RowType
            Case DataControlRowType.Header
                e.Row.Cells(9).Attributes.Add("title", "日期")
                e.Row.Cells(10).Attributes.Add("title", "會議室租借情形")
                e.Row.Cells(11).Attributes.Add("title", "修改")
                e.Row.Cells(12).Attributes.Add("title", "刪除")

            Case DataControlRowType.DataRow
                If e.Row.Cells(6).Text = "六" Or e.Row.Cells(6).Text = "日" Then
                    e.Row.Cells(9).Text = "<font color='red'>" & Right(e.Row.Cells(3).Text, 2) & "(" & e.Row.Cells(6).Text & ")</font>"
                Else
                    e.Row.Cells(9).Text = "<font color='#000044'>" & Right(e.Row.Cells(3).Text, 2) & "(" & e.Row.Cells(6).Text & ")</font>"
                End If

                If e.Row.Cells(0).Text = "&nbsp;" And e.Row.Cells(3).Text >= Format(Now, "yyyy/MM/dd") Then
                    '加入新增ImageButton
                    Dim tmpIB As New ImageButton
                    tmpIB.ToolTip = "新增租借"
                    tmpIB.Width = 32
                    tmpIB.Height = 32
                    tmpIB.ImageUrl = "~/images/icon/new3.png"
                    tmpIB.Attributes.Add("onclick", "subwindow=window.open('edit.aspx?type=ADD&date=" & e.Row.Cells(3).Text & "','MeetingOrder','height=700,width=1024,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no','');subwindow.focus();return false;")
                    e.Row.Cells(10).Controls.Add(tmpIB)
                    '清除刪除ImageButton
                    e.Row.Cells(12).Controls.Clear()
                ElseIf e.Row.Cells(0).Text = "&nbsp;" Then
                    '清除刪除ImageButton
                    e.Row.Cells(12).Controls.Clear()
                Else
                    e.Row.Cells(9).BackColor = Color.LightSkyBlue
                    e.Row.Cells(10).BackColor = Color.Wheat
                    e.Row.Cells(11).BackColor = Color.Wheat
                    e.Row.Cells(12).BackColor = Color.Wheat
                    e.Row.Cells(10).Text = "<font color='blue'>" & e.Row.Cells(1).Text & "-" & e.Row.Cells(2).Text & "</font>" & _
                                          "<font color='#663300'>[" & e.Row.Cells(4).Text & "]</font>" & _
                                          "<font color='black'>" & e.Row.Cells(5).Text & "</font>" & _
                                          "<font color='#006633'>[" & e.Row.Cells(7).Text & "]</font>" & _
                                          "<font color='#330066'>[" & e.Row.Cells(8).Text & "]</font>"
                    '加入修改ImageButton
                    Dim tmpIB As New ImageButton
                    tmpIB.ToolTip = "修改"
                    tmpIB.Width = 16
                    tmpIB.Height = 16
                    tmpIB.ImageUrl = "~/images/icon/edit2.png"
                    tmpIB.Attributes.Add("onclick", "subwindow=window.open('edit.aspx?type=MDY&date=" & e.Row.Cells(3).Text & "&key=" & e.Row.Cells(0).Text & "','MeetingOrder','height=700,width=1024,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no','');subwindow.focus();return false;")
                    e.Row.Cells(11).Controls.Add(tmpIB)
                    '清除刪除ImageButton 科室管理員及登記者可管理該筆登資料
                    If Session("usr_code") <> "WOHER" And e.Row.Cells(4).Text <> Session("usr_name") And Session("grp_code") <> "G002" And Session("grp_code") <> "G001" Then
                        e.Row.Cells(12).Controls.Clear()
                    End If
                End If
        End Select
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
    End Sub

    Protected Sub GridView1_SelectedIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSelectEventArgs) Handles GridView1.SelectedIndexChanging
        'GridView1.SelectedIndex = -1
    End Sub
End Class
