
Partial Class basic_LSI008A_Default
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "LSI008A"
    Dim s_prgname As String = ws.Get_PrgName(s_prgcode)
    Dim UserLimit1 As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) Then
            Set_DDL_Option("", "3A", ddl_room_code, Get_SqlStr(s_prgcode, "RoomCode"), "---請選擇---", "")
            Set_DDLyyyymm()
        End If

        '程式變數設定
        UserLimit1 = ws.Get_limit(Session("usr_code"), s_prgcode)
        setLimit(UserLimit1)

        '程式抬頭
        Label1.Text = s_prgname & "(" & s_prgcode & ")"
    End Sub

    Protected Sub Page_LoadComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LoadComplete
        MeetingOrder1.GridViewDataBind()
        MeetingOrder2.GridViewDataBind()
        MeetingOrder3.GridViewDataBind()
    End Sub

    Protected Sub setLimit(ByVal Climit As String)
        '依據權限控制物件
        MeetingOrder1.Visible = CBool(InStr(Climit, "S") > 0)
        MeetingOrder2.Visible = CBool(InStr(Climit, "S") > 0)
        MeetingOrder3.Visible = CBool(InStr(Climit, "S") > 0)
    End Sub

    Protected Sub Set_DDLyyyymm()
        Dim sRange As String = ws.Get_SysPara("LSI008A_Month", "12")
        For i As Integer = -2 To val(sRange)
            ddl_yyyymm.Items.Add(New ListItem("西元 " & Format(Now.AddMonths(i), "yyyy") & " 年 " & Format(Now.AddMonths(i), "MM") & " 月"))
        Next
        ddl_yyyymm.SelectedIndex = 2
        '設定月份抬頭
        l_MeetingOrder1.Text = "西元 " & Format(Now, "yyyy") & " 年 " & Format(Now, "MM") & " 月"
        l_MeetingOrder2.Text = "西元 " & Format(Now.AddMonths(1), "yyyy") & " 年 " & Format(Now.AddMonths(1), "MM") & " 月"
        l_MeetingOrder3.Text = "西元 " & Format(Now.AddMonths(2), "yyyy") & " 年 " & Format(Now.AddMonths(2), "MM") & " 月"
        '設定usercontrol
        MeetingOrder1.set_yyyymm(Format(Now, "yyyy/MM"), ddl_room_code.SelectedValue)
        MeetingOrder2.set_yyyymm(Format(Now.AddMonths(1), "yyyy/MM"), ddl_room_code.SelectedValue)
        MeetingOrder3.set_yyyymm(Format(Now.AddMonths(2), "yyyy/MM"), ddl_room_code.SelectedValue)
    End Sub

    Protected Sub ddl_yyyymm_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddl_yyyymm.SelectedIndexChanged
        '設定月份抬頭
        l_MeetingOrder1.Text = "西元 " & Format(Now.AddMonths(ddl_yyyymm.SelectedIndex - 2), "yyyy") & " 年 " & Format(Now.AddMonths(ddl_yyyymm.SelectedIndex - 2), "MM") & " 月"
        l_MeetingOrder2.Text = "西元 " & Format(Now.AddMonths(ddl_yyyymm.SelectedIndex - 1), "yyyy") & " 年 " & Format(Now.AddMonths(ddl_yyyymm.SelectedIndex - 1), "MM") & " 月"
        l_MeetingOrder3.Text = "西元 " & Format(Now.AddMonths(ddl_yyyymm.SelectedIndex - 0), "yyyy") & " 年 " & Format(Now.AddMonths(ddl_yyyymm.SelectedIndex - 0), "MM") & " 月"
        '設定usercontrol
        MeetingOrder1.set_yyyymm(Format(Now.AddMonths(ddl_yyyymm.SelectedIndex - 2), "yyyy/MM"), ddl_room_code.SelectedValue)
        MeetingOrder2.set_yyyymm(Format(Now.AddMonths(ddl_yyyymm.SelectedIndex - 1), "yyyy/MM"), ddl_room_code.SelectedValue)
        MeetingOrder3.set_yyyymm(Format(Now.AddMonths(ddl_yyyymm.SelectedIndex - 0), "yyyy/MM"), ddl_room_code.SelectedValue)
    End Sub

    Protected Sub ddl_room_code_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddl_room_code.SelectedIndexChanged
        '設定usercontrol
        MeetingOrder1.set_yyyymm(Format(Now.AddMonths(ddl_yyyymm.SelectedIndex - 2), "yyyy/MM"), ddl_room_code.SelectedValue)
        MeetingOrder2.set_yyyymm(Format(Now.AddMonths(ddl_yyyymm.SelectedIndex - 1), "yyyy/MM"), ddl_room_code.SelectedValue)
        MeetingOrder3.set_yyyymm(Format(Now.AddMonths(ddl_yyyymm.SelectedIndex - 0), "yyyy/MM"), ddl_room_code.SelectedValue)
    End Sub

    Private Sub bt_1_click(sender As Object, e As EventArgs) Handles bt_1.Click
        Panel1.Visible = True
        Panel1_1.Visible = True
        Panel2.Visible = False
        Panel2_1.Visible = False
        Panel3.Visible = False
        Panel3_1.Visible = False

    End Sub

    Private Sub bt_2_click(sender As Object, e As EventArgs) Handles bt_2.Click
        Panel1.Visible = True
        Panel1_1.Visible = True

        Panel2.Visible = True
        Panel2_1.Visible = True

        Panel3.Visible = False
        Panel3_1.Visible = False

    End Sub

    Private Sub bt_3_click(sender As Object, e As EventArgs) Handles bt_3.Click
        Panel1.Visible = True
        Panel1_1.Visible = True

        Panel2.Visible = True
        Panel2_1.Visible = True

        Panel3.Visible = True
        Panel3_1.Visible = True

    End Sub
End Class
