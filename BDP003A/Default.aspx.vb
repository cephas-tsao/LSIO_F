Imports System.Data

Partial Class management_BDP003A_Default
    Inherits PageBase

    Dim s_prgcode As String = "BDP003A"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            setDDL1()
        End If
    End Sub

    Protected Sub setDDL1()
        If Not IsPostBack Then
            Dim sSql As String = Get_SqlStr(s_prgcode, "")
            Set_DDL_Option("", "", DDL1, sSql, " ", "")
        End If

    End Sub

    Protected Sub setPrgTable()
        Dim dt_tmp As DataTable = Get_DataSet(s_prgcode, "GetMenu").Tables(0)
        Dim dt_tmp2 As DataTable
        Dim limit_str As String

        Table1.BorderStyle = BorderStyle.Double
        Table1.BackColor = Drawing.Color.White
        Table1.BorderStyle = BorderStyle.Solid
        Table1.BorderColor = Drawing.Color.Black

        '第一列
        Dim NewRow As New TableRow
        Dim NewCell As TableCell
        Dim A_check As CheckBox
        Dim all_check As CheckBox
        'Dim all_button As New Button

        NewCell = New TableCell
        NewCell.Text = "程式名稱"
        NewRow.Cells.Add(NewCell)
        '增加兩欄
        NewCell = New TableCell
        NewCell.Text = ""
        NewRow.Cells.Add(NewCell)
        NewCell = New TableCell
        NewCell.Text = ""
        NewRow.Cells.Add(NewCell)

        NewCell = New TableCell
        all_check = New CheckBox
        all_check.Attributes.Add("onclick", "Chk_All_E()")
        NewCell.Controls.Add(all_check)
        NewCell.Text = "E.執行"
        NewRow.Cells.Add(NewCell)

        NewCell = New TableCell
        NewCell.Text = "A.新增"
        NewRow.Cells.Add(NewCell)

        NewCell = New TableCell
        NewCell.Text = "M.修改"
        NewRow.Cells.Add(NewCell)

        NewCell = New TableCell
        NewCell.Text = "D.刪除"
        NewRow.Cells.Add(NewCell)

        NewCell = New TableCell
        NewCell.Text = "S.閱讀"
        NewRow.Cells.Add(NewCell)

        Table1.Rows.Add(NewRow)

        '---------------------------------
        NewRow = New TableRow
        '第二列
        '增加兩欄
        NewCell = New TableCell
        NewCell.Text = ""
        NewRow.Cells.Add(NewCell)
        NewCell = New TableCell
        NewCell.Text = ""
        NewRow.Cells.Add(NewCell)
        'NewCell.HorizontalAlign = HorizontalAlign.Right
        NewCell = New TableCell
        NewCell.Text = "全選/取消"
        NewRow.Cells.Add(NewCell)

        NewCell = New TableCell
        all_check = New CheckBox
        all_check.Attributes.Add("onclick", "Chk_All_E()")
        all_check.ID = "check_box_E"
        NewCell.Controls.Add(all_check)
        'all_button = New Button
        'all_button.Attributes.Add("onclick", "event.returnValue=false;Chk_All_E()")
        'all_button.ID = "check_button_E"
        'all_button.Text = "欄全選"
        'NewCell.Controls.Add(all_button)
        NewRow.Cells.Add(NewCell)

        NewCell = New TableCell
        all_check = New CheckBox
        all_check.Attributes.Add("onclick", "Chk_All_A()")
        all_check.ID = "check_box_A"
        NewCell.Controls.Add(all_check)
        'all_button = New Button
        'all_button.Attributes.Add("onclick", "event.returnValue=false;Chk_All_A()")
        'all_button.ID = "check_button_A"
        'all_button.Text = "欄全選"
        'NewCell.Controls.Add(all_button)
        NewRow.Cells.Add(NewCell)

        NewCell = New TableCell
        all_check = New CheckBox
        all_check.Attributes.Add("onclick", "Chk_All_M()")
        all_check.ID = "check_box_M"
        NewCell.Controls.Add(all_check)
        'all_button = New Button
        'all_button.Attributes.Add("onclick", "event.returnValue=false;Chk_All_M()")
        'all_button.ID = "check_button_M"
        'all_button.Text = "欄全選"
        'NewCell.Controls.Add(all_button)
        NewRow.Cells.Add(NewCell)

        NewCell = New TableCell
        all_check = New CheckBox
        all_check.Attributes.Add("onclick", "Chk_All_D()")
        all_check.ID = "check_box_D"
        NewCell.Controls.Add(all_check)
        'all_button = New Button
        'all_button.Attributes.Add("onclick", "event.returnValue=false;Chk_All_D()")
        'all_button.ID = "check_button_D"
        'all_button.Text = "欄全選"
        'NewCell.Controls.Add(all_button)
        NewRow.Cells.Add(NewCell)

        NewCell = New TableCell
        all_check = New CheckBox
        all_check.Attributes.Add("onclick", "Chk_All_S()")
        all_check.ID = "check_box_S"
        NewCell.Controls.Add(all_check)
        'all_button = New Button
        'all_button.Attributes.Add("onclick", "event.returnValue=false;Chk_All_S()")
        'all_button.ID = "check_button_S"
        'all_button.Text = "欄全選"
        'NewCell.Controls.Add(all_button)
        NewRow.Cells.Add(NewCell)

        Table1.Rows.Add(NewRow)

        '第三列
        'NewRow = New TableRow
        'NewCell = New TableCell
        'NewRow.Cells.Add(NewCell)
        'NewCell = New TableCell
        'NewCell.Text = "　"
        'NewCell.ColumnSpan = 5
        'NewRow.Cells.Add(NewCell)
        'Table1.Rows.Add(NewRow)

        '--------------------------------


        For ilop As Integer = 0 To dt_tmp.Rows.Count - 1
            NewRow = New TableRow
            A_check = New CheckBox
            NewCell = New TableCell
            If ilop Mod 2 = 0 Then
                NewRow.BackColor = Drawing.Color.LightBlue
            Else
                NewRow.BackColor = Drawing.Color.White
            End If
            NewCell.Text = Trim(dt_tmp.Rows(ilop).Item("prg_code") & "")
            NewRow.Cells.Add(NewCell)

            NewCell = New TableCell
            NewCell.Text = Trim(dt_tmp.Rows(ilop).Item("prg_name") & "")
            NewRow.Cells.Add(NewCell)

            '取得該帳號預設權限
            dt_tmp2 = Get_DataSet(s_prgcode, "GetLimit", DDL1.SelectedValue, dt_tmp.Rows(ilop).Item("prg_code") & "").Tables(0)

            If dt_tmp2.Rows.Count > 0 Then
                limit_str = Trim(dt_tmp2.Rows(0).Item("limit_str") & "")
            Else
                limit_str = ""
            End If

            NewCell = New TableCell
            all_check = New CheckBox
            all_check.Attributes.Add("onclick", "Chk_Row('" & Trim(dt_tmp.Rows(ilop).Item("prg_code") & "") & "')")
            all_check.ID = "C|" & Trim(dt_tmp.Rows(ilop).Item("prg_code") & "")
            NewCell.Controls.Add(all_check)
            'all_button = New Button
            'all_button.Attributes.Add("onclick", "if (window.event) {window.event.returnValue = false;} else {e.returnValue = false;}Chk_Row('" & Trim(dt_tmp.rows(ilop).item("prg_code") & "") & "')")
            'all_button.ID = "BT|" & Trim(dt_tmp.rows(ilop).item("prg_code") & "")
            'all_button.Text = "行全選"
            'NewCell.Controls.Add(all_button)
            NewRow.Cells.Add(NewCell)
            'NewCell.HorizontalAlign = HorizontalAlign.Right

            A_check = New CheckBox
            NewCell = New TableCell
            NewCell.BorderStyle = BorderStyle.None
            NewCell.BorderWidth = 0
            A_check.ID = "E|" & Trim(dt_tmp.Rows(ilop).Item("prg_code") & "")
            A_check.Checked = CBool(InStr(limit_str, "E") > 0)
            NewCell.Controls.Add(A_check)
            NewRow.Cells.Add(NewCell)

            A_check = New CheckBox
            NewCell = New TableCell
            A_check.ID = "A|" & Trim(dt_tmp.Rows(ilop).Item("prg_code") & "")
            A_check.Checked = CBool(InStr(limit_str, "A") > 0)
            NewCell.Controls.Add(A_check)
            NewRow.Cells.Add(NewCell)

            A_check = New CheckBox
            NewCell = New TableCell
            A_check.ID = "M|" & Trim(dt_tmp.Rows(ilop).Item("prg_code") & "")
            A_check.Checked = CBool(InStr(limit_str, "M") > 0)
            NewCell.Controls.Add(A_check)
            NewRow.Cells.Add(NewCell)

            A_check = New CheckBox
            NewCell = New TableCell
            A_check.ID = "D|" & Trim(dt_tmp.Rows(ilop).Item("prg_code") & "")
            A_check.Checked = CBool(InStr(limit_str, "D") > 0)
            NewCell.Controls.Add(A_check)
            NewRow.Cells.Add(NewCell)

            A_check = New CheckBox
            NewCell = New TableCell
            A_check.ID = "S|" & Trim(dt_tmp.Rows(ilop).Item("prg_code") & "")
            A_check.Checked = CBool(InStr(limit_str, "S") > 0)
            NewCell.Controls.Add(A_check)
            NewRow.Cells.Add(NewCell)

            Table1.Rows.Add(NewRow)
        Next

    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        If DDL1.SelectedValue = "" Then
            Table1.Visible = False
            Button1.Visible = False
            Button3.Visible = False
            Button4.Visible = False
            Bt_save.Visible = False
            Bt_cancel.Visible = False
            Exit Sub
        End If

        setPrgTable()
        DDL1.Enabled = False
        Table1.Visible = True
        Button1.Visible = True
        Button3.Visible = True
        Button4.Visible = True
        Bt_save.Visible = True
        Bt_cancel.Visible = True
    End Sub

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        DDL1.Enabled = True
        setPrgTable()
    End Sub

    Protected Sub Button_Cancel(ByVal sender As Object, ByVal e As System.EventArgs)
        Table1.Visible = False
        Button1.Visible = False
        Button3.Visible = False
        Button4.Visible = False
        Bt_save.Visible = False
        Bt_cancel.Visible = False
        DDL1.Enabled = True
    End Sub

    Protected Sub Save_User_Data(ByVal sender As Object, ByVal e As System.EventArgs) Handles Bt_save.Click
        Dim dt_tmp As DataTable = Get_DataSet(s_prgcode, "GetMenu").Tables(0)
        Dim i_prgcode As String
        Dim limit_str As String

        DDL1.Enabled = True
        If DDL1.SelectedValue <> "" Then

            ''先刪後增
            SaveEditData(s_prgcode, "DELETE", DDL1.SelectedValue)

            '寫入程式權限
            For xlop As Integer = 0 To dt_tmp.Rows.Count - 1
                i_prgcode = Trim(dt_tmp.Rows(xlop).Item("prg_code") & "")
                limit_str = Get_Str(i_prgcode)
                SaveEditData(s_prgcode, "INSERT", DDL1.SelectedValue, i_prgcode, limit_str, "Y")
            Next

            '取得帳號權限設定
            dt_tmp = Get_DataSet(s_prgcode, "GetType", DDL1.SelectedValue).Tables(0)
            If dt_tmp.Rows.Count > 0 Then
                If dt_tmp.Rows(0).Item(0).ToString = "B" Then
                    Me.Page.Form.Controls.Add(New LiteralControl("<script>Chk_Limit_Type();</script>"))
                Else
                    Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('權限變更完成');</script>"))
                End If
            Else
                Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('無此帳號!');</script>"))
            End If
        Else
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('內容不正確');</script>"))
            Table1.Visible = False
        End If

        setPrgTable()
        Button1.Visible = False
        Button3.Visible = False
        Button4.Visible = False
        Bt_save.Visible = False
        Bt_cancel.Visible = False
    End Sub

    Protected Function Get_Str(ByVal cPrgCode As String) As String
        '將上頁的設定串成權限字串
        Dim Tmp_str As String = ""

        If Me.Request.Form("A|" & cPrgCode) = "on" Then
            Tmp_str = Tmp_str & "A"
        End If
        If Me.Request.Form("M|" & cPrgCode) = "on" Then
            Tmp_str = Tmp_str & "M"
        End If
        If Me.Request.Form("D|" & cPrgCode) = "on" Then
            Tmp_str = Tmp_str & "D"
        End If
        If Me.Request.Form("S|" & cPrgCode) = "on" Then
            Tmp_str = Tmp_str & "S"
        End If
        If Me.Request.Form("E|" & cPrgCode) = "on" Then
            Tmp_str = Tmp_str & "E"
        End If

        Return Tmp_str
    End Function

    Protected Sub Button_Hidden_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button_Hidden.Click
        SaveEditData(s_prgcode, "UPDATE", DDL1.SelectedValue)
        setPrgTable()
    End Sub

End Class
