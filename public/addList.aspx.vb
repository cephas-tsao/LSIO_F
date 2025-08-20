Imports System.Data
Imports System.Drawing

Partial Class public_addlist
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim s_prgcode As String = "addList"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) Then
            '取得變數資料
            AddCode.Text = RepSql(Request("add_code"))
            TextBox.Text = RepSql(Request("textbox"))


            If Request.Browser.Browser = "Firefox" Then
                TextBox.Text = Replace(TextBox.Text, "$", "_")
            End If
            '設定查詢欄位
            SqlDS2.ConnectionString = ws.Get_ConnStr("con_db")
            SqlDS2.SelectCommand = Get_SqlStr(s_prgcode, "", Request("add_code"))

            If Trim(Request("subwhere") & "") <> "" Then txtSubWhere2.Text = RepSql(Replace(Request("subwhere"), "@", "'"))
            If Trim(Request("checked") & "") <> "" Then txtChecked.Text = RepSql(Replace(Request("checked"), "@", "'"))
            If Trim(Request("ChkName") & "") <> "" Then txtChkName.Text = RepSql(Replace(Request("ChkName"), "@", "'"))
            If Trim(Request("column") & "") <> "" Then txtColumn.Text = RepSql(Replace(Request("column"), "@", "'"))
            '取得資料
            setCheckBoxListData()
            setChecked()
        End If

    End Sub

    Protected Sub setCheckBoxListData()
        '取得CheckBoxList資料
        Dim sSql As String = ""
        Dim dt_tmp1 As DataTable = Get_DataSet(s_prgcode, "", AddCode.Text).Tables(0)

        If txtSubWhere.Text <> "" Then
            If txtSubWhere2.Text = "" Then
                If InStr(Trim(dt_tmp1.Rows(0).Item("sel_str") & ""), "WHERE") > 0 Then
                    sSql = Trim(dt_tmp1.Rows(0).Item("sel_str") & "") & Replace(txtSubWhere.Text, " WHERE ", " AND ")
                Else
                    sSql = Trim(dt_tmp1.Rows(0).Item("sel_str") & "") & txtSubWhere.Text
                End If
            Else
                If InStr(Trim(dt_tmp1.Rows(0).Item("sel_str") & ""), "WHERE") > 0 Then
                    sSql = Trim(dt_tmp1.Rows(0).Item("sel_str") & "") & Replace(txtSubWhere.Text, " WHERE ", " AND ") & " AND " & txtSubWhere2.Text
                Else
                    sSql = Trim(dt_tmp1.Rows(0).Item("sel_str") & "") & txtSubWhere.Text & " AND " & txtSubWhere2.Text
                End If
            End If
        Else
            If txtSubWhere2.Text = "" Then
                sSql = Trim(dt_tmp1.Rows(0).Item("sel_str") & "")
            Else
                sSql = Trim(dt_tmp1.Rows(0).Item("sel_str") & "") & " WHERE " & txtSubWhere2.Text
            End If
        End If

        Dim sOrder = RepSql(dt_tmp1.Rows(0).Item("order_str"))
        If sOrder <> "" Then sSql &= " ORDER BY " & sOrder

        '設定CheckBoxList欄數
        If txtColumn.Text <> "" Then CheckBoxList1.RepeatColumns = CInt(txtColumn.Text)

        Dim dt_tmp2 As DataTable = Get_DataSet(s_prgcode, "SqlStr", sSql).Tables(0)
        '設定CheckBoxList資料來源
        CheckBoxList1.Dispose()
        CheckBoxList1.DataSource = dt_tmp2
        If txtColumn.Text = "" Then
            '若無設定欄數則自動調整
            If dt_tmp2.Rows.Count > 19 Then CheckBoxList1.RepeatColumns = 2
            'If dt_tmp2.Rows.Count > 37 Then CheckBoxList1.RepeatColumns = 3
        End If
        CheckBoxList1.DataValueField = dt_tmp1.Rows(0).Item("list_value")
        CheckBoxList1.DataTextField = dt_tmp1.Rows(0).Item("list_name")
        CheckBoxList1.DataBind()
    End Sub

    Protected Sub setChecked()
        If txtChecked.Text <> "" Then
            Dim sub_Checked As String() = Split(txtChecked.Text, ",")
            For i As Integer = 0 To sub_Checked.Length - 1
                For j As Integer = 0 To CheckBoxList1.Items.Count - 1
                    If CheckBoxList1.Items(j).Value = sub_Checked(i) Then
                        CheckBoxList1.Items(j).Selected = True
                        Exit For
                    End If
                Next
            Next
        End If
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '資料搜尋-單一條件
        If Trim(Replace(TxtFind.Text, "'", "") & "") <> "" Then
            txtSubWhere.Text = " WHERE " & ddlField.SelectedValue & " LIKE '%" & TxtFind.Text & "%'"
        Else
            txtSubWhere.Text = " WHERE 1=1"
        End If
        setCheckBoxListData()
    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim para As String() = Split(TextBox.Text, "|")
        Dim sScript As String
        Dim sTmpStr As String()
        Dim first As Boolean
        Dim bRepeat As Boolean
        '設定顯示名稱
        first = True
        sScript = "opener.document.getElementById('" & para(0) & "').value = '"
        '附加資料
        If CheckBox2.Checked = True And txtChkName.Text <> "" Then
            sScript &= txtChkName.Text
            first = False
        End If

        For i As Integer = 0 To CheckBoxList1.Items.Count - 1
            bRepeat = False
            If CheckBoxList1.Items(i).Selected = True Then
                '檢查重覆資料
                If CheckBox2.Checked = True And txtChkName.Text <> "" Then
                    sTmpStr = Split(txtChkName.Text, ",")

                    For j As Integer = 0 To sTmpStr.Length - 1
                        If sTmpStr(j) = CheckBoxList1.Items(i).Text Then
                            bRepeat = True
                            Exit For
                        End If
                    Next
                End If
                If bRepeat = True Then Continue For '若資料重覆則跳下一筆

                If first = False Then sScript &= ","
                first = False
                sScript &= CheckBoxList1.Items(i).Text
            End If
        Next
        sScript &= "';"
        '設定隱藏code
        first = True
        sScript &= "opener.document.getElementById('" & para(1) & "').value = '"

        '附加資料
        If CheckBox2.Checked = True And txtChecked.Text <> "" Then
            sScript &= txtChecked.Text
            first = False
        End If

        For i As Integer = 0 To CheckBoxList1.Items.Count - 1
            bRepeat = False
            If CheckBoxList1.Items(i).Selected = True Then
                '檢查重覆資料
                If CheckBox2.Checked = True And txtChecked.Text <> "" Then
                    sTmpStr = Split(txtChecked.Text, ",")

                    For j As Integer = 0 To sTmpStr.Length - 1
                        If sTmpStr(j) = CheckBoxList1.Items(i).Value Then
                            bRepeat = True
                            Exit For
                        End If
                    Next
                End If
                If bRepeat = True Then Continue For '若資料重覆則跳下一筆

                If first = False Then sScript &= ","
                first = False
                sScript &= CheckBoxList1.Items(i).Value
            End If
        Next
        sScript &= "';"

        sScript &= "window.close();"
        Response.Write("<script language='javascript'>" & sScript & "</script>")
        Response.End()
    End Sub

    Protected Sub CheckBox1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        For i As Integer = 0 To CheckBoxList1.Items.Count - 1
            CheckBoxList1.Items(i).Selected = CheckBox1.Checked
        Next
    End Sub
End Class


