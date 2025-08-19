Imports System.Data
Imports System.Drawing

Partial Class basic_LSI008A_UsrList
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim tmpDB As DataTable
    Dim s_prgcode As String = "addList"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) Then
            '取得變數資料
            MrdCode.Text = RepSql(Request("MrdCode"))
            tmpDB = Get_DataSet("", "SqlStr", Get_SqlStr("LSI008A_detail", "SELECT", MrdCode.Text)).Tables(0)

            '設定查詢欄位
            SqlDS2.ConnectionString = ws.Get_ConnStr("con_db")
            SqlDS2.SelectCommand = Get_SqlStr(s_prgcode, "", "UsrList")

            '取得資料
            setCheckBoxListData()
            setChecked()
        End If

        If MrdCode.Text = "" Then Response.Redirect("Default.aspx")
    End Sub

    Protected Sub setCheckBoxListData()
        '取得CheckBoxList資料
        Dim sSql As String = ""
        Dim dt_tmp1 As DataTable = Get_DataSet(s_prgcode, "", "UsrList").Tables(0)

        If txtSubWhere.Text <> "" Then
            If InStr(Trim(dt_tmp1.Rows(0).Item("sel_str") & ""), "WHERE") > 0 Then
                sSql = Trim(dt_tmp1.Rows(0).Item("sel_str") & "") & Replace(txtSubWhere.Text, " WHERE ", " AND ")
            Else
                sSql = Trim(dt_tmp1.Rows(0).Item("sel_str") & "") & txtSubWhere.Text
            End If
        Else
            sSql = Trim(dt_tmp1.Rows(0).Item("sel_str") & "")
        End If

        Dim sOrder = RepSql(dt_tmp1.Rows(0).Item("order_str"))
        If sOrder <> "" Then sSql &= " ORDER BY " & sOrder

        Dim dt_tmp2 As DataTable = Get_DataSet(s_prgcode, "SqlStr", sSql).Tables(0)
        '設定CheckBoxList資料來源
        CheckBoxList1.Dispose()
        CheckBoxList1.DataSource = dt_tmp2
        '設定CheckBoxList欄數
        If dt_tmp2.Rows.Count > 19 Then CheckBoxList1.RepeatColumns = 2
        If dt_tmp2.Rows.Count > 37 Then CheckBoxList1.RepeatColumns = 3
        CheckBoxList1.DataValueField = dt_tmp1.Rows(0).Item("list_value")
        CheckBoxList1.DataTextField = dt_tmp1.Rows(0).Item("list_name")
        CheckBoxList1.DataBind()
    End Sub

    Protected Sub setChecked()
        For i As Integer = 0 To tmpDB.Rows.Count - 1
            For j As Integer = 0 To CheckBoxList1.Items.Count - 1
                If CheckBoxList1.Items(j).Value = tmpDB.Rows(i).Item("usr_code") Then
                    CheckBoxList1.Items(j).Selected = True
                    Exit For
                End If
            Next
        Next
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
        tmpDB = Get_DataSet("", "SqlStr", Get_SqlStr("LSI008A_detail", "SELECT", MrdCode.Text)).Tables(0)

        For i As Integer = 0 To CheckBoxList1.Items.Count - 1

            If CheckBoxList1.Items(i).Selected = True Then
                Dim bSaved As Boolean = False

                For j As Integer = 0 To tmpDB.Rows.Count - 1
                    If CheckBoxList1.Items(i).Value = tmpDB.Rows(j).Item("usr_code") Then
                        bSaved = True
                        Exit For
                    End If
                Next

                If bSaved = False Then
                    Dim dt_usr As DataTable = Get_DataSet("UsrList", "", CheckBoxList1.Items(i).Value).Tables(0)

                    If dt_usr.Rows.Count > 0 Then SaveEditData("LSI008A_detail", "ADD", "", MrdCode.Text, _
                                                  dt_usr.Rows(0).Item("usr_code") & "", _
                                                  dt_usr.Rows(0).Item("usr_name") & "", _
                                                  dt_usr.Rows(0).Item("usr_mail") & "", _
                                                  dt_usr.Rows(0).Item("dep_name") & "")
                End If
            End If
        Next

        Session("LSI008A_MrdCode") = MrdCode.Text
        'If Session("UsrListDo") = "Y" Then Session("UsrListDo") = "N"
        Session("UsrListDo") = "Y"
        Response.Write("<script>opener.window.location.href = opener.window.location.href;window.close();</script>")
        'Response.Write("<script>self.opener.location.reload();window.close();</script>")
    End Sub

    Protected Sub CheckBox1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        For i As Integer = 0 To CheckBoxList1.Items.Count - 1
            CheckBoxList1.Items(i).Selected = CheckBox1.Checked
        Next
    End Sub
End Class


