Partial Class usercontrol_WebSubMenu
    Inherits System.Web.UI.UserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) Then
            setDDL()
        End If
    End Sub

    Public Sub setName(ByVal sGridViewName As String, ByVal sSqlDSName As String, ByVal sOrderTBName As String, ByVal sSearchTBName As String)
        txtGridViewName.Text = sGridViewName
        txtSqlDSName.Text = sSqlDSName
        txtOrderTBName.Text = sOrderTBName
        txtSearchTBName.Text = sSearchTBName
    End Sub

    Public Sub show()
        If Panel1.Visible = True Then Panel1.Visible = False Else Panel1.Visible = True
    End Sub

    Private Sub setDDL()
        '設定DDL
        Dim p1 As New PageBase
        '排序欄位
        setFieldDDL(ddl_field1)
        p1.CopyDDL(ddl_field1, ddl_field2)
        p1.CopyDDL(ddl_field1, ddl_field3)
        p1.CopyDDL(ddl_field1, ddl_field4)
        p1.CopyDDL(ddl_field1, ddl_field5)
        p1.CopyDDL(ddl_field1, ddl_field6)
        '排序選項
        p1.Set_DDL_Option("Order", "", ddl_order1, "", "", "")
        p1.CopyDDL(ddl_order1, ddl_order2)
        p1.CopyDDL(ddl_order1, ddl_order3)
        p1.CopyDDL(ddl_order1, ddl_order4)
        p1.CopyDDL(ddl_order1, ddl_order5)
        p1.CopyDDL(ddl_order1, ddl_order6)
        '搜尋欄位
        Dim tmpSqlDS As SqlDataSource = CType(Me.Parent.Parent.FindControl(txtSqlDSName.Text), SqlDataSource)
        ddl_search1.DataSourceID = tmpSqlDS.ID
        ddl_search2.DataSourceID = tmpSqlDS.ID
        ddl_search3.DataSourceID = tmpSqlDS.ID
        ddl_search4.DataSourceID = tmpSqlDS.ID
        ddl_search5.DataSourceID = tmpSqlDS.ID
        ddl_search6.DataSourceID = tmpSqlDS.ID
        '數學符號
        p1.Set_DDL_Option("Math_s", "", ddl_Math_s1, "", "", "")
        p1.CopyDDL(ddl_Math_s1, ddl_Math_s2)
        p1.CopyDDL(ddl_Math_s1, ddl_Math_s3)
        p1.CopyDDL(ddl_Math_s1, ddl_Math_s4)
        p1.CopyDDL(ddl_Math_s1, ddl_Math_s5)
        p1.CopyDDL(ddl_Math_s1, ddl_Math_s6)
        '布林運算
        p1.Set_DDL_Option("bool", "", ddl_bool2, "", "", "")
        p1.CopyDDL(ddl_bool2, ddl_bool1)
        p1.CopyDDL(ddl_bool2, ddl_bool3)
        p1.CopyDDL(ddl_bool2, ddl_bool4)
        p1.CopyDDL(ddl_bool2, ddl_bool5)
        p1.CopyDDL(ddl_bool2, ddl_bool6)
    End Sub

    Private Sub setFieldDDL(ByVal ddl As DropDownList)
        Dim tmpGridView As GridView = CType(Me.Parent.Parent.FindControl(txtGridViewName.Text), GridView)

        '設定排序欄位DDL
        ddl.Items.Add(New ListItem(" ", ""))
        For i As Integer = 0 To tmpGridView.Columns.Count - 1
            If tmpGridView.Columns(i).SortExpression <> "" Then
                ddl.Items.Add(New ListItem(tmpGridView.Columns(i).HeaderText, tmpGridView.Columns(i).SortExpression))
            End If
        Next
    End Sub

    Private Sub Bt_ok_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Bt_ok.Click
        Dim b_first As Boolean
        Dim tmpTBox As TextBox
        '設定排序條件
        b_first = True
        tmpTBox = CType(Me.Parent.Parent.FindControl(txtOrderTBName.Text), TextBox)
        tmpTBox.Text = ""

        setOrder(ddl_field1.SelectedValue, ddl_order1.SelectedValue, b_first, tmpTBox)
        setOrder(ddl_field2.SelectedValue, ddl_order2.SelectedValue, b_first, tmpTBox)
        setOrder(ddl_field3.SelectedValue, ddl_order3.SelectedValue, b_first, tmpTBox)
        setOrder(ddl_field4.SelectedValue, ddl_order4.SelectedValue, b_first, tmpTBox)
        setOrder(ddl_field5.SelectedValue, ddl_order5.SelectedValue, b_first, tmpTBox)
        setOrder(ddl_field6.SelectedValue, ddl_order6.SelectedValue, b_first, tmpTBox)

        '設定搜尋條件
        tmpTBox = CType(Me.Parent.Parent.FindControl(txtSearchTBName.Text), TextBox)
        If txt_Search1.Text <> "" Or txt_Search2.Text <> "" Or txt_Search3.Text <> "" Or txt_Search4.Text <> "" Or txt_Search5.Text <> "" Or txt_Search6.Text <> "" Then
            b_first = True
            setSearch(ddl_bool1.SelectedValue, ddl_search1.SelectedValue, ddl_Math_s1.SelectedValue, txt_Search1.Text, b_first, tmpTBox)
            setSearch(ddl_bool2.SelectedValue, ddl_search2.SelectedValue, ddl_Math_s2.SelectedValue, txt_Search2.Text, b_first, tmpTBox)
            setSearch(ddl_bool3.SelectedValue, ddl_search3.SelectedValue, ddl_Math_s3.SelectedValue, txt_Search3.Text, b_first, tmpTBox)
            setSearch(ddl_bool4.SelectedValue, ddl_search4.SelectedValue, ddl_Math_s4.SelectedValue, txt_Search4.Text, b_first, tmpTBox)
            setSearch(ddl_bool5.SelectedValue, ddl_search5.SelectedValue, ddl_Math_s5.SelectedValue, txt_Search5.Text, b_first, tmpTBox)
            setSearch(ddl_bool6.SelectedValue, ddl_search6.SelectedValue, ddl_Math_s6.SelectedValue, txt_Search6.Text, b_first, tmpTBox)
        Else
            tmpTBox.Text = " WHERE 1=1"
        End If

        Dim tmpGridView As GridView = CType(Me.Parent.Parent.FindControl(txtGridViewName.Text), GridView)
        tmpGridView.Sort("", SortDirection.Ascending)
        tmpGridView.DataBind()
        tmpGridView.DataBind()
    End Sub

    Private Sub setOrder(ByVal sField As String, ByVal sOrder As String, ByRef bFirst As Boolean, ByRef tmpTBox As TextBox)
        If sField <> "" And sOrder <> "" And InStr(tmpTBox.Text, sField & " ") <= 0 Then
            If bFirst = True Then
                tmpTBox.Text = " ORDER BY " & sField & " " & sOrder
                bFirst = False
            Else
                tmpTBox.Text &= "," & sField & " " & sOrder
            End If
        End If
    End Sub

    Private Sub setSearch(ByVal sBool As String, ByVal sField As String, ByVal sMath_s As String, ByVal sSearch As String, ByRef bFirst As Boolean, ByRef tmpTBox As TextBox)
        If sSearch <> "" Then
            Dim p1 As New PageBase
            sSearch = p1.RepSql(sSearch)
            If bFirst = True Then
                If sMath_s <> "LIKE" Then
                    tmpTBox.Text = " WHERE " & sField & " " & sMath_s & " N'" & sSearch & "' "
                Else
                    tmpTBox.Text = " WHERE " & sField & " " & sMath_s & " N'%" & sSearch & "%' "
                End If
                bFirst = False
            Else
                If sMath_s <> "LIKE" Then
                    tmpTBox.Text &= " " & sBool & " " & sField & " " & sMath_s & " N'" & sSearch & "' "
                Else
                    tmpTBox.Text &= " " & sBool & " " & sField & " " & sMath_s & " N'%" & sSearch & "%' "
                End If
            End If
        End If
    End Sub

    Private Sub Bt_no_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Bt_no.Click
        Panel1.Visible = False
    End Sub
End Class
