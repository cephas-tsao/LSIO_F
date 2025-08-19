Imports System.Data

Partial Class management_BDP220A_edit
    Inherits PageBase


    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '資料存檔
        Dim Pgm_Type As String
        Dim Sql_str As String
        Dim PrgMemo As String
        Dim i_prgcode As String

        Pgm_Type = txt_pgm_type.Text
        Sql_str = ""
        PrgMemo = ""
        i_prgcode = "BDP220A"

        '設定記錄資料
        Sql_str = "select * from BDP220 where bdp220='" & txtPrimary.Text & "'"
        Dim sFieldName As String = Get_FieldName(i_prgcode)
        Dim sOldData As String = Get_OldData(Sql_str, i_prgcode)
        Dim sNewData As String = Get_NewData(i_prgcode)

        '檢查資料正確性
        If Chk_data() = False Then Exit Sub

        Select Case Pgm_Type
            Case "ADD", "COPY"
                Sql_str = "insert into BDP220(" & _
                          "            prg_code,field_code,field_name,data_source,data_type,scr_no,is_use)" & _
                          " values(N'" & _
                          RepSql(txt_prg_code.Text) & "',N'" & _
                          RepSql(txt_field_code.Text) & "',N'" & _
                          RepSql(txt_field_name.Text) & "',N'" & _
                          RepSql(txt_data_source.Text) & "',N'" & _
                          RepSql(ddl_data_type.SelectedValue) & "',N'" & _
                          RepSql(txt_scr_no.Text) & "',N'" & _
                          RepSql(ddl_is_use.SelectedValue) & "')"
                PrgMemo = RepSql(txt_prg_code.Text)
            Case "MDY"
                Sql_str = "update BDP220 " & _
                          "   set " & _
                          "       prg_code=N'" & RepSql(txt_prg_code.Text) & "'," & _
                          "       field_code=N'" & RepSql(txt_field_code.Text) & "'," & _
                          "       field_name=N'" & RepSql(txt_field_name.Text) & "'," & _
                          "       data_source=N'" & RepSql(txt_data_source.Text) & "'," & _
                          "       data_type=N'" & RepSql(ddl_data_type.SelectedValue) & "'," & _
                          "       scr_no=N'" & RepSql(txt_scr_no.Text) & "'," & _
                          "       is_use=N'" & RepSql(ddl_is_use.SelectedValue) & "'" & _
                          " where bdp220='" & RepSql(txtPrimary.Text) & "'"
                PrgMemo = RepSql(txt_prg_code.Text)
        End Select

        If Sql_str <> "" Then
            If SaveData(Sql_str) Then
                InsUsrRec(PrgMemo, Pgm_Type, i_prgcode, sFieldName, sOldData, sNewData)
                Dim SubWhere As String
                SubWhere = ""
                If Pgm_Type = "MDY" Then
                    SubWhere = txtSubWhere.Text
                Else
                    '取得最新條件
                    SubWhere = txtSubWhere.Text & " or prg_code='" & RepSql(txt_prg_code.Text) & "' "
                End If
                '存檔正確導回default頁面
                Response.Redirect("default.aspx?subwhere=" & SubWhere)
            Else
                '存檔出錯就導向錯誤頁面
                Response.Redirect("error.aspx")
            End If
        End If
    End Sub

    Protected Function Chk_data() As Boolean
        'Server端資料檢查 自訂函數
        Chk_data = True

        '檢查重覆資料
        If txt_pgm_type.Text = "ADD" Or txt_pgm_type.Text = "COPY" Then
            If Chk_RelData("select * from BDP220 where prg_code='" & RepSql(txt_prg_code.Text) & "' and field_code='" & RepSql(txt_field_code.Text) & "'") = False Then
                Chk_data = False
                Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('有重覆編號資料，請確認');</script>"))
            End If
        End If

        '檢查必填欄位是否空白
        If txt_prg_code.Text = "" Or txt_field_code.Text = "" Or txt_data_source.Text = "" Or txt_scr_no.Text = "" Then
            Chk_data = False
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('必填欄位不得為空白，請確認');</script>"))
        End If
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '接收default頁面變數

        If Not IsPostBack Then
            '標頭
            Label1.Text = CType(PreviousPage.FindControl("Label1"), Label).Text
            '編輯模式
            txt_pgm_type.Text = CType(PreviousPage.FindControl("txtPgmType"), TextBox).Text
            '首頁條件
            txtSubWhere.Text = CType(PreviousPage.FindControl("txtsubwhere"), TextBox).Text
            '設定下拉式選單
            Set_DDL_Option("IsUse", "", ddl_is_use, "", "", "")
            Set_DDL_Option("DataType", "", ddl_data_type, "", "", "")

            '非新增模式則預設取得資料
            Dim Pgm_type As String
            Pgm_type = CType(PreviousPage.FindControl("txtPgmType"), TextBox).Text
            txtPrimary.Text = CType(PreviousPage.FindControl("txtPrimary"), TextBox).Text


            Select Case Pgm_type
                Case "MDY", "COPY"
                    Dim Sql_str As String
                    Dim Dt_tmp As DataTable

                    Sql_str = "select * from BDP220 where bdp220='" & txtPrimary.Text & "'"
                    Dt_tmp = Get_DataTable(Sql_str)
                    If Dt_tmp.Rows.Count > 0 Then
                        txt_prg_code.Text = Trim(Dt_tmp.Rows(0).Item("prg_code") & "")
                        txt_field_code.Text = Trim(Dt_tmp.Rows(0).Item("field_code") & "")
                        txt_field_name.Text = Trim(Dt_tmp.Rows(0).Item("field_name") & "")
                        txt_data_source.Text = Trim(Dt_tmp.Rows(0).Item("data_source") & "")
                        ddl_data_type.SelectedValue = Trim(Dt_tmp.Rows(0).Item("data_type") & "")
                        txt_scr_no.Text = Trim(Dt_tmp.Rows(0).Item("scr_no") & "")
                        ddl_is_use.SelectedValue = Trim(Dt_tmp.Rows(0).Item("is_use") & "")
                    End If
            End Select


        End If
    End Sub

    Protected Sub Bt_BackUp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Bt_BackUp.Click
        Response.Redirect("default.aspx?subwhere=" & txtSubWhere.Text)
    End Sub
End Class
