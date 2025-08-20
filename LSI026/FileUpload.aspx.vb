Imports System.Data
Imports System.Data.OleDb
Imports System.Web.Helpers

Partial Class basic_LSI026_FileUpload
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    'Dim sSavePath As String = Get_SysPara("FileUpload") & "LSI026_sample.xls"
    'Dim sSavePath As String = Get_SysPara("FileUpload") & "uploadLSI026_company.xls"
    Dim sSavePath As String = "D:\web\LSIO_WebService\new_V2\Publish\upload\File_LSI026\uploadLSI026_company.xls"

    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
        'checks database connection string and handles error if there is one
        'anticfrs.InnerHtml = AntiForgery.GetHtml().ToString()
    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) Then
            '取得變數資料
            MrdCode.Text = RepSql(Request("MrdCode"))
        End If

        'If MrdCode.Text = "" Then Response.Redirect("Default.aspx")
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        'AntiForgery.Validate()
        Dim PostedFile As HttpPostedFile = FileUpload1.PostedFile
        Dim sFileName As String = Trim(PostedFile.FileName)
        Dim sExt As String = IO.Path.GetExtension(sFileName).ToLower
        If (unit.SelectedValue = 0) Then
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('請選擇單位!!');</script>"))
        End If
        'If IO.File.Exists(sSavePath) Then IO.File.Delete(sSavePath)
        If Trim(PostedFile.FileName.ToString) <> "" Then
            If FileUpload1.HasFile = True Then
                '檢查副檔名
                If Chk_Ext(sExt) = False Then
                    Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('上傳檔案 " & sFileName & " 附檔名非允許格式,請重新選擇檔案');</script>"))
                    Exit Sub
                End If
                '上傳檔案
                Try
                    If IO.File.Exists(sSavePath) Then IO.File.Delete(sSavePath)
                    PostedFile.SaveAs(sSavePath)
                Catch ex As Exception
                    Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('上傳失敗，請重新上傳');</script>"))
                    Exit Sub
                End Try
                '匯入資料
                If Update_ExcelFile() Then
                    Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('匯入成功');</script>"))
                Else
                    Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('" & e.ToString() & "')</script>")) '檔案格式有誤，請確認 
                    Exit Sub
                End If
                '刪除檔案
                If IO.File.Exists(sSavePath) Then IO.File.Delete(sSavePath)
            Else
                Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('該路徑無檔案');</script>"))
                Exit Sub
            End If
        Else
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('請選擇欲上傳檔案');</script>"))
            Exit Sub
        End If

        Session("LSI026_MrdCode") = MrdCode.Text
        Session("UsrListDo") = "Y"
        Response.Write("<script>opener.window.location.href = opener.window.location.href;</script>") 'window.close();
    End Sub

    Protected Function Chk_Ext(ByVal sExt As String) As Boolean
        If sExt <> ".xls" And sExt <> ".xlsx" Then Return False
        Return True
    End Function

    Protected Function Update_ExcelFile() As Boolean
        '讀取Excel
        Dim v_sheet1 As String = ""
        Dim dsTmp As DataSet = New DataSet()
        Dim MyCommand As OleDbDataAdapter
        Dim MyConnection As OleDbConnection
        Dim schemaTable As DataTable

        'MyConnection = New OleDbConnection( _
        '"provider=Microsoft.ACE.OLEDB.12.0; " & _
        '"data source=" & sSavePath & "; " & _
        '"Extended Properties=Excel 12.0;")
        MyConnection = New OleDbConnection(
        "provider=Microsoft.ACE.OLEDB.12.0; " &
        "data source=" & sSavePath & "; " &
        "Extended Properties=Excel 12.0;")
        '讀取sheel
        Try
            MyConnection.Open()
            schemaTable = MyConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
            MyConnection.Close()
            v_sheet1 = schemaTable.Rows(0).Item("TABLE_NAME").ToString()
            schemaTable.Clear()

        Catch ex As Exception
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('" + ex.Tostring() + "');</script>")) '此檔案非EXCEL檔
            Return False
        End Try
        '讀取資料至DS
        MyCommand = New System.Data.OleDb.OleDbDataAdapter("SEL" & "ECT * FROM [" & v_sheet1 & "]", MyConnection)
        MyCommand.Fill(dsTmp)
        MyConnection.Close()

        '新增修改資料
        Dim usr_code As String = ""
        Dim usr_name As String = ""
        Dim usr_com As String = ""
        Dim usr_uint As String = ""
        Dim uniform_num As String = ""
        Dim start_year As String = ""
        Dim end_year As String = ""
        '計算結果
        Dim iTotalCount As Integer = 0
        Dim iFailCount As Integer = 0
        Dim sMessage As String = ""

        Try
            For Each row As DataRow In dsTmp.Tables(0).Rows
                usr_code = row.Item("序號").ToString.Trim
                usr_com = row.Item("事業單位").ToString.Trim
                If row.Item("檢查員").ToString.Trim <> "" Then usr_name = row.Item("檢查員").ToString.Trim Else usr_name = "未填寫"
                uniform_num = row.Item("統編").ToString.Trim
                start_year = row.Item("起始年度").ToString.Trim
                end_year = row.Item("結束年度").ToString.Trim
                usr_uint = unit.SelectedItem.Text
                'usr_mail = row.Item("連絡mail").ToString.Trim
                If usr_name = "" Or usr_com = "" Then Continue For
                '計算總筆數
                iTotalCount += 1

                If Chk_RelData("LSI026_com", "", usr_com, usr_uint) = True Then
                    If SaveEditData("Self_com", "ADD", usr_code, usr_com, usr_name, usr_uint, uniform_num, start_year, end_year) = False Then
                        iFailCount += 1
                        sMessage &= "第 " & iTotalCount & " 筆:資料有誤\n"
                    End If
                Else
                    iFailCount += 1
                    sMessage &= "第 " & iTotalCount & " 筆:資料重覆\n"
                End If
            Next

            sMessage &= "總共匯入: " & iTotalCount & " 件\n"
            sMessage &= "匯入成功: " & iTotalCount - iFailCount & " 件\n"
            sMessage &= "匯入失敗 " & iFailCount & " 件\n"
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('" & sMessage & "');</script>"))
        Catch ex As Exception
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('" & ex.Message.ToString() & "');</script>")) 'EXCEL檔案欄位資料有誤
            'Return False
        End Try

        Return True
    End Function
End Class
