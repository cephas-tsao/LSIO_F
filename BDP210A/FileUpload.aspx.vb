Imports System.Data
Imports System.Data.OleDb

Partial Class management_BDP210A_FileUpload
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim sSavePath As String = Get_SysPara("FileUpload") & "BDP210A_tmp.xls"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Not IsPostBack) Then
        End If
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim PostedFile As HttpPostedFile = FileUpload1.PostedFile
        Dim sFileName As String = Trim(PostedFile.FileName)
        Dim sExt As String = IO.Path.GetExtension(sFileName).ToLower

        If Trim(PostedFile.FileName.ToString) <> "" Then
            If FileUpload1.HasFile = True Then
                '檢查副檔名
                If Chk_Ext(sExt) = False Then
                    Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('上傳檔案 " & sFileName & " 附檔名非允許格式,請重新選擇檔案');</script>"))
                    Exit Sub
                End If
                '上傳檔案
                Try
                    PostedFile.SaveAs(sSavePath)
                Catch ex As Exception
                    Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('上傳失敗，請重新上傳');</script>"))
                    Exit Sub
                End Try
                '匯入資料
                If Update_ExcelFile() Then
                    Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('匯入成功');</script>"))
                Else
                    Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('檔案格式有誤，請確認');</script>"))
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

        Response.Write("<script>opener.window.location.href = opener.window.location.href;window.close();</script>")
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

        MyConnection = New OleDbConnection( _
        "provider=Microsoft.ACE.OLEDB.12.0; " & _
        "data source=" & sSavePath & "; " & _
        "Extended Properties=Excel 12.0;")

        '讀取sheel
        Try
            MyConnection.Open()
            schemaTable = MyConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
            MyConnection.Close()
            v_sheet1 = schemaTable.Rows(0).Item("TABLE_NAME").ToString()
            schemaTable.Clear()

        Catch ex As Exception
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('此檔案非EXCEL檔');</script>"))
            Return False
        End Try
        '讀取資料至DS
        MyCommand = New System.Data.OleDb.OleDbDataAdapter("SEL" & "ECT * FROM [" & v_sheet1 & "]", MyConnection)
        MyCommand.Fill(dsTmp)
        MyConnection.Close()

        '新增修改資料
        Dim code_code As String = ""
        Dim code_name As String = ""
        Dim scr_no As String = ""
        Dim Field_code As String = ""
        Dim Field_name As String = ""
        Dim Code_type As String = ""
        Dim is_use As String = ""
        
        Try
            For Each row As DataRow In dsTmp.Tables(0).Rows
                Select Case row.Item("地區").ToString
                    Case "中正區"
                        code_code = "100"
                        code_name = "中正區"
                    Case "大同區"
                        code_code = "103"
                        code_name = "大同區"
                    Case "中山區"
                        code_code = "104"
                        code_name = "中山區"
                    Case "松山區"
                        code_code = "105"
                        code_name = "松山區"
                    Case "大安區"
                        code_code = "106"
                        code_name = "大安區"
                    Case "萬華區"
                        code_code = "108"
                        code_name = "萬華區"
                    Case "信義區"
                        code_code = "110"
                        code_name = "信義區"
                    Case "士林區"
                        code_code = "111"
                        code_name = "士林區"
                    Case "北投區"
                        code_code = "112"
                        code_name = "北投區"
                    Case "內湖區"
                        code_code = "114"
                        code_name = "內湖區"
                    Case "南港區"
                        code_code = "115"
                        code_name = "南港區"
                    Case "文山區"
                        code_code = "116"
                        code_name = "文山區"

                End Select
                scr_no = Get_AutoIntMax("BDP210", "scr_no", " WHERE code_code ='" & code_code & "' ") + 1
                Field_code = Right("0000" & scr_no, 4)
                Field_name = row.Item("路段").ToString
                Code_type = "B"

                '自動編號
                'dan_code = "DA" & Format(Now, "yyMMdd") & Right("0000" & Get_AutoIntMax("Danger_work", "RIGHT(dan_code, 10)", " WHERE dan_code LIKE 'DA" & Format(Now, "yyMMdd") & "%'") + 1, 4)

                If Chk_RelData("BDP210A", "", code_code, Field_name) = True Then
                    SaveEditData("BDP210A", "ADD", "", code_code, code_name, scr_no, Field_code, Field_name, Code_type)
                End If
            Next
        Catch ex As Exception
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('EXCEL檔案欄位資料有誤" & ex.Message & "');</script>"))
            Return False
        End Try

        Return True
    End Function
End Class
