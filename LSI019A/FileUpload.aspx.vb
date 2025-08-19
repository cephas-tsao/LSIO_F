Imports System.Data
Imports System.Data.OleDb

Partial Class basic_LSI019A_FileUpload
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim sSavePath As String = Get_SysPara("FileUpload") & "LSI019A_tmp.xls"

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
                    'Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('匯入成功');</script>"))
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

        'Response.Write("<script>opener.window.location.href = opener.window.location.href;</script>") 'window.close();
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
        Dim chk_date1 As String = ""
        Dim chk_date2 As String = ""
        Dim com_name As String = ""
        Dim eng_name As String = ""
        Dim work_staus As String = ""
        Dim usr_code As String = Session("usr_code")
        Dim ins_date As String = Format(Now, "yyyy/MM/dd")
        '計算結果
        Dim iTotalCount As Integer = 0
        Dim iFailCount As Integer = 0
        Dim sMessage As String = ""

        Try
            '清空就有資料
            SaveEditData("LSI019A", "DEL")

            For Each row As DataRow In dsTmp.Tables(0).Rows
                iTotalCount += 1
                chk_date1 = row.Item("第1次檢查日期").ToString
                chk_date2 = row.Item("第2次檢查日期").ToString
                com_name = row.Item("事業單位名稱").ToString
                eng_name = row.Item("工程名稱").ToString
                work_staus = row.Item("工地狀態").ToString

                'If Chk_RelData("LSI019A", "stop_code", stop_code) = True And Chk_RelData("LSI009A", "stop_number", stop_number) = True Then
                If SaveEditData("LSI019A", "ADD", "", chk_date1, chk_date2, com_name, eng_name _
                            , work_staus, usr_code, ins_date) = False Then
                    iFailCount += 1
                    sMessage &= "第 " & iTotalCount & " 筆:資料有誤\n"
                End If
                'Else
                'iFailCount += 1
                'sMessage &= "第 " & iTotalCount & " 筆:資料重覆\n"
                'End If
            Next

            sMessage &= "總共匯入: " & iTotalCount & " 件\n"
            sMessage &= "匯入成功: " & iTotalCount - iFailCount & " 件\n"
            sMessage &= "匯入失敗 " & iFailCount & " 件\n"
            sMessage &= "欲查看匯入資料請重新查詢"
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('" & sMessage & "');</script>"))
        Catch ex As Exception
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('EXCEL檔案欄位資料有誤');</script>"))
            Return False
        End Try

        Return True
    End Function
End Class
