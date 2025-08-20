Imports System.Data
Imports System.Data.OleDb

Partial Class basic_Downtime_FileUpload
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    'Dim sSavePath As String = Get_SysPara("FileUpload") & "Downtime.xls"
    Dim sSavePath As String = Server.MapPath("/upload/") & "Downtime.xls"

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
        Dim Down_institutions As String = ""
        Dim Down_title As String = ""
        Dim Down_address As String = ""
        Dim Down_X As String = ""
        Dim Down_Y As String = ""
        Dim Down_stopDate As String = ""
        Dim Down_returnDate As String = ""
        Dim Down_range As String = ""
        Dim Down_reason As String = ""
        Dim Down_remarks As String = ""
        'Dim Down_sDate As String = ""
        'Dim Down_eDate As String = ""
        'Dim Down_pDate As String = ""
        'Dim Down_cDate As String = ""
        'Dim Down_uDate As String = ""
        Dim usr_code As String = Session("usr_code")
        Dim ins_date As String = Format(Now, "yyyy/MM/dd")
        '計算結果
        Dim iTotalCount As Integer = 0
        Dim iFailCount As Integer = 0
        Dim sMessage As String = ""

        Try
            '清空舊有資料
            SaveEditData("Downtime", "DEL")

            For Each row As DataRow In dsTmp.Tables(0).Rows
                iTotalCount += 1
                Down_institutions = row.Item("*事業單位名稱").ToString
                Down_title = row.Item("*工程名稱").ToString
                Down_address = row.Item("*停工地點").ToString
                Down_X = row.Item("緯度").ToString
                Down_Y = row.Item("經度").ToString
                Down_stopDate = FormatDateTime(row.Item("*停工日期").ToString, DateFormat.ShortDate)
                'Down_returnDate = IIf(row.Item("復工日期").ToString = "", "", FormatDateTime(row.Item("復工日期").ToString, DateFormat.ShortDate))

                If row.Item("復工日期").ToString = "" Then
                    Down_returnDate = ""
                Else
                    Down_returnDate = FormatDateTime(row.Item("復工日期").ToString, DateFormat.ShortDate)
                End If

                Down_range = row.Item("*停工範圍").ToString
                Down_reason = row.Item("*停工原因").ToString
                Down_remarks = row.Item("備註").ToString
                'Down_sDate = row.Item("").ToString
                'Down_eDate = row.Item("").ToString
                'Down_pDate = row.Item("").ToString
                'Down_cDate = FormatDateTime(row.Item("建立日期").ToString, DateFormat.ShortDate)
                'Down_uDate = FormatDateTime(row.Item("修改日期").ToString, DateFormat.ShortDate)

                If SaveEditData("Downtime", "ADD", Down_institutions, _
                        Down_title, Down_address, _
                        Down_stopDate, Down_returnDate, Down_range, _
                        Down_reason, Down_remarks, _
                        Session("usr_code"), _
                        "", Down_X, Down_Y) = False Then
                    'insert into不需資料編號 => 傳空值
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
            'Response.Write("err=" & ex.Message)
            'Response.End()
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('EXCEL檔案欄位資料有誤');</script>"))
            Return False
        End Try

        Return True
    End Function
End Class
