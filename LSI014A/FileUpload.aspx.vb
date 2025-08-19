Imports System.Data
Imports System.Data.OleDb

Partial Class basic_LSI014A_FileUpload
    Inherits PageBase

    Dim ws As New WebReference.LSIO_WebService
    Dim sSavePath As String = Get_SysPara("FileUpload") & "LSI014A_tmp.xls"

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
        Dim dan_code As String = ""
        Dim com_name As String = ""
        Dim pet_name As String = ""
        Dim pet_date As String = ""
        Dim pet_time As String = ""
        Dim pet_tel1 As String = ""
        Dim att_name1 As String = ""
        Dim att_tel1 As String = ""
        Dim att_name2 As String = ""
        Dim att_tel2 As String = ""
        Dim att_add1 As String = ""
        Dim att_area As String = ""
        Dim att_road As String = ""
        Dim bulid_name As String = ""
        Dim bulid_code As String = ""
        Dim work_date_s As String = ""
        Dim work_date_e As String = ""
        Dim is_night As String = ""
        Dim work_floor_s As String = ""
        Dim work_floor_e As String = ""
        Dim work_type1 As String = ""
        Dim work_type_memo As String = ""
        Dim use_tool1 As String = ""
        Dim use_tool5_memo As String = ""
        Dim save_tool As String = ""
        Dim save_tool_memo As String = ""
        Dim dan_memo_memo As String = ""
        Dim dan_mail As String = ""
        Dim dan_mail2 As String = ""
        '計算結果
        Dim iTotalCount As Integer = 0
        Dim iFailCount As Integer = 0
        Dim sMessage As String = ""
        '檢查資料內容
        Dim bCheck As Boolean

        Try
            For Each row As DataRow In dsTmp.Tables(0).Rows
                iTotalCount += 1
                com_name = row.Item("事業單位名稱").ToString
                pet_name = row.Item("通報人").ToString
                pet_date = row.Item("通報日期").ToString
                pet_time = row.Item("通報時間").ToString
                pet_tel1 = row.Item("通報人聯絡電話").ToString
                att_name1 = row.Item("廠商名稱").ToString
                att_tel1 = row.Item("廠商電話").ToString
                att_name2 = row.Item("施工人員姓名").ToString
                att_tel2 = row.Item("施工人員電話").ToString
                att_add1 = row.Item("施工地址").ToString
                att_area = row.Item("施工區域").ToString
                att_road = row.Item("施工路段").ToString
                bulid_name = row.Item("大廈名稱").ToString
                bulid_code = row.Item("建號名稱").ToString
                work_date_s = row.Item("施工期間起").ToString
                work_date_e = row.Item("施工期間止").ToString
                is_night = row.Item("是否夜間施工").ToString
                work_floor_s = row.Item("作業樓層起").ToString
                work_floor_e = row.Item("作業樓層止").ToString
                work_type1 = row.Item("作業類別").ToString
                work_type_memo = row.Item("作業類別-其他說明").ToString
                use_tool1 = row.Item("使用機具").ToString
                use_tool5_memo = row.Item("使用機具-其他說明").ToString
                save_tool = row.Item("安全裝置").ToString
                save_tool_memo = row.Item("安全裝置-其他說明").ToString
                dan_memo_memo = row.Item("備註").ToString
                dan_mail = row.Item("通報回覆1").ToString
                dan_mail2 = row.Item("通報回覆2").ToString

                bCheck = True
                '檢查資料內容
                If Trim(com_name) = "" Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:事業單位名稱空白\n"
                End If
                If com_name.Length > 50 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:事業單位名稱超過50字\n"
                End If
                If pet_name.Length > 10 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:通報人超過10字\n"
                End If
                If pet_date.Length <> 10 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:通報日期格式錯誤(Ex:2012/01/01)\n"
                End If
                If pet_time.Length <> 5 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:通報時間格式錯誤(Ex:12:00)\n"
                End If
                If pet_tel1.Length > 16 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:通報人聯絡電話超過16字\n"
                End If
                If att_name1.Length > 50 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:廠商名稱超過50字\n"
                End If
                If att_tel1.Length > 16 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:廠商電話超過16字\n"
                End If
                If att_name2.Length > 50 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:施工人員姓名超過50字\n"
                End If
                If att_tel2.Length > 16 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:施工人員電話超過16字\n"
                End If
                If att_add1.Length > 100 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:施工地址超過100字\n"
                End If
                If att_area.Length > 10 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:施工區域超過10字\n"
                End If
                If att_road.Length > 10 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:施工路段超過10字\n"
                End If
                If bulid_name.Length > 20 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:大廈名稱超過20字\n"
                End If
                If bulid_code.Length > 20 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:建號名稱超過20字\n"
                End If
                If work_date_s.Length <> 10 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:施工期間起格式錯誤(Ex:2012/01/01)\n"
                End If
                If work_date_e.Length <> 10 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:施工期間止格式錯誤(Ex:2012/01/01)\n"
                End If
                If is_night <> "Y" And is_night <> "N" Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:是否夜間施工格式錯誤(Y/N)\n"
                End If
                If work_floor_s.Length > 10 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:作業樓層起超過10字\n"
                End If
                If work_floor_e.Length > 10 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:作業樓層止超過10字\n"
                End If
                If work_type1.Length > 20 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:作業類別超過20字\n"
                End If
                If work_type_memo.Length > 30 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:作業類別-其他說明超過30字\n"
                End If
                If use_tool1.Length > 20 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:使用機具超過20字\n"
                End If
                If use_tool5_memo.Length > 30 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:使用機具-其他說明超過30字\n"
                End If
                If save_tool.Length > 100 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:安全裝置超過100字\n"
                End If
                If save_tool_memo.Length > 30 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:安全裝置-其他說明超過30字\n"
                End If
                If dan_memo_memo.Length > 30 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:備註超過30字\n"
                End If
                If dan_mail.Length > 100 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:通報回覆1超過100字\n"
                End If
                If dan_mail2.Length > 100 Then
                    bCheck = False
                    sMessage &= "第 " & iTotalCount & " 筆:通報回覆2超過100字\n"
                End If

                '若資料有錯則跳下一筆
                If bCheck = False Then
                    iFailCount += 1
                    Continue For
                End If

                '自動編號
                dan_code = "DA" & Format(Now, "yyMMdd") & Right("0000" & Get_AutoIntMax("Danger_work", "RIGHT(dan_code, 10)", " WHERE dan_code LIKE 'DA" & Format(Now, "yyMMdd") & "%'") + 1, 4)

                If Chk_RelData("LSI014A", "", dan_code) = True Then
                    If SaveEditData("LSI014A", "ADD", dan_code, com_name, pet_name, pet_date, pet_time _
                        , pet_tel1, att_name1, att_tel1, att_name2, att_tel2, att_add1 _
                        , att_area, att_road, bulid_name, bulid_code, work_date_s _
                        , work_date_e, is_night, work_floor_s, work_floor_e, work_type1 _
                        , "", "", "", "" _
                        , "", "", "", "" _
                        , work_type_memo, use_tool1, "", "", "" _
                        , "", "", use_tool5_memo, save_tool, save_tool_memo _
                        , "", dan_memo_memo, dan_mail, dan_mail2) = False Then
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
            sMessage &= "欲查看匯入資料請重新查詢"
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('" & sMessage & "');</script>"))
        Catch ex As Exception
            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('EXCEL檔案欄位資料有誤');</script>"))
            Return False
        End Try

        Return True
    End Function
End Class
