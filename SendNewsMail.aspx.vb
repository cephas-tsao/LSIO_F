Imports System.Data

Partial Class SendNewsMail
    Inherits PageBase

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Response.CacheControl = "no-cache"
        Response.AddHeader("Pragma", "no-cache")
        Response.Expires = -1
        Response.Buffer = True

        Dim ws As New WebReference.LSIO_WebService
        Dim sPrgCode As String = "SendNewsMail"
        Dim bResult As String = ""
        '取得排程
        Dim dtSchedule As DataTable = Get_DataSet(sPrgCode, "Schedule").Tables(0)

        If dtSchedule.Rows.Count > 0 Then
            Response.Write("共有 " & dtSchedule.Rows.Count & " 筆送信排程(執行時間:" & Format(Now, "yyyy/MM/dd HH:mm") & ")<br>")
            '取得訂戶Mail List
            Dim dtMailList As DataTable = Get_DataSet(sPrgCode, "MailList").Tables(0)

            For i As Integer = 0 To dtSchedule.Rows.Count - 1
                bResult = "B"
                Dim sSchCode As String = dtSchedule.Rows(i).Item(0) & ""
                Dim sPaperCode As String = dtSchedule.Rows(i).Item(1) & ""
                Dim sTitle As String = dtSchedule.Rows(i).Item(2) & ""
                Dim sBody As String = dtSchedule.Rows(i).Item(3) & ""

                For j As Integer = 0 To dtMailList.Rows.Count - 1
                    Dim sMail As String = dtMailList.Rows(j).Item(0) & ""
                    Dim sName As String = dtMailList.Rows(j).Item(1) & ""

                    If Chk_Email(sMail) = False Then
                        bResult = "C"
                        '寫入Log檔
                        SaveEditData("LSI005A_detail", "ADD", "", sSchCode, sMail, sName)
                    Else
                        '發送Mail
                        Dim dtMail As New DataTable
                        Dim column1 As New DataColumn
                        Dim column2 As New DataColumn
                        column1.DataType = System.Type.GetType("System.String")
                        column2.DataType = System.Type.GetType("System.String")
                        dtMail.Columns.Add(column1)
                        dtMail.Columns.Add(column2)
                        Dim rowMail As DataRow = dtMail.NewRow
                        rowMail(0) = sMail
                        rowMail(1) = sName
                        dtMail.Rows.Add(rowMail)
                        '若發信失敗則離開程式
                        If SendNewsMail(sTitle, dtMail, sBody, sPrgCode) = False Then
                            Response.Write("發信失敗!!!(執行時間:" & Format(Now, "yyyy/MM/dd HH:mm") & ")<br>")
                            Exit Sub
                        End If
                    End If
                    'Response.Write(dtMailList.Rows(j).Item(0) & dtMailList.Rows(j).Item(1))
                Next
                Dim sDate = ""
                '若非第一次發信，則寫入重送日期
                If Chk_RelData("LSI005A", "", sSchCode) = True Then sDate = Format(Now, "yyyy/MM/dd")
                '更新發送標籤
                SaveEditData(sPrgCode, "Result", sSchCode, bResult, sDate)
                'Response.Write(dtSchedule.Rows(i).Item(0))
            Next
        Else
            Response.Write("查無送信排程(執行時間:" & Format(Now, "yyyy/MM/dd HH:mm") & ")<br>")
        End If
    End Sub
End Class
