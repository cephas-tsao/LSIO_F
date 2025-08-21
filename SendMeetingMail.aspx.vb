Imports System.Data

Partial Class SendMeetingMail
    Inherits PageBase

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Response.CacheControl = "no-cache"
        Response.AddHeader("Pragma", "no-cache")
        Response.Expires = -1
        Response.Buffer = True

        Dim ws As New WebReference.LSIO_WebService
        Dim sPrgCode As String = "SendMeetingMail"
        '取得排程
        Dim dtSchedule As DataTable = Get_DataSet(sPrgCode, "Schedule").Tables(0)

        If dtSchedule.Rows.Count > 0 Then
            Response.Write("共有 " & dtSchedule.Rows.Count & " 筆送信排程(執行時間:" & Format(Now, "yyyy/MM/dd HH:mm") & ")<br>")

            For i As Integer = 0 To dtSchedule.Rows.Count - 1
                Dim sMrdCode As String = dtSchedule.Rows(i).Item(0) & ""
                Dim sRoomName As String = dtSchedule.Rows(i).Item(1) & ""
                Dim sMrdDate As String = dtSchedule.Rows(i).Item(2) & ""
                Dim sMrdTimeS As String = dtSchedule.Rows(i).Item(3) & ""
                Dim sMrdTimeE As String = dtSchedule.Rows(i).Item(4) & ""
                Dim sMrdName As String = dtSchedule.Rows(i).Item(5) & ""
                Dim sMrdHost As String = dtSchedule.Rows(i).Item(6) & ""
                Dim sMrdMemo As String = dtSchedule.Rows(i).Item(7) & ""

                '取得訂戶Mail List
                Dim dtMailList As DataTable = Get_DataSet(sPrgCode, "MailList", sMrdCode).Tables(0)

                For j As Integer = 0 To dtMailList.Rows.Count - 1
                    Dim sMail As String = dtMailList.Rows(j).Item(0) & ""
                    Dim sName As String = dtMailList.Rows(j).Item(1) & ""

                    If Chk_Email(sMail) = False Then
                        ''寫入Log檔
                        'SaveEditData("LSI005A_detail", "ADD", "", sSchCode, sMail, sName)
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

                        '信計內容
                        Dim sBody As String = "" & _
                            "" & _
                            "<table cellpadding=5 cellspacing=0 bordercolor=#7EBAD1 border=1 bgcolor=#ffffff style='border-collapse:collapse;'" & _
                            "width='720' summary='排版表格' align=center class='c12'>" & _
                            "<tr bgcolor='#7EBAD1' align='center'><td colspan='2'><p><strong>開會通知</strong></p></td></tr>" & _
                            "<tr>" & _
                            "  <td>會議名稱：</td>" & _
                            "  <td>" & sMrdName & "</td>" & _
                            "</tr>" & _
                            "<tr>" & _
                            "  <td>主持人：</td>" & _
                            "  <td>" & sMrdHost & "</td>" & _
                            "</tr>" & _
                            "<tr>" & _
                            "  <td>會議室：</td>" & _
                            "  <td>" & sRoomName & "</td>" & _
                            "</tr>" & _
                            "<tr>" & _
                            "  <td>開會日期：</td>" & _
                            "  <td>" & sMrdDate & "</td>" & _
                            "</tr>" & _
                            "<tr>" & _
                            "  <td>開會時間：</td>" & _
                            "  <td>" & sMrdTimeS & " ～ " & sMrdTimeE & "</td>" & _
                            "</tr>" & _
                            "<tr>" & _
                            "  <td>備註：</td>" & _
                            "  <td>" & Replace(sMrdMemo, Chr(10), "<br />") & "</td>" & _
                            "</tr>" & _
                            "</table>"

                        '若發信失敗則離開程式
                        If SendMail(sMrdName & " - 開會通知", dtMail, sBody, sPrgCode) = False Then
                            Response.Write("發信失敗!!!(執行時間:" & Format(Now, "yyyy/MM/dd HH:mm") & ")<br>")
                            Exit Sub
                        End If
                    End If
                    'Response.Write(dtMailList.Rows(j).Item(0) & dtMailList.Rows(j).Item(1))
                Next
                'Dim sDate = ""
                ''若非第一次發信，則寫入重送日期
                'If Chk_RelData("LSI005A", "", sSchCode) = True Then sDate = Format(Now, "yyyy/MM/dd")
                ''更新發送標籤
                'SaveEditData(sPrgCode, "Result", sSchCode, bResult, sDate)
                'Response.Write(dtSchedule.Rows(i).Item(0))
            Next
        Else
            Response.Write("查無送信排程(執行時間:" & Format(Now, "yyyy/MM/dd HH:mm") & ")<br>")
        End If
    End Sub
End Class
