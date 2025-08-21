Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic

Partial Class Json
    Inherits PageBase

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '連線資料庫
        Dim Connection_Db As SqlConnection
        Connection_Db = New SqlConnection("Data Source=localhost;Initial Catalog=LSIODB;User ID=sa;Password=1208jsh;Pooling=True")
        Connection_Db.Open()

        Dim sSql As String
        sSql = "select news_type,news_title,news_code from e_news "

        '抓資料進DT
        Using (Connection_Db)
            '建立adapter
            Dim Fun_Adpt As SqlDataAdapter
            '建立DataTable
            Dim Fun_Dt1 As DataTable
            Fun_Dt1 = New DataTable

            Fun_Adpt = New SqlDataAdapter(sSql, Connection_Db)
            Fun_Adpt.Fill(Fun_Dt1)

            '轉換為Json資料格式
            Dim ilop As Integer
            Dim result As String
            result = ""

            result = CreateJsonParameters(Fun_Dt1, "news")
            'Response.Write(result + "<br>")
            'result = ""

            'result = result + "{'news':["

            'For ilop = 0 To Fun_Dt1.Rows.Count - 1
            '    result = result + "{'news_code':'" + Fun_Dt1.Rows(ilop).Item("news_code").ToString() + "',"
            '    result = result + "'news_title':'" + Fun_Dt1.Rows(ilop).Item("news_title").ToString() + "'},"
            '    'result = result + "{'news_type':'" + Fun_Dt1.Rows(ilop).Item("news_type").ToString() + "'},"
            'Next

            'result = Strings.Left(result, result.Length - 1)
            'result = result + "]}"
            Response.Write(result)
        End Using




    End Sub

    ''' <summary>
    ''' 傳入一DataTable轉換為Json格式字串
    ''' </summary>
    ''' <param name="dt">待轉換的Dt</param>
    ''' <param name="dt_name">資料表名稱 取用時使用</param>
    ''' <returns>Json字串</returns>
    ''' <remarks></remarks>
    Public Function CreateJsonParameters(ByVal dt As DataTable, ByVal dt_name As String) As String

        Dim JsonString As StringBuilder = New StringBuilder()
        If dt IsNot Nothing And dt.Rows.Count > 0 Then

            JsonString.Append("{ ")
            JsonString.Append("'" + dt_name + "'" + ":[ ")
            For i As Integer = 0 To dt.Rows.Count - 1

                JsonString.Append("{ ")
                For j As Integer = 0 To dt.Columns.Count - 1

                    If j < dt.Columns.Count - 1 Then

                        JsonString.Append("'" & dt.Columns(j).ColumnName.ToString() + "'" + ":" + "'" + dt.Rows(i)(j).ToString() + "'" + ",")

                    ElseIf j = dt.Columns.Count - 1 Then

                        JsonString.Append("'" & dt.Columns(j).ColumnName.ToString() + "'" + ":" + "'" + dt.Rows(i)(j).ToString() + "'")
                    End If
                Next

                '字串最後一個字元右上Json結尾符號
                If i = dt.Rows.Count - 1 Then

                    JsonString.Append("} ")

                Else

                    JsonString.Append("}, ")
                End If
            Next
            JsonString.Append("]}")
            Return JsonString.ToString()

        Else
            Return Nothing
        End If
    End Function





End Class
