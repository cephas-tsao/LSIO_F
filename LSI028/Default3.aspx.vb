Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Web.Configuration
Imports System.Web.Helpers
Partial Class basic_LSI028_Default3
    Inherits System.Web.UI.Page
    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
        'checks database connection string and handles error if there is one
        'anticfrs.InnerHtml = AntiForgery.GetHtml().ToString()
    End Sub
    Protected Sub drpSites_SelectedIndexChanged(sender As Object, e As EventArgs)
      
    End Sub

    Protected Sub search_Click(sender As Object, e As System.EventArgs)
        Dim Conn As SqlConnection = New SqlConnection
        Conn.ConnectionString = WebConfigurationManager.ConnectionStrings("AppSysConnectionString").ConnectionString
        Dim SqlStra1 As String
        Dim SqlStra2 As String
        Dim SqlStra3 As String
        Dim SqlStra4 As String
        Dim SqlStra5 As String
        lbyear.InnerHtml = HttpUtility.HtmlEncode(DDLYear.SelectedValue & "年")
        lbyear1.InnerHtml = HttpUtility.HtmlEncode(DDLYear.SelectedValue & "年")
        lbyear2.InnerHtml = HttpUtility.HtmlEncode(DDLYear.SelectedValue & "年")
        lbyear3.InnerHtml = HttpUtility.HtmlEncode(DDLYear.SelectedValue & "年")
        Dim year As String = DDLYear.SelectedValue
        If drpSites.SelectedValue = "一般科" Then
            SqlStra1 = "SELECT  *  FROM LSI026 where pc_type = N'一般科' and pc_season=1 and pc_year=" & year
            SqlStra2 = "SELECT  *  FROM LSI026 where pc_type = N'一般科' and pc_season=2 and pc_year=" & year
            SqlStra3 = "SELECT  *  FROM LSI026 where pc_type = N'一般科' and pc_season=3 and pc_year=" & year
            SqlStra4 = "SELECT  *  FROM LSI026 where pc_type = N'一般科' and pc_season=4 and pc_year=" & year
            SqlStra5 = "SELECT  *  FROM Self_com where unit = N'一般科' "

        ElseIf drpSites.SelectedValue = "勞條科" Then
            SqlStra1 = "SELECT  *  FROM LSI026 where pc_type = N'勞條科' and pc_season=1 and pc_year=" & year
            SqlStra2 = "SELECT  *  FROM LSI026 where pc_type = N'勞條科' and pc_season=2 and pc_year=" & year
            SqlStra3 = "SELECT  *  FROM LSI026 where pc_type = N'勞條科' and pc_season=3 and pc_year=" & year
            SqlStra4 = "SELECT  *  FROM LSI026 where pc_type = N'勞條科' and pc_season=4 and pc_year=" & year
            SqlStra5 = "SELECT  *  FROM Self_com where unit = N'勞條科' "
        End If

        Dim da1 As SqlDataAdapter = New SqlDataAdapter(SqlStra1, Conn)
        Dim da2 As SqlDataAdapter = New SqlDataAdapter(SqlStra2, Conn)
        Dim da3 As SqlDataAdapter = New SqlDataAdapter(SqlStra3, Conn)
        Dim da4 As SqlDataAdapter = New SqlDataAdapter(SqlStra4, Conn)
        Dim da5 As SqlDataAdapter = New SqlDataAdapter(SqlStra5, Conn)

        Dim DS1 As DataSet = New DataSet
        Dim DS2 As DataSet = New DataSet
        Dim DS3 As DataSet = New DataSet
        Dim DS4 As DataSet = New DataSet
        Dim DS5 As DataSet = New DataSet

        da1.Fill(DS1)    '-- 把SQL指令執行完成的結果，填入DataSet裡面。
        da2.Fill(DS2)    '-- 把SQL指令執行完成的結果，填入DataSet裡面。
        da3.Fill(DS3)    '-- 把SQL指令執行完成的結果，填入DataSet裡面。
        da4.Fill(DS4)    '-- 把SQL指令執行完成的結果，填入DataSet裡面。
        da5.Fill(DS5)    '-- 把SQL指令執行完成的結果，填入DataSet裡面。
        Dim DT1 As DataTable = DS1.Tables(0)
        Dim DT2 As DataTable = DS2.Tables(0)
        Dim DT3 As DataTable = DS3.Tables(0)
        Dim DT4 As DataTable = DS4.Tables(0)
        Dim DT5 As DataTable = DS5.Tables(0)

        '=============  ADO.NET / DataSet ==(End)======
        Dim num1 As Integer = 0
        Dim num2 As Integer = 0
        Dim num3 As Integer = 0
        Dim num4 As Integer = 0
        Dim num5 As Integer = 0
        Dim num6 As Integer = 0
        num1 = DT1.Rows.Count()
        num2 = DT2.Rows.Count()
        num3 = DT3.Rows.Count()
        num4 = DT4.Rows.Count()
        num5 = DT5.Rows.Count()


        num_1.Value = num1 & ""
        num_2.Value = num2 & ""
        num_3.Value = num3 & ""
        num_4.Value = num4 & ""

        total1.Value = num5
        total2.Value = num5
        total3.Value = num5
        total4.Value = num5

        rate1.Value = Int((num1 / num5) * 100) & "%"
        rate2.Value = Int((num2 / num5) * 100) & "%"
        rate3.Value = Int((num3 / num5) * 100) & "%"
        rate4.Value = Int((num4 / num5) * 100) & "%"
    End Sub
End Class
