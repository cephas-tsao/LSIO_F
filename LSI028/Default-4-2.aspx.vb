
Partial Class basic_LSI028_Default_4
    Inherits System.Web.UI.Page
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim Conn As SqlConnection = New SqlConnection
        Conn.ConnectionString = WebConfigurationManager.ConnectionStrings("AppSysConnectionString").ConnectionString
        Conn.open()

        Dim pic2, seq1, seq2, seq3, seq4, seq5, seq6, seq7, seq8, seqtt
        Dim p As String = Request("p")
        Dim snum As String = Request("snum")
        Dim c As String = Request("c")
        Dim IDA As String = Request("ID")

        Dim haveRec = False

        '-- p 就是「目前在第幾頁?」    
        Conn.open()
        Dim SqlStra1 As String

        SqlStra1 = "SELECT * FROM self_chk order by date desc"


        Dim da As SqlDataAdapter = New SqlDataAdapter(SqlStra1, Conn)
        Dim DS As DataSet = New DataSet
        da.Fill(DS)    '-- 把SQL指令執行完成的結果，填入DataSet裡面。

        Dim DT As DataTable = DS.Tables(0)
        '=============  ADO.NET / DataSet ==(End)======
        Dim hcnt As Integer = 0
        hcnt = DS
        Literal1.text = ""
        For int 
            'Response.Write("<h3>已送訂單筆數共計" & RecordCount & "筆</h3>")

            '--rowNo，目前畫面出現的這一頁，要撈出幾筆（列）資料
            Dim rowNo As Integer = 0
            Dim html As String = ""
    End Sub
End Class
