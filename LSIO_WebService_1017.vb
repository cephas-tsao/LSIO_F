Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

' 若要允許使用 ASP.NET AJAX 從指令碼呼叫此 Web 服務，請取消註解下一行。
' <System.Web.Script.Services.ScriptService()> _
<WebService(Namespace:="http://lsio.netdoing.com.tw:8080/LSIO_WebService/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class LSIO_WebService
    Inherits System.Web.Services.WebService

    ''' <summary>
    ''' 輸入程式代碼及參數，傳回DateSet
    ''' </summary>
    ''' <param name="sPrgcode">程式代碼</param>
    ''' <param name="sType">種類-select/delete/內頁編輯頁detail</param>
    ''' <param name="sPara1">參數1</param>
    ''' <param name="sPara2">參數2</param>
    ''' <param name="sPara3">參數3</param>
    ''' <param name="sPara4">參數4</param>
    ''' <param name="sPara5">參數5</param>
    ''' <param name="sPara6">參數6</param>
    ''' <returns>DateSet</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function Get_DataSet(ByVal sPrgcode As String, ByVal sType As String, ByVal sPara1 As String, ByVal sPara2 As String, ByVal sPara3 As String, ByVal sPara4 As String, ByVal sPara5 As String, ByVal sPara6 As String) As DataSet
        Dim sSql As String = ""
        Dim tmpDS As New DataSet
        Dim tmpDT As New DataTable
        Dim tmpAdpt As SqlDataAdapter
        Dim SqlCmd As New SqlCommand
        Dim SqlCon As SqlConnection = New SqlConnection(Get_ConnStr(""))

        If sType = "Detail" Then        '類型為Detail則共用SQL字串
            sSql = Get_SqlStr(sPrgcode, sType, sPara1, sPara2, sPara3, sPara4, sPara5, sPara6)
            SqlCmd = New SqlCommand(sSql, SqlCon)

        ElseIf sType = "SqlStr" Then    '類型為SqlStr則為SQL字串,直接使用
            sSql = sPara1
            SqlCmd = New SqlCommand(sSql, SqlCon)

        Else

            Select Case sPrgcode
                Case "Default"
                    Select Case sType
                        Case "A"
                            sSql = "SELECT * FROM BDP150 WHERE ip_type='A'"
                            SqlCmd = New SqlCommand(sSql, SqlCon)

                        Case "B"
                            sSql = "SELECT * FROM BDP150 WHERE ip_type='B'"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                    End Select

                Case "Login"
                    Select Case sType
                        Case "Check"
                            'sSql = "SELECT BDP080.* FROM BDP080 WHERE BDP080.usr_code='" & RepSql(sPara1) & "' AND is_use='Y'"
                            sSql = "SELECT BDP080.* FROM BDP080 WHERE BDP080.usr_code=@Para1 AND is_use='Y'"
                            'sSql = "SELECT BDP080.* FROM BDP080 WHERE is_use='Y'"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                            'SqlCmd.Parameters.AddWithValue("@Para2", RepSql(sPara2))

                        Case "ch_date"
                            'sSql = "SELECT * FROM BDP200 WHERE ((prg_code = 'BDP080A') AND (usr_code = '" & RepSql(sPara1) & "')) OR ((prg_code = 'KPB006A') AND (usr_code = '" & RepSql(sPara1) & "')) ORDER BY bdp200 DESC"
                            sSql = "SELECT * FROM BDP200 WHERE ((prg_code = 'BDP080A') AND (usr_code=@Para1)) OR ((prg_code = 'KPB006A') AND (usr_code =@Para1)) ORDER BY bdp200 DESC"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                        Case "IDno"
                            'sSql = "SELECT BDP080.* FROM BDP080 WHERE BDP080.usr_idno='" & RepSql(sPara1) & "' AND is_use='Y'"
                            sSql = "SELECT BDP080.* FROM BDP080"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            'SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                        Case "Mail"
                            'sSql = "SELECT BDP080.* FROM BDP080 WHERE BDP080.usr_mail='" & RepSql(sPara1) & "'"
                            sSql = "SELECT BDP080.* FROM BDP080 WHERE BDP080.usr_mail=@Para1"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                    End Select

                Case "WebUserMenu"
                    Select Case sType
                        Case "Lv0"
                            sSql = "SELECT prg_code,prg_name FROM BDP030 WHERE is_use='Y' AND menu_level='0' ORDER BY scr_no"
                            SqlCmd = New SqlCommand(sSql, SqlCon)

                        Case "Lv1"
                            'sSql = "SELECT * FROM BDP030 WHERE is_use='Y' AND menu_level='1' AND up_level='" & RepSql(sPara1) & "' ORDER BY scr_no"
                            sSql = "SELECT * FROM BDP030 WHERE is_use='Y' AND menu_level='1' AND up_level=@Para1 ORDER BY scr_no"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                        Case "Lv2"
                            'sSql = "SELECT * FROM BDP030 WHERE is_use='Y' AND menu_level='2' AND up_level='" & RepSql(sPara1) & "' ORDER BY scr_no"
                            sSql = "SELECT * FROM BDP030 WHERE is_use='Y' AND menu_level='2' AND up_level=@Para1 ORDER BY scr_no"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                    End Select

                Case "addForm"
                    Select Case sType
                        Case "GetField"
                            'sSql = "SELECT BDP160.*,field_code,field_name,field_type,scr_no,sel_field" & _
                            '       "  FROM BDP160 " & _
                            '       "  LEFT JOIN BDP161 ON BDP161.add_code=BDP160.add_code" & _
                            '       " WHERE BDP160.add_code='" & RepSql(sPara1) & "'" & _
                            '       " ORDER BY BDP161.scr_no"
                            sSql = "SELECT BDP160.*,field_code,field_name,field_type,scr_no,sel_field" & _
                                   "  FROM BDP160 " & _
                                   "  LEFT JOIN BDP161 ON BDP161.add_code=BDP160.add_code" & _
                                   " WHERE BDP160.add_code=@Para1" & _
                                   " ORDER BY BDP161.scr_no"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                        Case "GetGridView"
                            'sSql = "SELECT sel_str,order_str FROM BDP160 WHERE add_code='" & RepSql(sPara1) & "'"
                            sSql = "SELECT sel_str,order_str FROM BDP160 WHERE add_code=@Para1"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                    End Select

                Case "addList"
                    'sSql = "SELECT sel_str,order_str,list_value,list_name FROM BDP170 WHERE add_code='" & RepSql(sPara1) & "'"
                    sSql = "SELECT sel_str,order_str,list_value,list_name FROM BDP170 WHERE add_code=@Para1"
                    SqlCmd = New SqlCommand(sSql, SqlCon)
                    SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                Case "AutoComplete"
                    'sSql = "SELECT sel_str FROM BDP172 WHERE auto_code='" & RepSql(sPara1) & "' AND sel_code='" & RepSql(sPara2) & "'"
                    sSql = "SELECT sel_str FROM BDP172 WHERE auto_code=@Para1 AND sel_code=@Para2"
                    SqlCmd = New SqlCommand(sSql, SqlCon)
                    SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                    SqlCmd.Parameters.AddWithValue("@Para2", RepSql(sPara2))

                Case "RecordDetail"
                    Select Case sType
                        Case "GetData"
                            'sSql = "SELECT * FROM BDP200 WHERE bdp200='" & RepSql(sPara1) & "'"
                            sSql = "SELECT * FROM BDP200 WHERE bdp200=@Para1"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                        Case "GetName"
                            'sSql = "SELECT field_name FROM BDP220 WHERE prg_code='" & RepSql(sPara1) & "' AND field_code='" & RepSql(sPara2) & "'"
                            sSql = "SELECT field_name FROM BDP220 WHERE prg_code=@Para1 AND field_code=@Para2"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                            SqlCmd.Parameters.AddWithValue("@Para2", RepSql(sPara2))
                    End Select

                Case "Set_DDL_Option", "Set_CKL_Option"
                    'sSql = "SELECT * FROM BDP210 WHERE code_code='" & RepSql(sPara1) & "' ORDER BY scr_no"
                    sSql = "SELECT * FROM BDP210 WHERE code_code=@Para1 ORDER BY scr_no"
                    SqlCmd = New SqlCommand(sSql, SqlCon)
                    SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                Case "BDP002A"
                    Select Case sType
                        Case "GetMenu"
                            sSql = "SELECT * FROM BDP030 WHERE is_use='Y' AND is_group='N' ORDER BY prg_code,prg_type"
                            SqlCmd = New SqlCommand(sSql, SqlCon)

                        Case "GetLimit"
                            'sSql = "SELECT limit_str,prg_code FROM BDP091 WHERE grp_code='" & RepSql(sPara1) & "' and prg_code='" & RepSql(sPara2) & "'"
                            sSql = "SELECT limit_str,prg_code FROM BDP091 WHERE grp_code=@Para1 and prg_code=@Para2"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                            SqlCmd.Parameters.AddWithValue("@Para2", RepSql(sPara2))
                    End Select

                Case "BDP003A"
                    Select Case sType
                        Case "GetMenu"
                            sSql = "SELECT * FROM BDP030 WHERE is_use='Y' AND is_group='N' ORDER BY prg_code,prg_type"
                            SqlCmd = New SqlCommand(sSql, SqlCon)

                        Case "GetLimit"
                            'sSql = "SELECT limit_str,prg_code FROM BDP090 WHERE usr_code='" & RepSql(sPara1) & "' and prg_code='" & RepSql(sPara2) & "'"
                            sSql = "SELECT limit_str,prg_code FROM BDP090 WHERE usr_code=@Para1 and prg_code=@Para2"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                            SqlCmd.Parameters.AddWithValue("@Para2", RepSql(sPara2))

                        Case "GetType"
                            'sSql = "SELECT Limit_type FROM BDP080 WHERE usr_code='" & RepSql(sPara1) & "'"
                            sSql = "SELECT Limit_type FROM BDP080 WHERE usr_code=@Para1"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                    End Select

                Case "BDP003B"
                    Select Case sType
                        Case "GetMenu"
                            sSql = "SELECT * FROM BDP030 WHERE is_use='Y' AND is_group='N' ORDER BY prg_code,prg_type"
                            SqlCmd = New SqlCommand(sSql, SqlCon)

                        Case "GetLimit"
                            'sSql = "SELECT limit_str,prg_code FROM BDP092 WHERE dep_code='" & RepSql(sPara1) & "' and prg_code='" & RepSql(sPara2) & "'"
                            sSql = "SELECT limit_str,prg_code FROM BDP092 WHERE dep_code=@Para1 and prg_code=@Para2"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                            SqlCmd.Parameters.AddWithValue("@Para2", RepSql(sPara2))

                        Case "GetType"
                            'sSql = "SELECT Limit_type FROM BDP080 WHERE dep_code='" & RepSql(sPara1) & "'"
                            sSql = "SELECT Limit_type FROM BDP080 WHERE dep_code=@Para1"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                    End Select

                Case "BDP006A"
                    'sSql = "SELECT * FROM BDP080 WHERE usr_code ='" & RepSql(sPara1) & "'"
                    sSql = "SELECT * FROM BDP080 WHERE usr_code =@Para1"
                    SqlCmd = New SqlCommand(sSql, SqlCon)
                    SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                Case "LSI002A"
                    Select Case sType
                        Case "SRC"
                            Dim Ddate As Integer = Val(Get_SysPara("disable_date", "150"))
                            'sSql = "SELECT cmp_code FROM mail_complain WHERE 1=1  AND (is_agree='N' AND ins_date<'" & Format(Now.AddDays(0 - Ddate), "yyyy/MM/dd") & "') "
                            sSql = "SELECT cmp_code FROM mail_complain WHERE is_agree='N' AND ins_date<@Ddate"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Ddate", Format(Now.AddDays(0 - Ddate), "yyyy/MM/dd"))

                        Case "DEL"
                            'sSql = "DELETE FROM mail_complain WHERE cmp_code='" & RepSql(sPara1) & "'"
                            sSql = "DELETE FROM mail_complain WHERE cmp_code=@Para1"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                    End Select

                Case "LSI022A"
                    Select Case sType
                        Case "SRC"
                            Dim Ddate As Integer = Val(Get_SysPara("disable_date", "150"))
                            'sSql = "SELECT cmp_code FROM mail_complain WHERE 1=1  AND (is_agree='N' AND ins_date<'" & Format(Now.AddDays(0 - Ddate), "yyyy/MM/dd") & "') "
                            sSql = "SELECT cmp_code FROM mail WHERE ins_date<@Ddate"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Ddate", Format(Now.AddDays(0 - Ddate), "yyyy/MM/dd"))

                        Case "DEL"
                            'sSql = "DELETE FROM mail_complain WHERE cmp_code='" & RepSql(sPara1) & "'"
                            sSql = "DELETE FROM mail WHERE cmp_code=@Para1"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                    End Select

                Case "LSI004A"
                    'sSql = "SELECT DISTINCT sch_code FROM E_PAPER_SCH WHERE paper_code='" & RepSql(sPara1) & "'"
                    sSql = "SELECT DISTINCT sch_code FROM E_PAPER_SCH WHERE paper_code=@Para1"
                    SqlCmd = New SqlCommand(sSql, SqlCon)
                    SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                Case "LSI004A_detail"
                    Select Case sType
                        Case "Paper"
                            'sSql = "SELECT * FROM e_paper WHERE paper_code='" & RepSql(sPara1) & "'"
                            sSql = "SELECT * FROM e_paper WHERE paper_code=@Para1"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                        Case "PaperDetail01"
                            'sSql = "SELECT * FROM e_paper_d LEFT JOIN e_news ON e_paper_d.news_code=e_news.news_code" & _
                            '       " WHERE paper_code='" & RepSql(sPara1) & "' AND e_paper_d.news_type='01' ORDER BY e_paper_d.scr_no"
                            sSql = "SELECT * FROM e_paper_d LEFT JOIN e_news ON e_paper_d.news_code=e_news.news_code" & _
                                   " WHERE paper_code=@Para1 AND e_paper_d.news_type='01' ORDER BY e_paper_d.scr_no"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                        Case "PaperDetail02"
                            'sSql = "SELECT * FROM e_paper_d LEFT JOIN e_news ON e_paper_d.news_code=e_news.news_code" & _
                            '       " WHERE paper_code='" & RepSql(sPara1) & "' AND e_paper_d.news_type='02' ORDER BY e_paper_d.scr_no"
                            sSql = "SELECT * FROM e_paper_d LEFT JOIN e_news ON e_paper_d.news_code=e_news.news_code" & _
                                   " WHERE paper_code=@Para1 AND e_paper_d.news_type='02' ORDER BY e_paper_d.scr_no"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                        Case "PaperDetail03"
                            'sSql = "SELECT * FROM e_paper_d LEFT JOIN e_news ON e_paper_d.news_code=e_news.news_code" & _
                            '       " WHERE paper_code='" & RepSql(sPara1) & "' AND e_paper_d.news_type='03' ORDER BY e_paper_d.scr_no"
                            sSql = "SELECT * FROM e_paper_d LEFT JOIN e_news ON e_paper_d.news_code=e_news.news_code" & _
                                   " WHERE paper_code=@Para1 AND e_paper_d.news_type='03' ORDER BY e_paper_d.scr_no"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                        Case "PaperDetail04"
                            'sSql = "SELECT * FROM e_paper_d LEFT JOIN e_news ON e_paper_d.news_code=e_news.news_code" & _
                            '       " WHERE paper_code='" & RepSql(sPara1) & "' AND e_paper_d.news_type='04' ORDER BY e_paper_d.scr_no"
                            sSql = "SELECT * FROM e_paper_d LEFT JOIN e_news ON e_paper_d.news_code=e_news.news_code" & _
                                   " WHERE paper_code=@Para1 AND e_paper_d.news_type='04' ORDER BY e_paper_d.scr_no"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                        Case "PaperDetail05"
                            'sSql = "SELECT * FROM e_paper_d LEFT JOIN e_news ON e_paper_d.news_code=e_news.news_code" & _
                            '       " WHERE paper_code='" & RepSql(sPara1) & "' AND e_paper_d.news_type='05' ORDER BY e_paper_d.scr_no"
                            sSql = "SELECT * FROM e_paper_d LEFT JOIN e_news ON e_paper_d.news_code=e_news.news_code" & _
                                   " WHERE paper_code=@Para1 AND e_paper_d.news_type='05' ORDER BY e_paper_d.scr_no"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                        Case "PaperDetail06"
                            'sSql = "SELECT * FROM e_paper_d LEFT JOIN e_news ON e_paper_d.news_code=e_news.news_code" & _
                            '       " WHERE paper_code='" & RepSql(sPara1) & "' AND e_paper_d.news_type='06' ORDER BY e_paper_d.scr_no"
                            sSql = "SELECT * FROM e_paper_d LEFT JOIN e_news ON e_paper_d.news_code=e_news.news_code" & _
                                   " WHERE paper_code=@Para1 AND e_paper_d.news_type='06' ORDER BY e_paper_d.scr_no"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                        Case "InTime"
                            'sSql = "SELECT * FROM e_paper_d LEFT JOIN e_news ON e_paper_d.news_code=e_news.news_code" & _
                            '       " WHERE paper_code='" & RepSql(sPara1) & "' ORDER BY e_paper_d.scr_no"
                            sSql = "SELECT * FROM e_paper_d LEFT JOIN e_news ON e_paper_d.news_code=e_news.news_code" & _
                                   " WHERE paper_code=@Para1 ORDER BY e_paper_d.scr_no"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                        Case "SummaryL"
                            sSql = "SELECT DISTINCT field_name,news_type FROM E_PAPER_D " & _
                                   " LEFT JOIN BDP210 ON BDP210.field_code=E_PAPER_D.news_type AND BDP210.code_code='news_type'" & _
                                   " WHERE paper_code=@Para1 AND news_type='01' " & _
                                   "UNION SELECT news_title,news_code FROM E_NEWS" & _
                                   " WHERE news_code IN (SELECT news_code FROM E_PAPER_D WHERE paper_code=@Para1 AND news_type='01')" & _
                                   " ORDER BY news_type"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                        Case "SummaryR"
                            sSql = "SELECT DISTINCT field_name,news_type FROM E_PAPER_D" & _
                                   " LEFT JOIN BDP210 ON BDP210.field_code=E_PAPER_D.news_type AND BDP210.code_code='news_type'" & _
                                   " WHERE paper_code=@Para1 AND news_type<>'01' ORDER BY news_type"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                    End Select

                Case "LSI009A"
                    Select Case sType
                        Case "StopNumber"
                            'sSql = "SELECT * FROM Work_stop WHERE stop_number='" & RepSql(sPara1) & "'"
                            sSql = "SELECT * FROM Work_stop WHERE stop_number=@Para1"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                        Case Else
                            'sSql = "SELECT DISTINCT start_code FROM WORK_START" & _
                            '       " LEFT JOIN WORK_STOP ON WORK_START.stop_number=WORK_STOP.stop_number" & _
                            '       " WHERE stop_code='" & RepSql(sPara1) & "'"
                            sSql = "SELECT DISTINCT start_code FROM WORK_START" & _
                                   " LEFT JOIN WORK_STOP ON WORK_START.stop_number=WORK_STOP.stop_number" & _
                                   " WHERE stop_code=@Para1"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                    End Select

                Case "LSI016A"
                    'sSql = "SELECT DISTINCT paper_code FROM E_PAPER_D WHERE news_code='" & RepSql(sPara1) & "'"
                    sSql = "SELECT DISTINCT paper_code FROM E_PAPER_D WHERE news_code=@Para1"
                    SqlCmd = New SqlCommand(sSql, SqlCon)
                    SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                Case "Board"
                    'sSql = "SELECT * FROM Board WHERE bull_code='" & RepSql(sPara1) & "' AND is_hid='Y' ORDER BY Ins_date,Ins_time"
                    sSql = "SELECT * FROM Board WHERE bull_code=@Para1 AND is_hid='Y' ORDER BY Ins_date,Ins_time"
                    SqlCmd = New SqlCommand(sSql, SqlCon)
                    SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                Case "guidance"
                    sSql = "SELECT * FROM guidance"
                    SqlCmd = New SqlCommand(sSql, SqlCon)

                Case "guidance1"
                    sSql = "SELECT * FROM guidance1"
                    SqlCmd = New SqlCommand(sSql, SqlCon)
                Case "YComplain"
                    sSql = "SELECT * FROM YComplain"
                    SqlCmd = New SqlCommand(sSql, SqlCon)
                Case "NComplain"
                    sSql = "SELECT * FROM NComplain"
                    SqlCmd = New SqlCommand(sSql, SqlCon)
                Case "Complain"
                    sSql = "SELECT * FROM Complain"
                    SqlCmd = New SqlCommand(sSql, SqlCon)


                Case "SendNewsMail"
                    Select Case sType
                        Case "Schedule"
                            'sSql = "SELECT sch_code,e_paper.paper_code,e_paper.paper_name,e_paper.paper_html FROM e_paper_sch" & _
                            '       " LEFT JOIN e_paper ON e_paper.paper_code=e_paper_sch.paper_code" & _
                            '       " WHERE send_date='" & Format(Now, "yyyy/MM/dd") & "' AND send_time<='" & Format(Now, "HH:mm") & "' AND send_state='A'"
                            sSql = "SELECT sch_code,e_paper.paper_code,e_paper.paper_name,e_paper.paper_html FROM e_paper_sch" & _
                                   " LEFT JOIN e_paper ON e_paper.paper_code=e_paper_sch.paper_code" & _
                                   " WHERE send_date=@Date AND send_time<=@Time AND send_state='A'"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Date", Format(Now, "yyyy/MM/dd"))
                            SqlCmd.Parameters.AddWithValue("@Time", Format(Now, "HH:mm"))

                        Case "MailList"
                            sSql = "SELECT per_mail,com_name FROM e_paper_order WHERE Is_order='Y'"
                            SqlCmd = New SqlCommand(sSql, SqlCon)

                        Case "MailList_p1000"
                            sSql = "SELECT TOP 1000 per_mail,com_name FROM E_PAPER_ORDER WHERE Is_order='Y'" & _
                                   " AND per_mail NOT IN (SELECT per_mail FROM E_PAPER_TMP WHERE sch_code=@Para1 AND ins_date=@Date)" & _
                                   " ORDER BY ord_code"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                            SqlCmd.Parameters.AddWithValue("@Date", Format(Now, "yyyy/MM/dd"))

                        Case "SendLog"
                            'sSql = "SELECT mail_list FROM e_paper_log WHERE sch_code='" & RepSql(sPara1) & "'"
                            sSql = "SELECT mail_list FROM e_paper_log WHERE sch_code=@Para1"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                        Case "News"
                            'sSql = "SELECT paper_name,paper_html FROM e_paper WHERE paper_code='" & RepSql(sPara1) & "'"
                            sSql = "SELECT paper_name,paper_html FROM e_paper WHERE paper_code=@Para1"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                    End Select

                Case "SendMeetingMail"
                    Select Case sType
                        Case "Schedule"
                            'sSql = "SELECT mrd_code,room_name,mrd_date,mrd_time_s,mrd_time_e,mrd_name,mrd_host,mrd_memo FROM MEETING_ORDER" & _
                            '       " LEFT JOIN MEETING_ROOM ON MEETING_ORDER.room_code=MEETING_ROOM.room_code" & _
                            '       " WHERE (mrd_date='" & Format(Now.AddDays(7), "yyyy/MM/dd") & "' AND send_list LIKE '%A%')" & _
                            '       " OR  (mrd_date='" & Format(Now.AddDays(3), "yyyy/MM/dd") & "' AND send_list LIKE '%B%')" & _
                            '       " OR  (mrd_date='" & Format(Now.AddDays(1), "yyyy/MM/dd") & "' AND send_list LIKE '%C%') " & _
                            '       " OR  (mrd_date='" & Format(Now, "yyyy/MM/dd") & "' AND send_list LIKE '%D%')"
                            sSql = "SELECT mrd_code,room_name,mrd_date,mrd_time_s,mrd_time_e,mrd_name,mrd_host,mrd_memo FROM MEETING_ORDER" & _
                                   " LEFT JOIN MEETING_ROOM ON MEETING_ORDER.room_code=MEETING_ROOM.room_code" & _
                                   " WHERE (mrd_date=@Date7 AND send_list LIKE '%A%')" & _
                                   " OR  (mrd_date=@Date3 AND send_list LIKE '%B%')" & _
                                   " OR  (mrd_date=@Date1 AND send_list LIKE '%C%') " & _
                                   " OR  (mrd_date=@Date0 AND send_list LIKE '%D%')"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Date7", Format(Now.AddDays(7), "yyyy/MM/dd"))
                            SqlCmd.Parameters.AddWithValue("@Date3", Format(Now.AddDays(3), "yyyy/MM/dd"))
                            SqlCmd.Parameters.AddWithValue("@Date1", Format(Now.AddDays(1), "yyyy/MM/dd"))
                            SqlCmd.Parameters.AddWithValue("@Date0", Format(Now, "yyyy/MM/dd"))

                        Case "MailList"
                            'sSql = "SELECT usr_mail,usr_name FROM MEETING_USERS WHERE mrd_code='" & RepSql(sPara1) & "'"
                            sSql = "SELECT usr_mail,usr_name FROM MEETING_USERS WHERE mrd_code=@Para1"
                            SqlCmd = New SqlCommand(sSql, SqlCon)
                            SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                    End Select

                Case "ForgetPassword"
                    'sSql = "SELECT usr_name,usr_mail,usr_pass FROM BDP080 WHERE usr_code='" & RepSql(sPara1) & "' AND is_use='Y'"
                    sSql = "SELECT usr_name,usr_mail,usr_pass FROM BDP080 WHERE usr_code=@Para1 AND is_use='Y'"
                    SqlCmd = New SqlCommand(sSql, SqlCon)
                    SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                Case "WorkStart"
                    'sSql = "SELECT * FROM Work_stop WHERE stop_number='" & RepSql(sPara1) & "'"
                    sSql = "SELECT * FROM Work_stop WHERE stop_number=@Para1"
                    SqlCmd = New SqlCommand(sSql, SqlCon)
                    SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                Case "DanReport"
                    'sSql = "SELECT eng_name,eng_add as addr,pet_date,a.Field_name,b.Field_name FROM DANGER_CHECK " & _
                    '       " LEFT JOIN BDP210 a ON DANGER_CHECK.danc_state=a.field_code AND a.code_code='danc_state'" & _
                    '       " LEFT JOIN BDP210 b ON DANGER_CHECK.result=b.field_code AND b.code_code='result'" & _
                    '       " WHERE con_email='" & RepSql(sPara1) & "' AND con_pw='" & RepSql(sPara2) & "'"
                    sSql = "SELECT eng_name,eng_add as addr,pet_date,a.Field_name,b.Field_name FROM DANGER_CHECK " & _
                           " LEFT JOIN BDP210 a ON DANGER_CHECK.danc_state=a.field_code AND a.code_code='danc_state'" & _
                           " LEFT JOIN BDP210 b ON DANGER_CHECK.result=b.field_code AND b.code_code='result'" & _
                           " WHERE con_email=@Para1 AND con_pw=@Para2"
                    SqlCmd = New SqlCommand(sSql, SqlCon)
                    SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                    SqlCmd.Parameters.AddWithValue("@Para2", RepSql(sPara2))

                Case "ComReport"

                    sSql = "SELECT * FROM MAIL_COMPLAIN WHERE cmp_email=@Para1 and cmp_pass=@Para2"
                    SqlCmd = New SqlCommand(sSql, SqlCon)
                    SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                    SqlCmd.Parameters.AddWithValue("@Para2", RepSql(sPara2))

                Case "UsrList"
                    'sSql = "SELECT * FROM BDP080 LEFT JOIN BDP230 ON BDP080.dep_code=BDP230.dep_code" & _
                    '       " WHERE BDP080.is_use='Y' AND usr_code='" & RepSql(sPara1) & "'"
                    sSql = "SELECT * FROM BDP080 LEFT JOIN BDP230 ON BDP080.dep_code=BDP230.dep_code" & _
                           " WHERE BDP080.is_use='Y' AND usr_code=@Para1"
                    SqlCmd = New SqlCommand(sSql, SqlCon)
                    SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                Case "app_area"
                    If RepSql(sPara1) = "" Then
                        sSql = "SELECT Field_code,Field_name FROM BDP210 WHERE Code_code='ZipCode' ORDER BY Scr_no"
                        SqlCmd = New SqlCommand(sSql, SqlCon)
                    Else
                        'sSql = "SELECT Field_code,Field_name FROM BDP210 WHERE Code_code='" & RepSql(sPara1) & "' ORDER BY Scr_no"
                        sSql = "SELECT Field_code,Field_name FROM BDP210 WHERE Code_code=@Para1 ORDER BY Scr_no"
                        SqlCmd = New SqlCommand(sSql, SqlCon)
                        SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                    End If

                Case "app_DangerData"
                    'sSql = "SELECT com_name,att_add1 FROM DANGER_WORK WHERE att_area='" & RepSql(sPara1) & "' AND att_road='" & RepSql(sPara2) & "'"
                    sSql = "SELECT com_name,att_add1 FROM DANGER_WORK WHERE att_area=@Para1 AND att_road=@Para2" & _
                           " AND work_date_s<=@Para3 AND work_date_e>=@Para3"
                    SqlCmd = New SqlCommand(sSql, SqlCon)
                    SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                    SqlCmd.Parameters.AddWithValue("@Para2", RepSql(sPara2))
                    SqlCmd.Parameters.AddWithValue("@Para3", RepSql(Format(Now, "yyyy/MM/dd")))

                Case "app_DangerCheck"
                    sSql = "SELECT a.*,b.field_name AS WorkType FROM DANGER_WORK a" & _
                           " LEFT JOIN BDP210 b ON a.work_type1=b.field_code AND b.code_code='work_type1'" & _
                           " WHERE att_area=@Para1 AND att_road=@Para2 AND att_add1=@Para3"
                    SqlCmd = New SqlCommand(sSql, SqlCon)
                    SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                    SqlCmd.Parameters.AddWithValue("@Para2", RepSql(sPara2))
                    SqlCmd.Parameters.AddWithValue("@Para3", RepSql(sPara3))

                Case "app_UseTool"
                    sSql = "SELECT field_name FROM BDP210 WHERE code_code='use_tool1' AND field_code=@Para1"
                    SqlCmd = New SqlCommand(sSql, SqlCon)
                    SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                Case "app_Labor"
                    'sSql = "SELECT chk_date1,chk_date2,com_name,eng_name,work_staus FROM LABOR_INSPECT" & _
                    '       " WHERE chk_date2='' AND (chk_date1='' OR chk_date1<'" & RepSql(sPara1) & "')"
                    sSql = "SELECT chk_date1,chk_date2,com_name,eng_name,work_staus FROM LABOR_INSPECT" & _
                           " WHERE chk_date2='' AND (chk_date1='' OR chk_date1<@Para1)"
                    SqlCmd = New SqlCommand(sSql, SqlCon)
                    SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
            End Select
        End If

        Try
            tmpAdpt = New SqlDataAdapter(SqlCmd)
            tmpAdpt.Fill(tmpDT)
            tmpDS.Tables.Add(tmpDT)
        Catch ex As Exception
            tmpDS.Tables.Add(New DataTable)
        End Try
        Return tmpDS
    End Function

    ''' <summary>
    ''' 傳入程式代碼等參數取得SQL字串
    ''' </summary>
    ''' <param name="sPrgcode">程式代碼</param>
    ''' <param name="sType">種類-select/delete/內頁編輯頁detail</param>
    ''' <param name="sPara1">參數1</param>
    ''' <param name="sPara2">參數2</param>
    ''' <param name="sPara3">參數3</param>
    ''' <param name="sPara4">參數4</param>
    ''' <param name="sPara5">參數5</param>
    ''' <param name="sPara6">參數6</param>
    ''' <returns>SQL字串</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function Get_SqlStr(ByVal sPrgcode As String, ByVal sType As String, ByVal sPara1 As String, ByVal sPara2 As String, ByVal sPara3 As String, ByVal sPara4 As String, ByVal sPara5 As String, ByVal sPara6 As String) As String
        Dim sSql As String = ""
        '類型為ddlField則傳回通用SQL字串
        If sType = "ddlField" Then Return "SELECT * FROM BDP140 WHERE prg_code='" & RepSql(sPrgcode) & "' ORDER BY scr_no"

        Select Case sPrgcode
            Case "addForm"
                sSql = "SELECT * FROM BDP161 WHERE add_code='" & RepSql(sPara1) & "' ORDER BY scr_no"

            Case "addList"
                sSql = "SELECT * FROM BDP171 WHERE add_code='" & RepSql(sPara1) & "' ORDER BY scr_no"

            Case "BDP030option1"
                sSql = "SELECT prg_code AS field_code, prg_name AS field_name  FROM BDP030 WHERE is_group = 'Y' AND is_use = 'Y' ORDER BY scr_no "

            Case "BDP000A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT * FROM BDP000 " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM BDP000 WHERE bdp000=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM BDP000 WHERE bdp000='" & RepSql(sPara1) & "'"
                End Select

            Case "BDP002A"
                sSql = "SELECT grp_code AS field_code,grp_code + '.' + grp_name AS field_name FROM BDP070 ORDER BY grp_code"

            Case "BDP003A"
                sSql = "SELECT usr_code AS field_code,usr_code + '.' + usr_name AS field_name FROM BDP080 WHERE is_use='Y' ORDER BY grp_code,usr_code"

            Case "BDP003B"
                sSql = "SELECT dep_code AS field_code,dep_code + '.' + dep_name AS field_name FROM BDP230 WHERE is_use='Y' ORDER BY dep_code"


            Case "BDP007A"
                sSql = "SELECT * FROM BDP000 WHERE par_name ='css_sys'"

            Case "BDP030A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT * FROM BDP030 " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "Delete from BDP030 where prg_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM BDP030 WHERE prg_code='" & RepSql(sPara1) & "'"
                End Select

            Case "BDP070A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT * FROM BDP070 " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM BDP070 WHERE grp_code=@datakeys;DELETE FROM BDP091 WHERE grp_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM BDP070 WHERE grp_code='" & RepSql(sPara1) & "'"
                End Select

            Case "BDP080A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT * FROM BDP080 " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM BDP080 WHERE usr_code=@datakeys;DELETE FROM BDP090 WHERE usr_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM BDP080 WHERE usr_code='" & RepSql(sPara1) & "'"

                    Case "GrpCode"
                        sSql = "SELECT grp_name AS field_name, grp_code AS field_code FROM BDP070 WHERE is_use='Y'"

                    Case "DepCode"
                        sSql = "SELECT dep_code as field_code, dep_name as field_name FROM BDP230 WHERE is_use='Y'"

                End Select

            Case "BDP140A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT * FROM BDP140 " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM BDP140 WHERE bdp140=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM BDP140 WHERE bdp140='" & RepSql(sPara1) & "'"
                End Select

            Case "BDP150A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT *, CAST(ip_s1 AS varchar(10)) + '.' + CAST(ip_s2 AS varchar(10)) + '.' + CAST(ip_s3 AS varchar(10)) + '.' + CAST(ip_s4 AS varchar(10)) as ips, " & _
                               "  CAST(ip_e1 AS varchar(10)) + '.' + CAST(ip_e2 AS varchar(10)) + '.' + CAST(ip_e3 AS varchar(10)) + '.' + CAST(ip_e4 AS varchar(10)) as ipe, " & _
                               "  case ip_type when 'A' then '白名單' when 'B' then '黑名單' end as ip_types" & _
                               "  FROM BDP150 " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM BDP150 WHERE serial_no=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM BDP150 WHERE serial_no='" & RepSql(sPara1) & "'"
                End Select

            Case "BDP190A"
                Select Case sType
                    Case "SELECT"
                        'sSql = "SELECT * FROM Board WHERE bull_type='0'" & Replace(sPara1, "WHERE", "AND", , 1) & sPara2
                        sSql = "SELECT *,CASE WHEN LEN(theme) > 25 THEN SUBSTRING(theme, 1 , 25) + '...' ELSE theme END AS theme1," & _
                              "(SELECT COUNT(*) FROM Board WHERE bull_code=a.bull_code AND bull_type<>'0') AS cnt," & _
                              "(SELECT TOP 1 Ins_date FROM Board WHERE bull_code=a.bull_code AND bull_type<>'0' ORDER BY Ins_date DESC,Ins_time DESC) as last " & _
                              "FROM Board AS a WHERE bull_type='0' " & Replace(sPara1, "WHERE", "AND", , 1) & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM Board WHERE bull_code=@datakeys "

                    Case "Detail"
                        sSql = "SELECT * FROM Board WHERE bull_type='0' AND bull_code='" & RepSql(sPara1) & "'"
                End Select

            Case "BDP190A_detail"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT * FROM Board WHERE bull_type <> '0' AND bull_code ='" & RepSql(sPara1) & "'" & sPara2 & " ORDER BY ins_date,ins_time"

                    Case "DELETE"
                        sSql = "DELETE FROM Board WHERE bdp190=@datakeys;"

                    Case "Detail"
                        sSql = "SELECT * FROM Board WHERE bdp190='" & RepSql(sPara1) & "'"
                End Select

            Case "BDP200A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT BDP200.*,BDP030.prg_name FROM BDP200" & _
                               " LEFT JOIN BDP030 ON BDP200.prg_code=BDP030.prg_code " & sPara1 & " ORDER BY usr_date DESC,usr_time DESC"

                    Case "DELETE"
                        sSql = "DELETE FROM BDP200 WHERE bdp200=@datakeys"
                End Select

            Case "BDP210A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT * FROM BDP210 " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM BDP210 WHERE bdp210=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM BDP210 WHERE bdp210='" & RepSql(sPara1) & "'"

                End Select

            Case "BDP220A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT * FROM BDP220 " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM BDP220 WHERE bdp220=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM BDP220 WHERE bdp220='" & RepSql(sPara1) & "'"

                End Select

            Case "BDP230A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT * FROM BDP230 " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM BDP230 WHERE dep_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM BDP230 WHERE dep_code='" & RepSql(sPara1) & "'"

                End Select

            Case "LSI001A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT mail_asklaw.*,BDP210.field_name as ask_types FROM mail_asklaw" & _
                               " LEFT JOIN BDP210 ON mail_asklaw.ask_type=BDP210.field_code AND BDP210.code_code='AskType' " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM mail_asklaw WHERE ask_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM mail_asklaw WHERE ask_code='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI002A"
                Select Case sType
                    Case "SELECT"
                        Dim Ddate As Integer = Val(Get_SysPara("disable_date", "150"))
                        sPara1 &= " AND (is_agree<>'N' OR ins_date>='" & Format(Now.AddDays(0 - Ddate), "yyyy/MM/dd") & "') "

                        sSql = "SELECT mail_complain.*,BDP210.field_name as cmp_types " & _
                               " ,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(com_Contract" & _
                               " ,'A','定期契約'),'B','終止契約爭議'),'C','預告期間'),'D','資遣費'),'E','退休金'),'F','服務證明書'),'G','資遣通報') as com_Contract_c" & _
                               " ,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(com_paymon" & _
                               " ,'A','未達基本工資'),'B','積欠工資'),'C','加班費'),'D','扣薪(違約或賠償)'),'E','年終獎金') as com_paymon_c" & _
                               " ,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(com_relax" & _
                               " ,'A','超時'),'B','休息'),'C','例假'),'D','國定節日'),'E','特別休假'),'F','假日出勤未加給工資'),'G','婚喪病等假') as com_relax_c" & _
                               " ,REPLACE(REPLACE(REPLACE(com_child" & _
                               " ,'A','童工'),'B','女工夜間工作'),'C','產假') as com_child_c" & _
                               " ,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(com_Welfar" & _
                               " ,'A','職災補償'),'B','技術生'),'C','工作規則'),'D','職工福利'),'E','安全衛生') as com_Welfar_c" & _
                               " FROM mail_complain" & _
                               " LEFT JOIN BDP210 ON mail_complain.cmp_type=BDP210.field_code AND BDP210.code_code='CmpType' " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM mail_complain WHERE cmp_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM mail_complain WHERE cmp_code='" & RepSql(sPara1) & "'"
                End Select

            Case "yumi2" '20160909 yumi 新增
                Select Case sType
                    Case "SELECT"
                        'sSql = " SELECT yumi2.* FROM yumi2 " & _
                        '       sPara1 & sPara2
                        sSql = " SELECT y.book_code,y.book_title,b.Field_name " & _
                        "[book_author],y.book_translation,y.book_publisher,y.book_area,y.book_road,y.book_address," & _
                        "y.book_phone,y.book_print,y.book_premiums,y.pet_date FROM yumi2 y " & _
                        " left join BDP210 b " & _
                        " on b.Field_code=y.book_author  and (b.Code_code='book_autho')" & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM yumi2 WHERE book_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM yumi2 WHERE book_code='" & RepSql(sPara1) & "'"
                End Select

            Case "yumi2-2" '20160908 yumi 新增
                Select Case sType
                    Case "SELECT"
                        Dim Ddate As Integer = Val(Get_SysPara("disable_date", "150"))
                        sPara1 &= " AND (is_agree<>'N' OR ins_date>='" & Format(Now.AddDays(0 - Ddate), "yyyy/MM/dd") & "') "

                        sSql = " SELECT yumi2_2.*,BDP210.field_name as per_types FROM yumi2_2 " & _
                               " LEFT JOIN BDP210 ON yumi2_2.per_type=BDP210.field_code AND BDP210.code_code='per_type' " & _
                               sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM yumi2_2 WHERE ord_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM yumi2_2 WHERE ord_code='" & RepSql(sPara1) & "'"
                End Select

            Case "Downtime" '20161004 yumi 新增
                Select Case sType
                    Case "SELECT"
                        sSql = " SELECT Down_code,Down_institutions,Down_title,Down_address,Down_X,Down_Y, " & _
                            " convert(varchar, Down_stopDate, 111) Down_stopDate, convert(varchar, Down_returnDate, 111) Down_returnDate, " & _
                            "Down_range, Down_reason, Down_remarks, " & _
                            "convert(varchar, Down_sDate, 111) Down_sDate, " & _
                            "convert(varchar, Down_eDate, 111) Down_eDate, " & _
                            "convert(varchar, Down_pDate, 111) Down_pDate, " & _
                            "convert(varchar, Down_cDate, 111) Down_cDate, " & _
                            "convert(varchar, Down_uDate, 111) Down_uDate, " & _
                            "Down_creatorID,Down_modifyID,Down_creatorIP,Down_modifyIP FROM Downtime " & _
                               sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM Downtime WHERE Down_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT Down_code,Down_institutions,Down_title,Down_address,Down_X,Down_Y, " & _
                            " convert(varchar, Down_stopDate, 111) Down_stopDate, convert(varchar, Down_returnDate, 111) Down_returnDate, " & _
                            "Down_range, Down_reason, Down_remarks, " & _
                            "convert(varchar, Down_sDate, 111) Down_sDate, " & _
                            "convert(varchar, Down_eDate, 111) Down_eDate, " & _
                            "convert(varchar, Down_pDate, 111) Down_pDate, " & _
                            "convert(varchar, Down_cDate, 111) Down_cDate, " & _
                            "convert(varchar, Down_uDate, 111) Down_uDate, " & _
                            "Down_creatorID,Down_modifyID,Down_creatorIP,Down_modifyIP " & _
                            " FROM Downtime WHERE Down_code='" & RepSql(sPara1) & "'"
                End Select

            Case "Banner" '20161129 yumi 新增
                Select Case sType
                    Case "SELECT"
                        sSql = " SELECT bn_code, title, imgName " & _
                            "FROM Banner " & _
                               sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM Banner WHERE bn_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT bn_code, title, imgName " & _
                            " FROM Banner WHERE bn_code='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI022A"
                Select Case sType
                    Case "SELECT"
                        Dim Ddate As Integer = Val(Get_SysPara("disable_date", "150"))
                        'sPara1 &= " AND (ins_date>='" & Format(Now.AddDays(0 - Ddate), "yyyy/MM/dd") & "') "

                        sSql = "SELECT * FROM mail " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM mail WHERE cmp_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM mail WHERE cmp_code='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI003A"
                Select Case sType
                    Case "SELECT"
                        sSql = " SELECT e_paper_order.*,BDP210.field_name as per_types FROM e_paper_order " & _
                               " LEFT JOIN BDP210 ON e_paper_order.per_type=BDP210.field_code AND BDP210.code_code='per_type' " & _
                               sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM e_paper_order WHERE ord_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM e_paper_order WHERE ord_code='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI004A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT e_paper.*,BDP080.usr_name FROM e_paper LEFT JOIN BDP080 ON e_paper.usr_code=BDP080.usr_code " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM e_paper_d WHERE paper_code=@datakeys;DELETE FROM e_paper WHERE paper_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM e_paper WHERE paper_code='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI004A_detail"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT e_paper_d.*,e_news.news_title,BDP210.field_name as news_types FROM e_paper_d " & _
                               " LEFT JOIN e_news ON e_paper_d.news_code=e_news.news_code " & _
                               " LEFT JOIN BDP210 ON BDP210.field_code=e_paper_d.news_type and code_code='news_type' " & _
                               " WHERE paper_code ='" & RepSql(sPara1) & "'" & sPara2 & " ORDER BY news_type,scr_no"

                    Case "DELETE"
                        sSql = "DELETE FROM e_paper_d WHERE serial_no=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM e_paper_d WHERE serial_no='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI005A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT e_paper_sch.*,BDP210.field_name as send_states,e_paper.paper_name as paper_names FROM e_paper_sch " & _
                               " LEFT JOIN BDP210 ON e_paper_sch.send_state=BDP210.field_code AND BDP210.code_code='sendstate' " & _
                               " LEFT JOIN E_PAPER ON e_paper_sch.paper_code=e_paper.paper_code " & _
                                sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM e_paper_sch WHERE sch_code=@datakeys;DELETE FROM e_paper_log WHERE sch_code=@datakeys;" & _
                               "DELETE FROM E_PAPER_TMP WHERE sch_code=@datakeys"


                    Case "Detail"
                        sSql = "SELECT * FROM e_paper_sch WHERE sch_code='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI005A_detail"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT e_paper_log.* FROM e_paper_log " & _
                               " WHERE sch_code ='" & RepSql(sPara1) & "'" & sPara2 & " ORDER BY serial_no"

                    Case "DELETE"
                        sSql = "DELETE FROM e_paper_log WHERE serial_no=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM e_paper_log WHERE serial_no='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI007A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT Meeting_room.*,BDP210.field_name as Is_uses FROM Meeting_room" & _
                               " LEFT JOIN BDP210 ON Meeting_room.Is_use=BDP210.field_code AND BDP210.code_code='IsUse' " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM Meeting_room WHERE room_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM Meeting_room WHERE room_code='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI008A"
                Select Case sType
                    Case "SELECT"
                        sSql = " SELECT mrd_code,mrd_time_s,mrd_time_e,mrd_name,mrd_host,Meeting_order.mrd_date,BDP080.usr_name,Meeting_room.room_name,Day FROM Meeting_order" & _
                               " LEFT JOIN BDP080 ON BDP080.usr_code=Meeting_order.usr_code" & _
                               " LEFT JOIN Meeting_room ON Meeting_room.room_code=Meeting_order.room_code" & _
                               " LEFT JOIN Meeting_Date ON Meeting_Date.mrd_date=Meeting_order.mrd_date" & _
                               " WHERE Meeting_order.mrd_date LIKE '" & RepSql(sPara1) & "%' AND Meeting_order.room_code LIKE '" & RepSql(sPara2) & "%'" & _
                               " UNION" & _
                               " SELECT '','新增','','','',mrd_date,'','',Day FROM Meeting_Date WHERE mrd_date LIKE '" & RepSql(sPara1) & "%'" & _
                               " ORDER BY Meeting_order.mrd_date,mrd_time_s"


                    Case "DELETE"
                        sSql = "DELETE FROM Meeting_order WHERE mrd_code=@datakeys;DELETE FROM Meeting_users WHERE mrd_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM Meeting_order WHERE mrd_code='" & RepSql(sPara1) & "'"

                    Case "RoomCode"
                        sSql = "SELECT room_name AS field_name, room_code AS field_code FROM Meeting_room WHERE is_use='Y'"

                    Case "UsrCode"
                        sSql = "SELECT usr_name AS field_name, usr_code AS field_code FROM BDP080 WHERE is_use='Y'"
                End Select

            Case "LSI008A_detail"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT * FROM Meeting_users WHERE mrd_code='" & RepSql(sPara1) & "'"

                    Case "DELETE"
                        sSql = "DELETE FROM Meeting_users WHERE serial_no=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM Meeting_users WHERE serial_no='" & RepSql(sPara1) & "'"

                End Select

            Case "LSI009A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT Work_stop.*,BDP080.usr_name FROM Work_stop LEFT JOIN BDP080 ON Work_stop.usr_code=BDP080.usr_code " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM Work_stop WHERE stop_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM Work_stop WHERE stop_code='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI010A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT Work_start.* FROM Work_start " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM Work_start WHERE start_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM Work_start WHERE start_code='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI011A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT Self_check.* FROM Self_check " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM Self_check WHERE self_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM Self_check WHERE self_code='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI012A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT Small_space.*,C.subname1 as subname1,BDP210.field_name as eng_areas  FROM Small_space" & _
                               " LEFT JOIN BDP210 ON Small_space.eng_area=BDP210.field_code AND BDP210.code_code='ZipCode'   LEFT JOIN V_IND_NAME C on Small_space.sub_category=C.code" & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM Small_space WHERE sma_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM Small_space WHERE sma_code='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI013A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT Fix_work.*,BDP210.field_name as att_areas FROM Fix_work" & _
                               " LEFT JOIN BDP210 ON Fix_work.att_area=BDP210.field_code AND BDP210.code_code='ZipCode' " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM Fix_work WHERE fix_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM Fix_work WHERE fix_code='" & RepSql(sPara1) & "'"
                End Select
            Case "LSI020A" 'may2013/10/16
                Select Case sType
                    Case "SELECT"
                        'sSql = "SELECT Fix_work.*,BDP210.field_name as att_areas FROM Fix_work" & _
                        '      " LEFT JOIN BDP210 ON Fix_work.att_area=BDP210.field_code AND BDP210.code_code='ZipCode' " & sPara1 & sPara2

                        sSql = "SELECT SCHOOL_WORK.*,BDP210.field_name as com_names,BDP210_1.field_name as att_areas" & _
                               " ,REPLACE(REPLACE(REPLACE(SCHOOL_WORK.work_floor_type" & _
                               " ,'A','屋頂'),'B','屋突'),'C','外牆') AS work_floor_type_c" & _
                                " ,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(SCHOOL_WORK.fix_type" & _
                               " ,'A','新建工程'),'B','修繕工程'),'C','機電設備維謢'),'D','結構補強'),'E','拆除工程'),'F','室內裝修'),'G','外牆修繕'),'H','屋頂作業'),'I','2公尺以上高處作業'),'J','使用合梯或施工架作業'),'K','其他') AS fix_type_c" & _
                                " FROM  BDP210 RIGHT OUTER JOIN  SCHOOL_WORK ON BDP210.Field_code = SCHOOL_WORK.com_name LEFT OUTER JOIN  BDP210 AS BDP210_1 ON SCHOOL_WORK.att_area = BDP210_1.Field_code " & sPara1 & sPara2
                    Case "DELETE"
                        sSql = "DELETE FROM SCHOOL_WORK WHERE fix_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM SCHOOL_WORK WHERE fix_code='" & RepSql(sPara1) & "'"
                End Select
            Case "LSI014A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT Danger_Work.*,a.field_name AS area_name,b.field_name AS road_name,c.field_name AS work_type" & _
                            " ,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(Danger_Work.use_tool1" & _
                            " ,'A','吊籠'),'B','高空工作車'),'C','高空繩索作業'),'D','移動式起重機')" & _
                            " ,'E','移動式起重機+搭乘設備(俗稱吊籃)'),'F','工作平台工法(從事升降梯安裝)')" & _
                            " ,'G','本體平台工法(從事升降梯安裝)') AS use_tool_c" & _
                            " ,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(Danger_Work.save_tool" & _
                            " ,'A','無'),'B','安全帽'),'C','腰帶式安全帶'),'E','背負式安全帶')" & _
                            " ,'F','背負式安全帶+雙鉤'),'G','雙繩作業+防墜器')" & _
                            " ,'H','雙繩作業+具有繩索訓練證照'),'I','防墜器') AS save_tool_c" & _
                            " FROM Danger_Work" & _
                            " LEFT JOIN BDP210 a ON Danger_Work.att_area=a.field_code AND a.code_code='ZipCode'" & _
                            " LEFT JOIN BDP210 b ON Danger_Work.att_road=b.field_code AND b.code_code=a.field_code" & _
                            " LEFT JOIN BDP210 c ON Danger_Work.work_type1=c.field_code AND c.code_code='work_type1'" & sPara1 & sPara2
                        'sSql = "SELECT Danger_Work.*,a.field_name AS area_name,b.field_name AS road_name,c.field_name AS work_type FROM Danger_Work" & _
                        '" LEFT JOIN BDP210 a ON Danger_Work.att_area=a.field_code AND a.code_code='ZipCode' " & _
                        '" LEFT JOIN BDP210 b ON Danger_Work.att_road=b.field_code AND b.code_code=a.field_code " & _
                        '" LEFT JOIN BDP210 c ON Danger_Work.work_type1=c.field_code AND c.code_code='work_type1' " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM Danger_Work WHERE dan_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM Danger_Work WHERE dan_code='" & RepSql(sPara1) & "'"

                End Select

            Case "LSI015A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT Danger_check.*,BDP210.field_name as danc_steps FROM Danger_check" & _
                               " LEFT JOIN BDP210 ON Danger_check.danc_step=BDP210.field_code AND BDP210.code_code='danc_step' " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM Danger_check WHERE danc_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM Danger_check WHERE danc_code='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI016A"
                Select Case sType
                    Case "SELECT"
                        sSql = " SELECT e_news.*,BDP210.field_name as news_types,BDP080.usr_name FROM e_news " & _
                               " LEFT JOIN BDP210 ON e_news.news_type=BDP210.field_code AND BDP210.code_code='news_type'" & _
                               " LEFT JOIN BDP080 ON e_news.usr_code = BDP080.usr_code " & _
                               sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM e_paper_d WHERE news_code=@datakeys;DELETE FROM e_news WHERE news_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM e_news WHERE news_code='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI017A"
                Select Case sType
                    Case "SELECT"
                        sSql = " SELECT APP_VERSION.*, BDP210.field_name as app_types FROM APP_VERSION " & _
                               "  LEFT JOIN BDP210 ON APP_VERSION.app_type=BDP210.field_code AND BDP210.code_code='app_type' " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM APP_VERSION WHERE app_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM APP_VERSION WHERE app_code='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI018A"
                Select Case sType
                    Case "SELECT"
                        sSql = " SELECT APP_LOG.*, BDP230.dep_name, BDP210.field_name as app_types FROM APP_LOG " & _
                               " LEFT JOIN BDP210 ON APP_LOG.app_type=BDP210.field_code AND BDP210.code_code='app_type' " & _
                               " LEFT JOIN BDP080 ON APP_LOG.usr_code = BDP080.usr_code  " & _
                               " LEFT JOIN BDP230 ON BDP080.dep_code = BDP230.dep_code " & sPara1 & " ORDER BY apl_code DESC"
                        'sSql = "SELECT BDP200.*,BDP030.prg_name FROM BDP200 LEFT JOIN BDP030 ON BDP200.prg_code=BDP030.prg_code " & sPara1 & " ORDER BY usr_date DESC,usr_time DESC"

                    Case "DELETE"
                        sSql = "DELETE FROM APP_LOG WHERE apl_code=@datakeys"

                End Select

            Case "LSI019A"
                Select Case sType
                    Case "SELECT"
                        sSql = " SELECT * FROM LABOR_INSPECT " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM LABOR_INSPECT WHERE labor_inspect=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM LABOR_INSPECT WHERE labor_inspect=" & RepSql(sPara1)
                End Select

            Case "LSI023A" '20160803 karry新增
                Select Case sType
                    Case "SELECT"
                        sSql = " SELECT * FROM [work_tb] " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM [work_tb] WHERE workid=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM [work_tb] WHERE workid=" & RepSql(sPara1)
                End Select

            Case "LSI021A" '20150709 allen新增
                Select Case sType
                    Case "SELECT"
                        'Dim Ddate As Integer = Val(Get_SysPara("disable_date", "150"))
                        'sPara1 &= " AND (is_agree<>'N' OR ins_date>='" & Format(Now.AddDays(0 - Ddate), "yyyy/MM/dd") & "') "

                        sSql = "SELECT Roof_work.*,a.field_name AS area_name,b.field_name AS road_name" & _
                               " ,REPLACE(REPLACE(REPLACE(REPLACE(roof_tool1" & _
                               " ,'A','使用施工架或高空工作車'),'B','安全網'),'C','安全帽'),'D','防墜器、安全帶') as roof_tool1_c" & _
                               " ,REPLACE(REPLACE(tear_tool1" & _
                               " ,'A','扶手先行工法'),'B','安全母索支柱工法') as tear_tool1_c" & _
                               " ,REPLACE(REPLACE(REPLACE(REPLACE(cage_tool1" & _
                               " ,'A','安全帽'),'B','背負式安全帶'),'C','防墜器'),'D','雙繩作業+防墜器') as cage_tool1_c" & _
                               " from Roof_work" & _
                               " LEFT JOIN BDP210 a ON Roof_work.att_area=a.field_code AND a.code_code='ZipCode'" & _
                               " LEFT JOIN BDP210 b ON Roof_work.att_road=b.field_code AND b.code_code=a.field_code" & sPara1 & sPara2
                    Case "DELETE"
                        sSql = "DELETE FROM Roof_work WHERE roof_cod=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM Roof_work WHERE roof_cod='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI024A" '20150709 allen新增
                Select Case sType
                    Case "SELECT"
                        'Dim Ddate As Integer = Val(Get_SysPara("disable_date", "150"))
                        'sPara1 &= " AND (is_agree<>'N' OR ins_date>='" & Format(Now.AddDays(0 - Ddate), "yyyy/MM/dd") & "') "

                        sSql = "SELECT Precise_Check.*,a.field_name AS area_name,b.field_name AS road_name" & _
                               " ,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(pc_kind1" & _
                               " ,'A','異常氣壓作業'),'B','高度7公尺以上模板組、拆作業'),'C','高度2公尺以上擋土支撐組、拆作業')" & _
                               " ,'D','外牆施工架組立、拆除'),'E','室內挑高排架'),'F','鋼骨工程吊料平臺爬升')" & _
                               " ,'G','地下結構物模板、鋼筋吊料作業') as pc_kind1_c" & _
							   " ,REPLACE(REPLACE(REPLACE(pc_kind2" & _
                               " ,'A','升降機組裝(施工電梯)'),'B','塔式起重機升高及拆卸作業'),'C','人字臂起重桿組、拆作業" & _
                               " ') as pc_kind2_c" & _
                               " FROM Precise_Check" & _
                               " LEFT JOIN BDP210 a ON Precise_Check.eng_area=a.field_code AND a.code_code='ZipCode'" & _
                               " LEFT JOIN BDP210 b ON Precise_Check.eng_road=b.field_code AND b.code_code=a.field_code" & sPara1 & sPara2
                    Case "DELETE"
                        sSql = "DELETE FROM Precise_Check WHERE pc_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM Precise_Check WHERE pc_code='" & RepSql(sPara1) & "'"
                End Select
				
				
            Case "LSI025" '20211102 Ken 新增
                Select Case sType
                    Case "SELECT"
                        'Dim Ddate As Integer = Val(Get_SysPara("disable_date", "150"))
                        'sPara1 &= " AND (is_agree<>'N' OR ins_date>='" & Format(Now.AddDays(0 - Ddate), "yyyy/MM/dd") & "') "

                        sSql = "SELECT LSI025.*,a.field_name AS area_name,b.field_name AS road_name" & _
                               " ,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(pc_kind1" & _
                               " ,'A','異常氣壓作業'),'B','高度7公尺以上模板組、拆作業'),'C','高度2公尺以上擋土支撐組、拆作業')" & _
                               " ,'D','外牆施工架組立、拆除'),'E','室內挑高排架'),'F','鋼骨工程吊料平臺爬升')" & _
                               " ,'G','地下結構物模板、鋼筋吊料作業') as pc_kind1_c" & _
							   " ,REPLACE(REPLACE(REPLACE(pc_kind2" & _
                               " ,'A','升降機組裝(施工電梯)'),'B','塔式起重機升高及拆卸作業'),'C','人字臂起重桿組、拆作業" & _
                               " ') as pc_kind2_c" & _
                               " FROM LSI025" & _
                               " LEFT JOIN BDP210 a ON LSI025.eng_area=a.field_code AND a.code_code='ZipCode'" & _
                               " LEFT JOIN BDP210 b ON LSI025.eng_road=b.field_code AND b.code_code=a.field_code" & sPara1 & sPara2
                    Case "DELETE"
                        sSql = "DELETE FROM LSI025 WHERE pc_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM LSI025 WHERE pc_code='" & RepSql(sPara1) & "'"
                End Select

             Case "LSI026" '20211102 Ken 新增
                Select Case sType
                    Case "SELECT"
                        'Dim Ddate As Integer = Val(Get_SysPara("disable_date", "150"))
                        'sPara1 &= " AND (is_agree<>'N' OR ins_date>='" & Format(Now.AddDays(0 - Ddate), "yyyy/MM/dd") & "') "

                        sSql = "SELECT LSI026.*,a.field_name AS area_name,b.field_name AS road_name" & _
                               " ,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(pc_kind1" & _
                               " ,'A','異常氣壓作業'),'B','高度7公尺以上模板組、拆作業'),'C','高度2公尺以上擋土支撐組、拆作業')" & _
                               " ,'D','外牆施工架組立、拆除'),'E','室內挑高排架'),'F','鋼骨工程吊料平臺爬升')" & _
                               " ,'G','地下結構物模板、鋼筋吊料作業') as pc_kind1_c" & _
							   " ,REPLACE(REPLACE(REPLACE(pc_kind2" & _
                               " ,'A','升降機組裝(施工電梯)'),'B','塔式起重機升高及拆卸作業'),'C','人字臂起重桿組、拆作業" & _
                               " ') as pc_kind2_c" & _
                               " FROM LSI026" & _
                               " LEFT JOIN BDP210 a ON LSI026.eng_area=a.field_code AND a.code_code='ZipCode'" & _
                               " LEFT JOIN BDP210 b ON LSI026.eng_road=b.field_code AND b.code_code=a.field_code" & sPara1 & sPara2
                    Case "DELETE"
                        sSql = "DELETE FROM LSI026 WHERE pc_code=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM LSI026 WHERE pc_code='" & RepSql(sPara1) & "'"
                End Select

            Case "Board"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT *,CASE WHEN LEN(theme) > 25 THEN SUBSTRING(theme, 1 , 25) + '...' ELSE theme END AS theme1," & _
                               "(SELECT COUNT(*) FROM Board WHERE bull_code=a.bull_code AND bull_type<>'0') AS cnt," & _
                               "(SELECT TOP 1 Ins_date FROM Board WHERE bull_code=a.bull_code AND bull_type<>'0' ORDER BY Ins_date DESC,Ins_time DESC) as last " & _
                               "FROM Board AS a WHERE bull_type='0' " & sPara1 & "  AND is_hid='Y' ORDER BY Ins_date DESC,Ins_time DESC"

                    Case "DELETE"
                        sSql = "DELETE FROM Board WHERE bdp190=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM Board WHERE bdp190='" & RepSql(sPara1) & "'"
                End Select

            Case "News"
                sSql = "SELECT DISTINCT e_paper.paper_code,e_paper.paper_name,e_paper.paper_num,e_paper.ins_date FROM e_paper" & _
                       " LEFT JOIN e_paper_d ON e_paper_d.paper_code=e_paper.paper_code" & _
                       " LEFT JOIN e_news ON e_news.news_code=e_paper_d.news_code " & sPara1 & " " & RepSql(sPara2)

            Case "Meeting"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT mrd_code,mrd_name,mrd_date,room_name,mrd_time_s+' - '+mrd_time_e as mrd_time" & _
                               " ,mrd_memo FROM MEETING_ORDER" & _
                               " LEFT JOIN MEETING_ROOM ON MEETING_ORDER.room_code=MEETING_ROOM.room_code" & _
                               " WHERE is_show='Y' " & sPara1 & " ORDER BY mrd_date DESC,mrd_time DESC"
                End Select

            Case "OTB030A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT * FROM OTB030 " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM OTB030 WHERE otb030=@datakeys"

                    Case "Detail"
                        sSql = "SELECT * FROM OTB030 WHERE otb030='" & RepSql(sPara1) & "'"
                End Select

            Case "OTB050A"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT * FROM OTB040 " & sPara1 & sPara2

                    Case "DELETE"
                        sSql = "DELETE FROM OTB040 WHERE tn_code=@datakeys;DELETE FROM OTB041 WHERE tn_code=@datakeys;"

                    Case "Detail"
                        sSql = "SELECT * FROM OTB040 WHERE tn_code='" & RepSql(sPara1) & "'"
                End Select

            Case "OTB041A_detail"
                Select Case sType
                    Case "SELECT"
                        sSql = "SELECT OTB041.*,OTB020.tsh_name FROM OTB041 LEFT JOIN OTB020 ON OTB041.tsh_code=OTB020.tsh_code " & _
                               "WHERE tn_code ='" & RepSql(sPara1) & "'" & sPara2 & " ORDER BY scr_no"

                    Case "DELETE"
                        sSql = "DELETE FROM OTB041 WHERE otb041=@datakeys;"

                    Case "Detail"
                        sSql = "SELECT * FROM OTB041 WHERE otb041='" & RepSql(sPara1) & "'"
                End Select
        End Select

        Return sSql
    End Function

    ''' <summary>
    ''' 檢查資料異動權限
    ''' </summary>
    ''' <param name="sUsrCodeL">登入使用者</param>
    ''' <param name="sUsrCodeP">資料擁有者</param>
    ''' <returns>true:有異動權限，false：無異動權限</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function Chk_UsrLimit(ByVal sUsrCodeL As String, ByVal sUsrCodeP As String, ByVal sGrpCodeL As String) As Boolean
        Dim sSql As String = ""
        Dim Dt_fun1 As New DataTable
        Dim Dt_fun2 As New DataTable
        Dim tmpAdpt As SqlDataAdapter
        Dim SqlCmd As New SqlCommand
        Dim SqlCon As SqlConnection = New SqlConnection(Get_ConnStr(""))

        'sSql = "SELECT dep_code FROM BDP080 WHERE usr_code='" & sUsrCodeL & "'"
        sSql = "SELECT dep_code FROM BDP080 WHERE usr_code=@UsrCodeL"
        SqlCmd = New SqlCommand(sSql, SqlCon)
        SqlCmd.Parameters.AddWithValue("@UsrCodeL", RepSql(sUsrCodeL))
        tmpAdpt = New SqlDataAdapter(SqlCmd)
        tmpAdpt.Fill(Dt_fun1)
        'Dt_fun1 = Me.Get_DataTable(sSql)

        'sSql = "SELECT dep_code FROM BDP080 WHERE usr_code='" & sUsrCodeP & "'"
        sSql = "SELECT dep_code FROM BDP080 WHERE usr_code=@UsrCodeP"
        SqlCmd = New SqlCommand(sSql, SqlCon)
        SqlCmd.Parameters.AddWithValue("@UsrCodeP", RepSql(sUsrCodeP))
        tmpAdpt = New SqlDataAdapter(SqlCmd)
        tmpAdpt.Fill(Dt_fun2)
        'Dt_fun2 = Me.Get_DataTable(sSql)

        '沒有資料擁有者時預設為可異動
        If Not Dt_fun2.Rows.Count > 0 Then
            Return True
        End If


        If sGrpCodeL = "G002" Then
            '科室相同則視為有異動權限
            If Dt_fun1.Rows.Count > 0 Then
                If (Dt_fun1.Rows(0).Item("dep_code") = Dt_fun2.Rows(0).Item("dep_code")) Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End If

        If sGrpCodeL = "G003" Then
            '科室及帳號相同則視為有異動權限
            If Dt_fun1.Rows.Count > 0 Then
                If (Dt_fun1.Rows(0).Item("dep_code") = Dt_fun2.Rows(0).Item("dep_code")) And (sUsrCodeL = sUsrCodeP) Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End If
    End Function


    ''' <summary>
    ''' 傳入程式代碼等參數檢查是否有重覆資料
    ''' </summary>
    ''' <param name="sPrgcode">程式代碼</param>
    ''' <param name="sType">種類</param>
    ''' <param name="sPara1">參數1</param>
    ''' <param name="sPara2">參數2</param>
    ''' <param name="sPara3">參數3</param>
    ''' <param name="sPara4">參數4</param>
    ''' <param name="sPara5">參數5</param>
    ''' <param name="sPara6">參數6</param>
    ''' <returns>true:沒有重覆，false：有重覆</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function Chk_RelData(ByVal sPrgcode As String, ByVal sType As String, ByVal sPara1 As String, ByVal sPara2 As String, ByVal sPara3 As String, ByVal sPara4 As String, ByVal sPara5 As String, ByVal sPara6 As String) As Boolean
        Dim sSql As String = ""
        Dim Dt_fun1 As New DataTable
        Dim tmpAdpt As SqlDataAdapter
        Dim SqlCmd As New SqlCommand
        Dim SqlCon As SqlConnection = New SqlConnection(Get_ConnStr(""))

        Select Case sPrgcode
            Case "BDP000A"
                'sSql = "SELECT * FROM BDP000 WHERE par_name='" & RepSql(sPara1) & "'"
                sSql = "SELECT * FROM BDP000 WHERE par_name=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "BDP070A"
                'sSql = "SELECT * FROM BDP070 WHERE grp_code='" & RepSql(sPara1) & "'"
                sSql = "SELECT * FROM BDP070 WHERE grp_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "BDP080A"
                Select Case sType
                    Case "IDno"
                        'sSql = "SELECT * FROM BDP080 WHERE usr_idno='" & RepSql(sPara1) & "'"
                        sSql = "SELECT * FROM BDP080 WHERE usr_idno=@Para1"
                        SqlCmd = New SqlCommand(sSql, SqlCon)
                        SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                    Case Else
                        'sSql = "SELECT * FROM BDP080 WHERE usr_code='" & RepSql(sPara1) & "'"
                        sSql = "SELECT * FROM BDP080 WHERE usr_code=@Para1"
                        SqlCmd = New SqlCommand(sSql, SqlCon)
                        SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                End Select

            Case "BDP030A"
                'sSql = "SELECT * FROM BDP030 WHERE prg_code='" & RepSql(sPara1) & "'"
                sSql = "SELECT * FROM BDP030 WHERE prg_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "BDP210A"
                'sSql = "SELECT * FROM BDP210 WHERE code_code='" & RepSql(sPara1) & "' and field_name ='" & RepSql(sPara2) & "'"
                sSql = "SELECT * FROM BDP210 WHERE code_code=@Para1 and field_name =@Para2"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                SqlCmd.Parameters.AddWithValue("@Para2", RepSql(sPara2))

            Case "BDP230A"
                'sSql = "SELECT * FROM BDP230 WHERE dep_code='" & RepSql(sPara1) & "'"
                sSql = "SELECT * FROM BDP230 WHERE dep_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "LSI001A"
                'sSql = "SELECT * FROM mail_asklaw WHERE ask_code='" & RepSql(sPara1) & "'"
                sSql = "SELECT * FROM mail_asklaw WHERE ask_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "LSI002A"
                'sSql = "SELECT * FROM mail_complain WHERE cmp_code='" & RepSql(sPara1) & "'"
                sSql = "SELECT * FROM mail_complain WHERE cmp_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "yumi2" '20160912 yumi新增
                sSql = "SELECT * FROM yumi2 WHERE book_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "yumi2-2" '20160912 yumi新增
                sSql = "SELECT * FROM yumi2_2 WHERE ord_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "LSI022A"
                'sSql = "SELECT * FROM mail_complain WHERE cmp_code='" & RepSql(sPara1) & "'"
                sSql = "SELECT * FROM mail WHERE cmp_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "LSI003A"
                'sSql = "SELECT * FROM E_PAPER_ORDER WHERE ord_code='" & RepSql(sPara1) & "'"
                sSql = "SELECT * FROM E_PAPER_ORDER WHERE ord_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "LSI004A"
                'sSql = "SELECT * FROM e_paper WHERE paper_code='" & RepSql(sPara1) & "'"
                sSql = "SELECT * FROM e_paper WHERE paper_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "LSI004A_detail"
                'sSql = "SELECT * FROM e_paper_d WHERE paper_code='" & RepSql(sPara1) & "' AND news_code='" & RepSql(sPara2) & "'"
                sSql = "SELECT * FROM e_paper_d WHERE paper_code=@Para1 AND news_code=@Para2"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                SqlCmd.Parameters.AddWithValue("@Para2", RepSql(sPara2))

            Case "LSI005A"
                'sSql = "SELECT * FROM e_paper_sch WHERE sch_code='" & RepSql(sPara1) & "' AND count=0"
                sSql = "SELECT * FROM e_paper_sch WHERE sch_code=@Para1 AND count=0"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "LSI007A"
                'sSql = "SELECT * FROM Meeting_room WHERE room_code='" & RepSql(sPara1) & "'"
                sSql = "SELECT * FROM Meeting_room WHERE room_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "LSI008A"
                Select Case sType
                    Case "MrdCode"
                        'sSql = "SELECT * FROM Meeting_order WHERE mrd_code='" & RepSql(sPara1) & "'"
                        sSql = "SELECT * FROM Meeting_order WHERE mrd_code=@Para1"
                        SqlCmd = New SqlCommand(sSql, SqlCon)
                        SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                    Case "MrdTime"
                        'sSql = "SELECT * FROM Meeting_order WHERE mrd_code<>'" & RepSql(sPara1) & "'" & _
                        '       " AND room_code='" & RepSql(sPara2) & "' AND mrd_date='" & RepSql(sPara3) & "'" & _
                        '       " AND (mrd_time_e>'" & RepSql(sPara4) & "' AND mrd_time_s<'" & RepSql(sPara5) & "')"
                        sSql = "SELECT * FROM Meeting_order WHERE mrd_code<>@Para1" & _
                               " AND room_code=@Para2 AND mrd_date=@Para3" & _
                               " AND (mrd_time_e>@Para4 AND mrd_time_s<@Para5)"
                        SqlCmd = New SqlCommand(sSql, SqlCon)
                        SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                        SqlCmd.Parameters.AddWithValue("@Para2", RepSql(sPara2))
                        SqlCmd.Parameters.AddWithValue("@Para3", RepSql(sPara3))
                        SqlCmd.Parameters.AddWithValue("@Para4", RepSql(sPara4))
                        SqlCmd.Parameters.AddWithValue("@Para5", RepSql(sPara5))
                End Select

            Case "LSI008A_detail"
                'sSql = "SELECT * FROM Meeting_users WHERE mrd_code='" & RepSql(sPara1) & "' AND usr_code='" & RepSql(sPara2) & "'" & _
                '       " AND usr_name='" & RepSql(sPara3) & "' AND usr_mail='" & RepSql(sPara4) & "'"
                sSql = "SELECT * FROM Meeting_users WHERE mrd_code=@Para1 AND usr_code=@Para2" & _
                       " AND usr_name=@Para3 AND usr_mail=@Para4"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                SqlCmd.Parameters.AddWithValue("@Para2", RepSql(sPara2))
                SqlCmd.Parameters.AddWithValue("@Para3", RepSql(sPara3))
                SqlCmd.Parameters.AddWithValue("@Para4", RepSql(sPara4))

            Case "LSI009A"
                Select Case sType
                    Case "stop_code"
                        'sSql = "SELECT * FROM Work_stop WHERE stop_code='" & RepSql(sPara1) & "'"
                        sSql = "SELECT * FROM Work_stop WHERE stop_code=@Para1"
                        SqlCmd = New SqlCommand(sSql, SqlCon)
                        SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                    Case "stop_number"
                        'sSql = "SELECT * FROM Work_stop WHERE stop_number='" & RepSql(sPara1) & "'"
                        sSql = "SELECT * FROM Work_stop WHERE stop_number=@Para1"
                        SqlCmd = New SqlCommand(sSql, SqlCon)
                        SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                End Select

            Case "LSI010A"
                Select Case sType
                    Case "start_code"
                        'sSql = "SELECT * FROM Work_start WHERE start_code='" & RepSql(sPara1) & "'"
                        sSql = "SELECT * FROM Work_start WHERE start_code=@Para1"
                        SqlCmd = New SqlCommand(sSql, SqlCon)
                        SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                    Case "stop_number"
                        'sSql = "SELECT * FROM Work_start WHERE stop_number='" & RepSql(sPara1) & "'"
                        sSql = "SELECT * FROM Work_start WHERE stop_number=@Para1"
                        SqlCmd = New SqlCommand(sSql, SqlCon)
                        SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

                End Select

            Case "LSI011A"
                'sSql = "SELECT * FROM Self_check WHERE self_code='" & RepSql(sPara1) & "'"
                sSql = "SELECT * FROM Self_check WHERE self_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "LSI012A"
                'sSql = "SELECT * FROM Small_space WHERE sma_code='" & RepSql(sPara1) & "'"
                sSql = "SELECT * FROM Small_space WHERE sma_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "LSI013A"
                'sSql = "SELECT * FROM Fix_work WHERE fix_code='" & RepSql(sPara1) & "'"
                sSql = "SELECT * FROM Fix_work WHERE fix_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "LSI020A" 'may 2013/10/15
                'sSql = "SELECT * FROM Fix_work WHERE fix_code='" & RepSql(sPara1) & "'"
                sSql = "SELECT * FROM SCHOOL_WORK WHERE fix_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "LSI014A"
                'sSql = "SELECT * FROM Danger_Work WHERE dan_code='" & RepSql(sPara1) & "'"
                sSql = "SELECT * FROM Danger_Work WHERE dan_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "LSI015A"
                'sSql = "SELECT * FROM Danger_check WHERE danc_code='" & RepSql(sPara1) & "'"
                sSql = "SELECT * FROM Danger_check WHERE danc_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "LSI016A"
                'sSql = "SELECT * FROM E_NEWS WHERE news_code='" & RepSql(sPara1) & "'"
                sSql = "SELECT * FROM E_NEWS WHERE news_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "LSI017A"
                'sSql = "SELECT * FROM APP_VERSION WHERE app_code='" & RepSql(sPara1) & "'"
                sSql = "SELECT * FROM APP_VERSION WHERE app_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "LSI021A" '20150709 allen 新增
                sSql = "SELECT * FROM Roof_work WHERE roof_cod=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "LSI024A"
                sSql = "SELECT * FROM Precise_Check WHERE pc_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
            Case "LSI025" '20211102 Ken 新增
                sSql = "SELECT * FROM LSI025 WHERE pc_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
				
            Case "LSI026" '20211102 Ken 新增
                sSql = "SELECT * FROM LSI026 WHERE pc_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))



            Case "SendNewsMail"
                sSql = "SELECT * FROM E_PAPER_LOG WHERE sch_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "DanSearch"
                'sSql = "SELECT * FROM Danger_check WHERE con_email='" & RepSql(sPara1) & "' AND con_pw='" & RepSql(sPara2) & "'"
                sSql = "SELECT * FROM Danger_check WHERE con_email=@Para1 AND con_pw=@Para2"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                SqlCmd.Parameters.AddWithValue("@Para2", RepSql(sPara2))

            Case "ComSearch"
                'sSql = "SELECT * FROM Danger_check WHERE con_email='" & RepSql(sPara1) & "' AND con_pw='" & RepSql(sPara2) & "'"
                sSql = "SELECT * FROM MAIL_COMPLAIN WHERE cmp_email=@Para1 AND cmp_pass=@Para2"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
                SqlCmd.Parameters.AddWithValue("@Para2", RepSql(sPara2))

            Case "MyService"
                'sSql = "SELECT * FROM GEO_LOCATION WHERE ms_code='" & RepSql(sPara1) & "'"
                sSql = "SELECT * FROM GEO_LOCATION WHERE ms_code=@Para1"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))

            Case "app_update"
                'sSql = "SELECT * FROM APP_VERSION WHERE version>'" & RepSql(sPara1) & "' AND app_type='1'"
                sSql = "SELECT * FROM APP_VERSION WHERE version>@Para1 AND app_type='1'"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@Para1", RepSql(sPara1))
        End Select

        tmpAdpt = New SqlDataAdapter(SqlCmd)
        tmpAdpt.Fill(Dt_fun1)
        'Dt_fun1 = Me.Get_DataTable(sSql)
        If Dt_fun1.Rows.Count > 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    ''' <summary>
    ''' 資料存檔
    ''' </summary>
    ''' <param name="sPrgcode">程式代碼</param>
    ''' <param name="sType">種類-add/copy/mdy</param>
    ''' <param name="sPara1">參數1~參數60</param>
    ''' <returns>true:存檔成功，false：存檔失敗</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function SaveEditData(ByVal sPrgcode As String, ByVal sType As String, ByVal sPara1 As String, ByVal sPara2 As String, ByVal sPara3 As String, ByVal sPara4 As String, ByVal sPara5 As String, ByVal sPara6 As String, ByVal sPara7 As String, ByVal sPara8 As String, ByVal sPara9 As String, ByVal sPara10 As String _
                                 , ByVal sPara11 As String, ByVal sPara12 As String, ByVal sPara13 As String, ByVal sPara14 As String, ByVal sPara15 As String, ByVal sPara16 As String, ByVal sPara17 As String, ByVal sPara18 As String, ByVal sPara19 As String, ByVal sPara20 As String _
                                 , ByVal sPara21 As String, ByVal sPara22 As String, ByVal sPara23 As String, ByVal sPara24 As String, ByVal sPara25 As String, ByVal sPara26 As String, ByVal sPara27 As String, ByVal sPara28 As String, ByVal sPara29 As String, ByVal sPara30 As String _
                                 , ByVal sPara31 As String, ByVal sPara32 As String, ByVal sPara33 As String, ByVal sPara34 As String, ByVal sPara35 As String, ByVal sPara36 As String, ByVal sPara37 As String, ByVal sPara38 As String, ByVal sPara39 As String, ByVal sPara40 As String _
                                 , ByVal sPara41 As String, ByVal sPara42 As String, ByVal sPara43 As String, ByVal sPara44 As String, ByVal sPara45 As String, ByVal sPara46 As String, ByVal sPara47 As String, ByVal sPara48 As String, ByVal sPara49 As String, ByVal sPara50 As String) As Boolean
        Dim sSql As String = ""
        'Dim SqlCmd As New SqlCommand
        'Dim SqlCon As SqlConnection = Me.ConnectDB

        Select Case sPrgcode
            Case "BDP000A"
                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO BDP000(" & _
                               "par_name,par_value,par_memo" & _
                               ") VALUES (N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "')"
                        'sSql = "INSERT INTO BDP000(" & _
                        '       "par_name,par_value,par_memo" & _
                        '       ") VALUES (@sPara2,@sPara3,@sPara4)"
                        'SqlCmd = New SqlCommand(sSql, SqlCon)
                        'SqlCmd.Parameters.AddWithValue("@sPara2", RepSql(sPara2))
                        'SqlCmd.Parameters.AddWithValue("@sPara3", RepSql(sPara3))
                        'SqlCmd.Parameters.AddWithValue("@sPara4", RepSql(sPara4))

                    Case "MDY"
                        sSql = "UPDATE BDP000 SET " & _
                               "    par_name=N'" & RepSql(sPara2) & "'," & _
                               "    par_value=N'" & RepSql(sPara3) & "'," & _
                               "    par_memo=N'" & RepSql(sPara4) & "'" & _
                               " WHERE bdp000='" & RepSql(sPara1) & "'"
                        'sSql = "UPDATE BDP000 SET " & _
                        '       "    par_name=@sPara2," & _
                        '       "    par_value=@sPara3," & _
                        '       "    par_memo=@sPara4" & _
                        '       " WHERE bdp000=@sPara1"
                        'SqlCmd = New SqlCommand(sSql, SqlCon)
                        'SqlCmd.Parameters.AddWithValue("@sPara1", RepSql(sPara1))
                        'SqlCmd.Parameters.AddWithValue("@sPara2", RepSql(sPara2))
                        'SqlCmd.Parameters.AddWithValue("@sPara3", RepSql(sPara3))
                        'SqlCmd.Parameters.AddWithValue("@sPara4", RepSql(sPara4))
                End Select

            Case "BDP002A"
                Select Case sType
                    Case "INSERT"
                        sSql = "INSERT INTO BDP091(" & _
                               "grp_code,prg_code,limit_str,is_use" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "')"
                        'sSql = "INSERT INTO BDP091(" & _
                        '       "grp_code,prg_code,limit_str,is_use" & _
                        '       ") VALUES (@sPara1,@sPara2,@sPara3,@sPara4)"
                        'SqlCmd = New SqlCommand(sSql, SqlCon)
                        'SqlCmd.Parameters.AddWithValue("@sPara1", RepSql(sPara1))
                        'SqlCmd.Parameters.AddWithValue("@sPara2", RepSql(sPara2))
                        'SqlCmd.Parameters.AddWithValue("@sPara3", RepSql(sPara3))
                        'SqlCmd.Parameters.AddWithValue("@sPara4", RepSql(sPara4))

                    Case "DELETE"
                        sSql = "DELETE FROM BDP091 WHERE grp_code='" & RepSql(sPara1) & "'"
                        'sSql = "DELETE FROM BDP091 WHERE grp_code=@sPara1"
                        'SqlCmd = New SqlCommand(sSql, SqlCon)
                        'SqlCmd.Parameters.AddWithValue("@sPara1", RepSql(sPara1))
                End Select

            Case "BDP030A"
                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO BDP030(" & _
                               "prg_code,prg_name,prg_type,is_use,is_group,scr_no,menu_level,up_level,prg_folder," & _
                               "s_date,e_date " & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "'," & _
                               RepSql(sPara6) & ",N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "',N'" & _
                               RepSql(sPara9) & "',N'" & _
                               RepSql(sPara10) & "',N'" & _
                               RepSql(sPara11) & "')"

                    Case "MDY"
                        sSql = "UPDATE BDP030 SET " & _
                               "    prg_name=N'" & RepSql(sPara2) & "'," & _
                               "    prg_type=N'" & RepSql(sPara3) & "'," & _
                               "    is_use=N'" & RepSql(sPara4) & "'," & _
                               "    is_group=N'" & RepSql(sPara5) & "'," & _
                               "    scr_no=" & RepSql(sPara6) & "," & _
                               "    menu_level=N'" & RepSql(sPara7) & "'," & _
                               "    up_level=N'" & RepSql(sPara8) & "'," & _
                               "    s_date=N'" & RepSql(sPara10) & "'," & _
                               "    e_date=N'" & RepSql(sPara11) & "'," & _
                               "    prg_folder=N'" & RepSql(sPara9) & "'" & _
                               " WHERE prg_code='" & RepSql(sPara1) & "'"
                End Select
            Case "BDP003A"
                Select Case sType
                    Case "INSERT"
                        sSql = "INSERT INTO BDP090(" & _
                               "usr_code,prg_code,limit_str,is_use" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "')"

                    Case "DELETE"
                        sSql = "DELETE FROM BDP090 WHERE usr_code='" & RepSql(sPara1) & "'"

                    Case "UPDATE"
                        sSql = "UPDATE BDP080 SET Limit_type='A' WHERE usr_code='" & RepSql(sPara1) & "'"
                End Select

            Case "BDP003B"
                Select Case sType
                    Case "INSERT"
                        sSql = "INSERT INTO BDP092(" & _
                               "dep_code,prg_code,limit_str,is_use" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "')"

                    Case "DELETE"
                        sSql = "DELETE FROM BDP092 WHERE dep_code='" & RepSql(sPara1) & "'"

                    Case "UPDATE"
                        sSql = "UPDATE BDP080 SET Limit_type='C' WHERE dep_code='" & RepSql(sPara1) & "'"
                End Select

            Case "BDP006A"
                sSql = "UPDATE BDP080 SET usr_pass='" & RepSql(sPara2) & "' WHERE usr_code ='" & RepSql(sPara1) & "'"

            Case "BDP007A"
                sSql = "UPDATE BDP000 SET par_value='" & RepSql(sPara1) & "' WHERE par_name ='css_sys'"

            Case "BDP070A"
                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO BDP070(" & _
                               "grp_code,grp_name,is_use" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "')"

                    Case "MDY"
                        sSql = "UPDATE BDP070 SET " & _
                               "    grp_name=N'" & RepSql(sPara2) & "'," & _
                               "    is_use=N'" & RepSql(sPara3) & "'" & _
                               " WHERE grp_code='" & RepSql(sPara1) & "'"
                End Select

            Case "BDP080A"
                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO BDP080(" & _
                               "usr_code,usr_name,usr_pass,grp_code,Limit_type," & _
                               "usr_mail,is_use,usr_memo,dep_code,limit_level,usr_idno,count" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "',N'" & _
                               RepSql(sPara9) & "',N'" & _
                               RepSql(sPara10) & "',N'" & _
                               RepSql(sPara11) & "',0)"

                    Case "MDY"
                        sSql = "UPDATE BDP080 SET " & _
                               "    usr_name=N'" & RepSql(sPara2) & "'," & _
                               "    usr_pass=N'" & RepSql(sPara3) & "'," & _
                               "    grp_code=N'" & RepSql(sPara4) & "'," & _
                               "    Limit_type=N'" & RepSql(sPara5) & "'," & _
                               "    usr_mail=N'" & RepSql(sPara6) & "'," & _
                               "    is_use=N'" & RepSql(sPara7) & "'," & _
                               "    usr_memo=N'" & RepSql(sPara8) & "'," & _
                               "    dep_code=N'" & RepSql(sPara9) & "'," & _
                               "    limit_level=N'" & RepSql(sPara10) & "'" & _
                               " WHERE usr_code='" & RepSql(sPara1) & "'"
                End Select

            Case "BDP140A"
                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO BDP140(" & _
                               "scr_no,prg_code,field_code,field_name,field_type" & _
                               ") VALUES (" & _
                               RepSql(sPara2) & ",N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "')"

                    Case "MDY"
                        sSql = "UPDATE BDP140 SET " & _
                               "    scr_no=" & RepSql(sPara2) & "," & _
                               "    prg_code=N'" & RepSql(sPara3) & "'," & _
                               "    field_code=N'" & RepSql(sPara4) & "'," & _
                               "    field_name=N'" & RepSql(sPara5) & "'," & _
                               "    field_type=N'" & RepSql(sPara6) & "'" & _
                               " WHERE bdp140=" & RepSql(sPara1)
                End Select

            Case "BDP150A"
                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO BDP150(" & _
                               "ip_type,ip_name,ip_s1,ip_s2,ip_s3,ip_s4,ip_e1,ip_e2,ip_e3,ip_e4" & _
                               ") VALUES (N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "'," & _
                               RepSql(sPara4) & "," & _
                               RepSql(sPara5) & "," & _
                               RepSql(sPara6) & "," & _
                               RepSql(sPara7) & "," & _
                               RepSql(sPara8) & "," & _
                               RepSql(sPara9) & "," & _
                               RepSql(sPara10) & "," & _
                               RepSql(sPara11) & ")"

                    Case "MDY"
                        sSql = "UPDATE BDP150 SET " & _
                               "    ip_type=N'" & RepSql(sPara2) & "'," & _
                               "    ip_name=N'" & RepSql(sPara3) & "'," & _
                               "    ip_s1=" & RepSql(sPara4) & "," & _
                               "    ip_s2=" & RepSql(sPara5) & "," & _
                               "    ip_s3=" & RepSql(sPara6) & "," & _
                               "    ip_s4=" & RepSql(sPara7) & "," & _
                               "    ip_e1=" & RepSql(sPara8) & "," & _
                               "    ip_e2=" & RepSql(sPara9) & "," & _
                               "    ip_e3=" & RepSql(sPara10) & "," & _
                               "    ip_e4=" & RepSql(sPara11) & "" & _
                               " WHERE serial_no=" & RepSql(sPara1)
                End Select

            Case "BDP190A"
                Select Case sType
                    Case "MDY"
                        sSql = "UPDATE Board SET " & _
                               "    is_hid=N'" & RepSql(sPara2) & "'" & _
                               " WHERE bull_code='" & RepSql(sPara1) & "' and bull_type='0'"

                End Select

            Case "BDP190A_detail"
                Select Case sType
                    Case "MDY"
                        sSql = "UPDATE Board SET " & _
                               "    is_hid=N'" & RepSql(sPara2) & "'" & _
                               " WHERE bdp190=" & RepSql(sPara1)

                End Select


            Case "BDP210A"
                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO BDP210(" & _
                               "Code_code,Code_name,Scr_no,Field_code,Field_name,Code_type" & _
                               ") VALUES (N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "'," & _
                               RepSql(sPara4) & ",N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "')"

                    Case "MDY"
                        sSql = "UPDATE BDP210 SET " & _
                               "    Code_code=N'" & RepSql(sPara2) & "'," & _
                               "    Code_name=N'" & RepSql(sPara3) & "'," & _
                               "    Scr_no=" & RepSql(sPara4) & "," & _
                               "    Field_code=N'" & RepSql(sPara5) & "'," & _
                               "    Field_name=N'" & RepSql(sPara6) & "'," & _
                               "    Code_type=N'" & RepSql(sPara7) & "'" & _
                               " WHERE bdp210='" & RepSql(sPara1) & "'"
                End Select

            Case "BDP220A"
                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO BDP220(" & _
                               "prg_code,field_code,field_name,data_source,data_type,scr_no,is_use" & _
                               ") VALUES (N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "')"

                    Case "MDY"
                        sSql = "UPDATE BDP220 SET " & _
                               "    prg_code=N'" & RepSql(sPara2) & "'," & _
                               "    field_code=N'" & RepSql(sPara3) & "'," & _
                               "    field_name=N'" & RepSql(sPara4) & "'," & _
                               "    data_source=N'" & RepSql(sPara5) & "'," & _
                               "    data_type=N'" & RepSql(sPara6) & "'," & _
                               "    scr_no=N'" & RepSql(sPara7) & "'," & _
                               "    is_use=N'" & RepSql(sPara8) & "'" & _
                               " WHERE bdp220='" & RepSql(sPara1) & "'"
                End Select

            Case "BDP230A"
                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO BDP230(" & _
                               "dep_code,dep_name,dep_level,is_use" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "')"

                    Case "MDY"
                        sSql = "UPDATE BDP230 SET " & _
                               "    dep_name=N'" & RepSql(sPara2) & "'," & _
                               "    is_use=N'" & RepSql(sPara4) & "'," & _
                               "    dep_level=N'" & RepSql(sPara3) & "'" & _
                               " WHERE dep_code='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI001A"
                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO mail_asklaw(" & _
                               "ask_code,ask_name,ask_tel1,ask_email,ask_type,ask_memo,usr_code,ins_date" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "')"

                    Case "MDY"
                        sSql = "UPDATE mail_asklaw SET " & _
                               "    ask_name=N'" & RepSql(sPara2) & "'," & _
                               "    ask_tel1=N'" & RepSql(sPara3) & "'," & _
                               "    ask_email=N'" & RepSql(sPara4) & "'," & _
                               "    ask_type=N'" & RepSql(sPara5) & "'," & _
                               "    ask_memo=N'" & RepSql(sPara6) & "'" & _
                               " WHERE ask_code='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI002A"
                If String.IsNullOrEmpty(sPara2) Then sPara2 = "未填寫"
                If String.IsNullOrEmpty(sPara3) Then sPara3 = "未填寫"
                If String.IsNullOrEmpty(sPara4) Then sPara4 = "未填寫"
                If String.IsNullOrEmpty(sPara5) Then sPara5 = "未填寫"
                If String.IsNullOrEmpty(sPara6) Then sPara6 = "未填寫"
                If String.IsNullOrEmpty(sPara7) Then sPara7 = "未填寫"
                If String.IsNullOrEmpty(sPara8) Then sPara8 = "未填寫"
                If String.IsNullOrEmpty(sPara9) Then sPara9 = "未填寫"
                If String.IsNullOrEmpty(sPara32) Then sPara32 = "未填寫"
                If String.IsNullOrEmpty(sPara36) Then sPara36 = "未填寫"
                If String.IsNullOrEmpty(sPara40) Then sPara40 = "未填寫"
                Dim temp() As String = sPara50.Split("#") '是否為代他人申訴#申訴人部門#申訴人職稱

                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO mail_complain(" & _
                               "cmp_code,cmp_name,cmp_tel1,cmp_email,cmp_idno,com_name,com_add,com_idno,cmp_memo,cmp_type" & _
                               ",cmp_file1,cmp_file2,cmp_file3,cmp_file4,cmp_file5,usr_code,ins_date,att_area,att_road,is_agree" & _
                               ",cmp_date,cmp_idclass,cmp_sdate,cmp_jobcontent,cmp_edate,com_type,com_payarea,com_payroad,com_payadd" & _
                               ",com_Contract,com_Contract_chk,com_Contract_other,com_paymon,com_relax,com_relax_chk,com_relax_other" & _
                               ",com_child,com_Welfar,com_Welfare_chk,com_Welfare_other,com_disputetype,com_disputecon,com_Mediationtype" & _
                               ",com_Mediationcon,cmp_pass,com_schedule,com_sectype,com_sex,cmp_age,cmp_entrust,cmp_depart,cmp_job" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "',N'" & _
                               RepSql(sPara9) & "',N'" & _
                               RepSql(sPara10) & "',N'" & _
                               RepSql(sPara11) & "',N'" & _
                               RepSql(sPara12) & "',N'" & _
                               RepSql(sPara13) & "',N'" & _
                               RepSql(sPara14) & "',N'" & _
                               RepSql(sPara15) & "',N'" & _
                               RepSql(sPara16) & "',N'" & _
                               RepSql(sPara17) & "',N'" & _
                               RepSql(sPara18) & "',N'" & _
                               RepSql(sPara19) & "',N'" & _
                               RepSql(sPara20) & "',N'" & _
                               RepSql(sPara21) & "',N'" & _
                               RepSql(sPara22) & "',N'" & _
                               RepSql(sPara23) & "',N'" & _
                               RepSql(sPara24) & "',N'" & _
                               RepSql(sPara25) & "',N'" & _
                               RepSql(sPara26) & "',N'" & _
                               RepSql(sPara27) & "',N'" & _
                               RepSql(sPara28) & "',N'" & _
                               RepSql(sPara29) & "',N'" & _
                               RepSql(sPara30) & "',N'" & _
                               RepSql(sPara31) & "',N'" & _
                               RepSql(sPara32) & "',N'" & _
                               RepSql(sPara33) & "',N'" & _
                               RepSql(sPara34) & "',N'" & _
                               RepSql(sPara35) & "',N'" & _
                               RepSql(sPara36) & "',N'" & _
                               RepSql(sPara37) & "',N'" & _
                               RepSql(sPara38) & "',N'" & _
                               RepSql(sPara39) & "',N'" & _
                               RepSql(sPara40) & "',N'" & _
                               RepSql(sPara41) & "',N'" & _
                               RepSql(sPara42) & "',N'" & _
                               RepSql(sPara43) & "',N'" & _
                               RepSql(sPara44) & "',N'" & _
                               RepSql(sPara45) & "',N'" & _
                               RepSql(sPara46) & "',N'" & _
                               RepSql(sPara47) & "',N'" & _
                               RepSql(sPara48) & "',N'" & _
                               RepSql(sPara49) & "'," & _
                               RepSql(temp(0)) & ",N'" & _
                               RepSql(temp(1)) & "',N'" & _
                               RepSql(temp(2)) & "')"

                    Case "MDY"
                        sSql = "UPDATE mail_complain SET " & _
                               "    cmp_name=N'" & RepSql(sPara2) & "'," & _
                               "    cmp_tel1=N'" & RepSql(sPara3) & "'," & _
                               "    cmp_email=N'" & RepSql(sPara4) & "'," & _
                               "    cmp_idno=N'" & RepSql(sPara5) & "'," & _
                               "    com_name=N'" & RepSql(sPara6) & "'," & _
                               "    com_add=N'" & RepSql(sPara7) & "'," & _
                               "    com_idno=N'" & RepSql(sPara8) & "'," & _
                               "    cmp_memo=N'" & RepSql(sPara9) & "'," & _
                               "    cmp_type=N'" & RepSql(sPara10) & "'," & _
                               "    att_area=N'" & RepSql(sPara18) & "'," & _
                               "    att_road=N'" & RepSql(sPara19) & "'," & _
                               "    is_agree=N'" & RepSql(sPara20) & "'," & _
                               "    cmp_date=N'" & RepSql(sPara21) & "'," & _
                               "    cmp_idclass=N'" & RepSql(sPara22) & "'," & _
                               "    cmp_sdate=N'" & RepSql(sPara23) & "'," & _
                               "    cmp_jobcontent=N'" & RepSql(sPara24) & "'," & _
                               "    cmp_edate=N'" & RepSql(sPara25) & "'," & _
                               "    com_type=N'" & RepSql(sPara26) & "'," & _
                               "    com_payarea=N'" & RepSql(sPara27) & "'," & _
                               "    com_payroad=N'" & RepSql(sPara28) & "'," & _
                               "    com_payadd=N'" & RepSql(sPara29) & "'," & _
                               "    com_Contract=N'" & RepSql(sPara30) & "'," & _
                               "    com_Contract_chk=N'" & RepSql(sPara31) & "'," & _
                               "    com_Contract_other=N'" & RepSql(sPara32) & "'," & _
                               "    com_paymon=N'" & RepSql(sPara33) & "'," & _
                               "    com_relax=N'" & RepSql(sPara34) & "'," & _
                               "    com_relax_chk=N'" & RepSql(sPara35) & "'," & _
                               "    com_relax_other=N'" & RepSql(sPara36) & "'," & _
                               "    com_child=N'" & RepSql(sPara37) & "'," & _
                               "    com_Welfar=N'" & RepSql(sPara38) & "'," & _
                               "    com_Welfare_chk=N'" & RepSql(sPara39) & "'," & _
                               "    com_Welfare_other=N'" & RepSql(sPara40) & "'," & _
                               "    com_disputetype=N'" & RepSql(sPara41) & "'," & _
                               "    com_disputecon=N'" & RepSql(sPara42) & "'," & _
                               "    com_Mediationtype=N'" & RepSql(sPara43) & "'," & _
                               "    com_Mediationcon=N'" & RepSql(sPara44) & "'," & _
                               "    cmp_pass=N'" & RepSql(sPara45) & "'," & _
                               "    com_schedule=N'" & RepSql(sPara46) & "'," & _
                               "    com_sectype=N'" & RepSql(sPara47) & "'," & _
                               "    com_sex=N'" & RepSql(sPara48) & "'," & _
                               "    cmp_age=N'" & RepSql(sPara49) & "'," & _
                               "    cmp_entrust=" & RepSql(temp(0)) & "," & _
                               "    cmp_depart=N'" & RepSql(temp(1)) & "'," & _
                               "    cmp_job=N'" & RepSql(temp(2)) & "'" & _
                               " WHERE cmp_code='" & RepSql(sPara1) & "'"

                    Case "YComplain"
                        sSql = "UPDATE YComplain SET " & _
                               "      content=N'" & RepSql(sPara1) & "'"
                    Case "NComplain"
                        sSql = "UPDATE NComplain SET " & _
                               "      content=N'" & RepSql(sPara1) & "'"
                End Select

            Case "LSI022A"
                If String.IsNullOrEmpty(sPara2) Then sPara2 = "未填寫"
                If String.IsNullOrEmpty(sPara3) Then sPara3 = "未填寫"
                If String.IsNullOrEmpty(sPara4) Then sPara4 = "未填寫"
                If String.IsNullOrEmpty(sPara5) Then sPara5 = "未填寫"
                If String.IsNullOrEmpty(sPara6) Then sPara6 = "未填寫"
                If String.IsNullOrEmpty(sPara7) Then sPara7 = "未填寫"

                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO mail(" & _
                               "cmp_code,ins_date,cmp_date,cmp_memo,cmp_name,cmp_email,cmp_tel1" & _
                               ",cmp_file1,cmp_file2,cmp_file3,cmp_file4,cmp_file5" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "',N'" & _
                               RepSql(sPara9) & "',N'" & _
                               RepSql(sPara10) & "',N'" & _
                               RepSql(sPara11) & "',N'" & _
                               RepSql(sPara12) & "')"

                    Case "MDY"
                        sSql = "UPDATE mail SET " & _
                               "    ins_date=N'" & RepSql(sPara2) & "'," & _
                               "    cmp_date=N'" & RepSql(sPara3) & "'," & _
                               "    cmp_memo=N'" & RepSql(sPara4) & "'," & _
                               "    cmp_name=N'" & RepSql(sPara5) & "'," & _
                               "    cmp_email=N'" & RepSql(sPara6) & "'," & _
                               "    cmp_email=N'" & RepSql(sPara7) & "'," & _
                               "    cmp_email=N'" & RepSql(sPara8) & "'," & _
                               "    cmp_email=N'" & RepSql(sPara9) & "'," & _
                               "    cmp_email=N'" & RepSql(sPara10) & "'," & _
                               "    cmp_email=N'" & RepSql(sPara11) & "'," & _
                               "    cmp_tel1=N'" & RepSql(sPara12) & "'" & _
                               " WHERE cmp_code='" & RepSql(sPara1) & "'"

                    Case "Complain"
                        sSql = "UPDATE Complain SET " & _
                               "      content=N'" & RepSql(sPara1) & "'"
                End Select

            Case "LSI003A"
                If String.IsNullOrEmpty(sPara2) Then sPara2 = "未填寫"
                If String.IsNullOrEmpty(sPara5) Then sPara5 = "未填寫"
                If String.IsNullOrEmpty(sPara8) Then sPara8 = "未填寫"
                If String.IsNullOrEmpty(sPara9) Then sPara9 = "未填寫"
                If String.IsNullOrEmpty(sPara10) Then sPara10 = "未填寫"
                If String.IsNullOrEmpty(sPara11) Then sPara11 = "未填寫"

                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO e_paper_order(" & _
                               "ord_code,com_name,com_type1,com_type2,com_type2_memo," & _
                               "per_type,job_kind,job_kind_memo,per_tel1,per_mail, " & _
                               "pet_date,Is_order,usr_code,ins_date" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "',N'" & _
                               RepSql(sPara9) & "',N'" & _
                               RepSql(sPara10) & "',N'" & _
                               RepSql(sPara11) & "',N'" & _
                               RepSql(sPara12) & "',N'" & _
                               RepSql(sPara13) & "',N'" & _
                               RepSql(sPara14) & "')"

                    Case "MDY"
                        sSql = "UPDATE e_paper_order SET " & _
                               "    com_name=N'" & RepSql(sPara2) & "'," & _
                               "    com_type1=N'" & RepSql(sPara3) & "'," & _
                               "    com_type2=N'" & RepSql(sPara4) & "'," & _
                               "    com_type2_memo=N'" & RepSql(sPara5) & "'," & _
                               "    per_type=N'" & RepSql(sPara6) & "'," & _
                               "    job_kind=N'" & RepSql(sPara7) & "'," & _
                               "    job_kind_memo=N'" & RepSql(sPara8) & "'," & _
                               "    per_tel1=N'" & RepSql(sPara9) & "'," & _
                               "    per_mail=N'" & RepSql(sPara10) & "'," & _
                               "    pet_date=N'" & RepSql(sPara11) & "'," & _
                               "    Is_order=N'" & RepSql(sPara12) & "'" & _
                               " WHERE ord_code='" & RepSql(sPara1) & "'"

                    Case "RefundOrder"
                        sSql = "UPDATE e_paper_order SET " & _
                               "    Is_order=N'" & RepSql(sPara1) & "'" & _
                               " WHERE com_name=N'" & RepSql(sPara2) & "'" & _
                               " AND per_mail='" & RepSql(sPara3) & "'"
                End Select

            Case "yumi2" '20160912 yumi 新增
                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO yumi2(" & _
                               "book_code,book_title,book_author,book_translation,book_publisher," & _
                               "book_area,book_road,book_address,book_phone,book_print,book_premiums,pet_date" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "',N'" & _
                               RepSql(sPara9) & "',N'" & _
                               RepSql(sPara10) & "',N'" & _
                               RepSql(sPara11) & "',N'" & _
                               RepSql(sPara12) & "')"

                    Case "MDY"
                        sSql = "UPDATE yumi2 SET " & _
                               "    book_title=N'" & RepSql(sPara2) & "'," & _
                               "    book_author=N'" & RepSql(sPara3) & "'," & _
                               "    book_translation=N'" & RepSql(sPara4) & "'," & _
                               "    book_publisher=N'" & RepSql(sPara5) & "'," & _
                               "    book_area=N'" & RepSql(sPara6) & "'," & _
                               "    book_road=N'" & RepSql(sPara7) & "'," & _
                               "    book_address=N'" & RepSql(sPara8) & "'," & _
                               "    book_phone=N'" & RepSql(sPara9) & "'," & _
                               "    book_print=N'" & RepSql(sPara10) & "'," & _
                               "    book_premiums=N'" & RepSql(sPara11) & "'," & _
                               "    pet_date=N'" & RepSql(sPara12) & "'" & _
                               " WHERE book_code='" & RepSql(sPara1) & "'"

                        'Case "RefundOrder"
                        '    sSql = "UPDATE yumi2 SET " & _
                        '           "    Is_order=N'" & RepSql(sPara1) & "'" & _
                        '           " WHERE com_name=N'" & RepSql(sPara2) & "'" & _
                        '           " AND per_mail='" & RepSql(sPara3) & "'"
                End Select

            Case "yumi2-2" '20160909 yumi 新增
                If String.IsNullOrEmpty(sPara2) Then sPara2 = "未填寫"
                If String.IsNullOrEmpty(sPara5) Then sPara5 = "未填寫"
                If String.IsNullOrEmpty(sPara8) Then sPara8 = "未填寫"
                If String.IsNullOrEmpty(sPara9) Then sPara9 = "未填寫"
                If String.IsNullOrEmpty(sPara10) Then sPara10 = "未填寫"
                If String.IsNullOrEmpty(sPara11) Then sPara11 = "未填寫"

                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO yumi2_2(" & _
                               "ord_code,com_name,com_type1,com_type2,com_type2_memo," & _
                               "per_type,job_kind,job_kind_memo,per_tel1,per_mail, " & _
                               "pet_date,Is_order,usr_code,ins_date, com_sex " & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "',N'" & _
                               RepSql(sPara9) & "',N'" & _
                               RepSql(sPara10) & "',N'" & _
                               RepSql(sPara11) & "',N'" & _
                               RepSql(sPara12) & "',N'" & _
                               RepSql(sPara13) & "',N'" & _
                               RepSql(sPara14) & "',N'" & _
                               RepSql(sPara15) & "')"

                    Case "MDY"
                        sSql = "UPDATE yumi2_2 SET " & _
                               "    com_name=N'" & RepSql(sPara2) & "'," & _
                               "    com_type1=N'" & RepSql(sPara3) & "'," & _
                               "    com_type2=N'" & RepSql(sPara4) & "'," & _
                               "    com_type2_memo=N'" & RepSql(sPara5) & "'," & _
                               "    per_type=N'" & RepSql(sPara6) & "'," & _
                               "    job_kind=N'" & RepSql(sPara7) & "'," & _
                               "    job_kind_memo=N'" & RepSql(sPara8) & "'," & _
                               "    per_tel1=N'" & RepSql(sPara9) & "'," & _
                               "    per_mail=N'" & RepSql(sPara10) & "'," & _
                               "    pet_date=N'" & RepSql(sPara11) & "'," & _
                               "    Is_order=N'" & RepSql(sPara12) & "'," & _
                               "    com_sex=" & RepSql(sPara15) & "" & _
                               " WHERE ord_code='" & RepSql(sPara1) & "'"

                    Case "RefundOrder"
                        sSql = "UPDATE yumi2_2 SET " & _
                               "    Is_order=N'" & RepSql(sPara1) & "'" & _
                               " WHERE com_name=N'" & RepSql(sPara2) & "'" & _
                               " AND per_mail='" & RepSql(sPara3) & "'"
                End Select

            Case "Downtime" '20161004 yumi 新增 & 20161229 edit

                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO Downtime(" & _
                               "Down_institutions,Down_title,Down_address,Down_stopDate," & _
                               "Down_returnDate,Down_range,Down_reason,Down_remarks," & _
                               "Down_cDate,Down_creatorID,Down_X,Down_Y " & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "','" & _
                               RepSql(sPara4) & "',"

                        sSql &= IIf(sPara5 = "", "NULL,N'", "'" & RepSql(sPara5) & "',N'")

                        sSql &= RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "'," & _
                               "getdate()" & ",N'" & _
                               RepSql(sPara9) & "',N'" & _
                               RepSql(sPara11) & "',N'" & _
                               RepSql(sPara12) & "')"
                        'HttpContext.Current.Response.Write(sSql)
                        'HttpContext.Current.Response.End()

                    Case "MDY"
                        sSql = "UPDATE Downtime SET " & _
                               "    Down_X=N'" & RepSql(sPara11) & "'," & _
                               "    Down_Y=N'" & RepSql(sPara12) & "'," & _
                               "    Down_institutions=N'" & RepSql(sPara1) & "'," & _
                               "    Down_title=N'" & RepSql(sPara2) & "'," & _
                               "    Down_address=N'" & RepSql(sPara3) & "'," & _
                               "    Down_stopDate='" & RepSql(sPara4) & "'," & _
                               "    Down_range=N'" & RepSql(sPara6) & "'," & _
                               "    Down_reason=N'" & RepSql(sPara7) & "'," & _
                               "    Down_remarks=N'" & RepSql(sPara8) & "'," & _
                               "    Down_uDate=getdate()," & _
                               "    Down_modifyID=N'" & RepSql(sPara9) & "',"

                        sSql &= IIf(sPara5 = "", "    Down_returnDate=" & "NULL", "    Down_returnDate='" & RepSql(sPara5) & "'")

                        sSql &= " WHERE Down_code=" & RepSql(sPara10) & ""
                        'HttpContext.Current.Response.Write(sSql)
                        'HttpContext.Current.Response.End()

                    Case "DEL"
                        sSql = "DELETE FROM Downtime"

                        'Case "RefundOrder"
                        '    sSql = "UPDATE yumi2 SET " & _
                        '           "    Is_order=N'" & RepSql(sPara1) & "'" & _
                        '           " WHERE com_name=N'" & RepSql(sPara2) & "'" & _
                        '           " AND per_mail='" & RepSql(sPara3) & "'"
                End Select

            Case "Banner" '20161129 yumi 新增
                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO Banner(" & _
                               "title,imgName " & _
                               ") VALUES (N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "')"

                    Case "MDY"
                        sSql = "UPDATE Banner SET " & _
                               "    title=N'" & RepSql(sPara1) & "', " & _
                               "    imgName=N'" & RepSql(sPara2) & "' " & _
                               " WHERE bn_code=N'" & RepSql(sPara3) & "'"

                        'Case "RefundOrder"
                        '    sSql = "UPDATE yumi2 SET " & _
                        '           "    Is_order=N'" & RepSql(sPara1) & "'" & _
                        '           " WHERE com_name=N'" & RepSql(sPara2) & "'" & _
                        '           " AND per_mail='" & RepSql(sPara3) & "'"
                End Select

            Case "LSI004A"
                If String.IsNullOrEmpty(sPara2) Then sPara2 = "未填寫"
                If String.IsNullOrEmpty(sPara3) Then sPara3 = "未填寫"
                If String.IsNullOrEmpty(sPara5) Then sPara5 = "未填寫"
                If String.IsNullOrEmpty(sPara6) Then sPara6 = "未填寫"
                If String.IsNullOrEmpty(sPara10) Then sPara10 = "未填寫"

                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO e_paper(" & _
                               "paper_code,paper_name,paper_num,paper_logo,paper_bottom1,paper_bottom2,usr_code,ins_date,paper_type" & _
                               ",paper_html,paper_css" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               Replace(sPara5, "'", "''") & "',N'" & _
                               Replace(sPara6, "'", "''") & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "',N'" & _
                               RepSql(sPara9) & "',N'" & _
                               RepSql(sPara10) & "',N'" & _
                               RepSql(sPara11) & "')"

                    Case "MDY"
                        sSql = "UPDATE e_paper SET " & _
                               "    paper_name=N'" & RepSql(sPara2) & "'," & _
                               "    paper_num=N'" & RepSql(sPara3) & "'," & _
                               "    paper_logo=N'" & RepSql(sPara4) & "'," & _
                               "    paper_bottom1=N'" & Replace(sPara5, "'", "''") & "'," & _
                               "    paper_bottom2=N'" & Replace(sPara6, "'", "''") & "'," & _
                               "    usr_code=N'" & RepSql(sPara7) & "'," & _
                               "    paper_type=N'" & RepSql(sPara9) & "'," & _
                               "    paper_css=N'" & RepSql(sPara11) & "'" & _
                               " WHERE paper_code='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI004A_detail"
                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO e_paper_d(" & _
                               "paper_code,news_code,news_type,scr_no" & _
                               ") VALUES (N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "')"

                    Case "MDY"
                        sSql = "UPDATE e_paper_d SET " & _
                               "    paper_code=N'" & RepSql(sPara2) & "'," & _
                               "    news_code=N'" & RepSql(sPara3) & "'," & _
                               "    news_type=N'" & RepSql(sPara4) & "'," & _
                               "    scr_no=N'" & RepSql(sPara5) & "'" & _
                               " WHERE serial_no='" & RepSql(sPara1) & "'"

                    Case "PaperHtml"
                        sSql = "UPDATE e_paper SET " & _
                               "    paper_html=N'" & Replace(sPara2, "'", "''") & "'" & _
                               " WHERE paper_code='" & RepSql(sPara1) & "'"

                    Case "NewsContect"
                        sSql = "UPDATE E_NEWS SET " & _
                               "    news_contect=N'" & Replace(sPara2, "'", "''") & "'" & _
                               " WHERE news_code='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI005A"
                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO e_paper_sch(" & _
                               "sch_code,paper_code,send_date,send_time,send_state" & _
                               ",usr_code,ins_date,last_date,is_show,count" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "',N'" & _
                               RepSql(sPara9) & "',0)"

                    Case "MDY"
                        sSql = "UPDATE e_paper_sch SET " & _
                               "    paper_code=N'" & RepSql(sPara2) & "'," & _
                               "    send_date=N'" & RepSql(sPara3) & "'," & _
                               "    send_time=N'" & RepSql(sPara4) & "'," & _
                               "    send_state=N'" & RepSql(sPara5) & "'" & _
                               " WHERE sch_code='" & RepSql(sPara1) & "'"

                    Case "Result"
                        sSql = "UPDATE e_paper_sch SET " & _
                               "    last_date=N'" & RepSql(sPara2) & "'," & _
                               "    count=count+1" & _
                               " WHERE sch_code='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI005A_detail"
                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO e_paper_log(" & _
                               "sch_code,mail_list,cmemo" & _
                               ") VALUES (N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "')"

                    Case "MDY"
                        sSql = "UPDATE e_paper_log SET " & _
                               "    sch_code=N'" & RepSql(sPara2) & "'," & _
                               "    mail_list=N'" & RepSql(sPara3) & "'," & _
                               "    cmemo=N'" & RepSql(sPara4) & "'" & _
                               " WHERE serial_no='" & RepSql(sPara1) & "'"

                    Case "PaperHtml"
                        sSql = "UPDATE e_paper SET " & _
                               "    paper_html=N'" & Replace(sPara2, "'", "''") & "'" & _
                               " WHERE paper_code='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI006A"
                sSql = " UPDATE BDP000 SET par_value='" & RepSql(sPara1) & "' WHERE par_name ='sender';"
                sSql &= " UPDATE BDP000 SET par_value='" & RepSql(sPara2) & "' WHERE par_name ='sender_mail';"
                sSql &= " UPDATE BDP000 SET par_value='" & RepSql(sPara3) & "' WHERE par_name ='resender_mail'"

            Case "LSI007A"
                If String.IsNullOrEmpty(sPara2) Then sPara2 = "未填寫"
                If String.IsNullOrEmpty(sPara4) Then sPara4 = "未填寫"

                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO Meeting_room(" & _
                               "room_code,room_name,is_use,room_memo,usr_code,ins_date" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "')"

                    Case "MDY"
                        sSql = "UPDATE Meeting_room SET " & _
                               "    room_name=N'" & RepSql(sPara2) & "'," & _
                               "    is_use=N'" & RepSql(sPara3) & "'," & _
                               "    room_memo=N'" & RepSql(sPara4) & "'" & _
                               " WHERE room_code='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI008A"
                If String.IsNullOrEmpty(sPara7) Then sPara7 = "未填寫"
                If String.IsNullOrEmpty(sPara8) Then sPara8 = "未填寫"
                If String.IsNullOrEmpty(sPara9) Then sPara9 = "未填寫"

                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO Meeting_order(" & _
                               "mrd_code,room_code,mrd_date,mrd_time_s,mrd_time_e,usr_code,mrd_name,mrd_host,mrd_memo,send_list,is_show" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "',N'" & _
                               RepSql(sPara9) & "',N'" & _
                               RepSql(sPara10) & "',N'" & _
                               RepSql(sPara11) & "')"

                    Case "MDY"
                        sSql = "UPDATE Meeting_order SET " & _
                               "    room_code=N'" & RepSql(sPara2) & "'," & _
                               "    mrd_date=N'" & RepSql(sPara3) & "'," & _
                               "    mrd_time_s=N'" & RepSql(sPara4) & "'," & _
                               "    mrd_time_e=N'" & RepSql(sPara5) & "'," & _
                               "    usr_code=N'" & RepSql(sPara6) & "'," & _
                               "    mrd_name=N'" & RepSql(sPara7) & "'," & _
                               "    mrd_host=N'" & RepSql(sPara8) & "'," & _
                               "    mrd_memo=N'" & RepSql(sPara9) & "'," & _
                               "    send_list=N'" & RepSql(sPara10) & "'," & _
                               "    is_show=N'" & RepSql(sPara11) & "'" & _
                               " WHERE mrd_code='" & RepSql(sPara1) & "'"
                End Select

            Case "LSI008A_detail"
                If String.IsNullOrEmpty(sPara3) Then sPara3 = "未填寫"
                If String.IsNullOrEmpty(sPara4) Then sPara4 = "未填寫"
                If String.IsNullOrEmpty(sPara5) Then sPara5 = "未填寫"
                If String.IsNullOrEmpty(sPara6) Then sPara6 = "未填寫"

                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO Meeting_users(" & _
                               "mrd_code,usr_code,usr_name,usr_mail,usr_dep" & _
                               ") VALUES (N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "')"

                    Case "MDY"
                        sSql = "UPDATE Meeting_users SET " & _
                               "    usr_code=N'" & RepSql(sPara3) & "'," & _
                               "    usr_name=N'" & RepSql(sPara4) & "'," & _
                               "    usr_mail=N'" & RepSql(sPara5) & "'," & _
                               "    usr_dep=N'" & RepSql(sPara6) & "'" & _
                               " WHERE serial_no=N'" & RepSql(sPara1) & "'"
                End Select

            Case "LSI009A"
                If String.IsNullOrEmpty(sPara6) Then sPara6 = "未填寫"
                If String.IsNullOrEmpty(sPara8) Then sPara8 = "未填寫"
                If String.IsNullOrEmpty(sPara9) Then sPara9 = "未填寫"
                If String.IsNullOrEmpty(sPara10) Then sPara10 = "未填寫"
                If String.IsNullOrEmpty(sPara11) Then sPara11 = "未填寫"
                If String.IsNullOrEmpty(sPara12) Then sPara12 = "未填寫"
                If String.IsNullOrEmpty(sPara13) Then sPara13 = "未填寫"

                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO Work_stop(" & _
                               "stop_code,stop_number,usr_code,ins_date,stop_date,stop_time,stop_date_e,stop_time_e,com_name" & _
                               " ,com_per,eng_name,eng_code,eng_gps,chk_date,ver_code" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "',N'" & _
                               RepSql(sPara9) & "',N'" & _
                               RepSql(sPara10) & "',N'" & _
                               RepSql(sPara11) & "',N'" & _
                               RepSql(sPara12) & "',N'" & _
                               RepSql(sPara13) & "',N'" & _
                               RepSql(sPara14) & "',N'" & _
                               RepSql(sPara15) & "')"

                    Case "MDY"
                        sSql = "UPDATE Work_stop SET " & _
                               "    stop_number=N'" & RepSql(sPara2) & "'," & _
                               "    usr_code=N'" & RepSql(sPara3) & "'," & _
                               "    ins_date=N'" & RepSql(sPara4) & "'," & _
                               "    stop_date=N'" & RepSql(sPara5) & "'," & _
                               "    stop_time=N'" & RepSql(sPara6) & "'," & _
                               "    stop_date_e=N'" & RepSql(sPara7) & "'," & _
                               "    stop_time_e=N'" & RepSql(sPara8) & "'," & _
                               "    com_name=N'" & RepSql(sPara9) & "'," & _
                               "    com_per=N'" & RepSql(sPara10) & "'," & _
                               "    eng_name=N'" & RepSql(sPara11) & "'," & _
                               "    eng_code=N'" & RepSql(sPara12) & "'," & _
                               "    eng_gps=N'" & RepSql(sPara13) & "'," & _
                               "    chk_date=N'" & RepSql(sPara14) & "'," & _
                               "    ver_code=N'" & RepSql(sPara15) & "'" & _
                               " WHERE stop_code=N'" & RepSql(sPara1) & "'"
                End Select

            Case "LSI010A"
                If String.IsNullOrEmpty(sPara3) Then sPara3 = "未填寫"
                If String.IsNullOrEmpty(sPara4) Then sPara4 = "未填寫"
                If String.IsNullOrEmpty(sPara6) Then sPara6 = "未填寫"
                If String.IsNullOrEmpty(sPara7) Then sPara7 = "未填寫"
                If String.IsNullOrEmpty(sPara8) Then sPara8 = "未填寫"
                If String.IsNullOrEmpty(sPara9) Then sPara9 = "未填寫"
                If String.IsNullOrEmpty(sPara13) Then sPara13 = "未填寫"
                If String.IsNullOrEmpty(sPara14) Then sPara14 = "未填寫"
                If String.IsNullOrEmpty(sPara15) Then sPara15 = "未填寫"
                If String.IsNullOrEmpty(sPara16) Then sPara16 = "未填寫"
                If String.IsNullOrEmpty(sPara20) Then sPara20 = "未填寫"

                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO Work_start(" & _
                               "start_code,stop_number,com_name,com_per,start_date,start_number,eng_name,eng_code,eng_gps,chk_date" & _
                               ",ins_date,stop_date,stop_time,att_name,att_tel1,att_email,start_path1,start_path2,stop_date_e,stop_time_e,usr_code" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "',N'" & _
                               RepSql(sPara9) & "',N'" & _
                               RepSql(sPara10) & "',N'" & _
                               RepSql(sPara11) & "',N'" & _
                               RepSql(sPara12) & "',N'" & _
                               RepSql(sPara13) & "',N'" & _
                               RepSql(sPara14) & "',N'" & _
                               RepSql(sPara15) & "',N'" & _
                               RepSql(sPara16) & "',N'" & _
                               RepSql(sPara17) & "',N'" & _
                               RepSql(sPara18) & "',N'" & _
                               RepSql(sPara19) & "',N'" & _
                               RepSql(sPara20) & "',N'" & _
                               RepSql(sPara21) & "')"

                    Case "MDY"
                        sSql = "UPDATE Work_start SET " & _
                               "    stop_number=N'" & RepSql(sPara2) & "'," & _
                               "    com_name=N'" & RepSql(sPara3) & "'," & _
                               "    com_per=N'" & RepSql(sPara4) & "'," & _
                               "    start_date=N'" & RepSql(sPara5) & "'," & _
                               "    start_number=N'" & RepSql(sPara6) & "'," & _
                               "    eng_name=N'" & RepSql(sPara7) & "'," & _
                               "    eng_code=N'" & RepSql(sPara8) & "'," & _
                               "    eng_gps=N'" & RepSql(sPara9) & "'," & _
                               "    chk_date=N'" & RepSql(sPara10) & "'," & _
                               "    ins_date=N'" & RepSql(sPara11) & "'," & _
                               "    stop_date=N'" & RepSql(sPara12) & "'," & _
                               "    stop_time=N'" & RepSql(sPara13) & "'," & _
                               "    att_name=N'" & RepSql(sPara14) & "'," & _
                               "    att_tel1=N'" & RepSql(sPara15) & "'," & _
                               "    att_email=N'" & RepSql(sPara16) & "'," & _
                               "    stop_date_e=N'" & RepSql(sPara19) & "'," & _
                               "    stop_time_e=N'" & RepSql(sPara20) & "'" & _
                               " WHERE start_code=N'" & RepSql(sPara1) & "'"
                End Select

            Case "LSI011A"
                If String.IsNullOrEmpty(sPara2) Then sPara2 = "未填寫"
                If String.IsNullOrEmpty(sPara3) Then sPara3 = "未填寫"
                If String.IsNullOrEmpty(sPara4) Then sPara4 = "未填寫"
                If String.IsNullOrEmpty(sPara5) Then sPara5 = "未填寫"
                If String.IsNullOrEmpty(sPara6) Then sPara6 = "未填寫"
                If String.IsNullOrEmpty(sPara8) Then sPara8 = "未填寫"
                If String.IsNullOrEmpty(sPara9) Then sPara9 = "未填寫"

                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO Self_check(" & _
                               "self_code,com_name,eng_name,eng_add,com_tel1,com_tel2,self_date,att_name,att_dep,eng_area" & _
                               ",start_path1,usr_code,ins_date,eng_road" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "',N'" & _
                               RepSql(sPara9) & "',N'" & _
                               RepSql(sPara10) & "',N'" & _
                               RepSql(sPara11) & "',N'" & _
                               RepSql(sPara12) & "',N'" & _
                               RepSql(sPara13) & "',N'" & _
                               RepSql(sPara14) & "')"

                    Case "MDY"
                        sSql = "UPDATE Self_check SET " & _
                               "    com_name=N'" & RepSql(sPara2) & "'," & _
                               "    eng_name=N'" & RepSql(sPara3) & "'," & _
                               "    eng_add=N'" & RepSql(sPara4) & "'," & _
                               "    com_tel1=N'" & RepSql(sPara5) & "'," & _
                               "    com_tel2=N'" & RepSql(sPara6) & "'," & _
                               "    self_date=N'" & RepSql(sPara7) & "'," & _
                               "    att_name=N'" & RepSql(sPara8) & "'," & _
                               "    att_dep=N'" & RepSql(sPara9) & "'," & _
                               "    eng_area=N'" & RepSql(sPara10) & "'" & _
                               " WHERE self_code=N'" & RepSql(sPara1) & "'"
                End Select

            Case "LSI012A"
                If String.IsNullOrEmpty(sPara2) Then sPara2 = "未填寫"
                If String.IsNullOrEmpty(sPara4) Then sPara4 = "未填寫"
                If String.IsNullOrEmpty(sPara5) Then sPara5 = "未填寫"
                If String.IsNullOrEmpty(sPara6) Then sPara6 = "未填寫"
                If String.IsNullOrEmpty(sPara7) Then sPara7 = "未填寫"
                If String.IsNullOrEmpty(sPara8) Then sPara8 = "未填寫"
                If String.IsNullOrEmpty(sPara9) Then sPara9 = "未填寫"
                If String.IsNullOrEmpty(sPara10) Then sPara10 = "未填寫"
                If String.IsNullOrEmpty(sPara11) Then sPara11 = "未填寫"
                If String.IsNullOrEmpty(sPara12) Then sPara12 = "未填寫"
                If String.IsNullOrEmpty(sPara13) Then sPara13 = "未填寫"
                If String.IsNullOrEmpty(sPara14) Then sPara14 = "未填寫"
                If String.IsNullOrEmpty(sPara15) Then sPara15 = "未填寫"
                If String.IsNullOrEmpty(sPara16) Then sPara16 = "未填寫"
                If String.IsNullOrEmpty(sPara17) Then sPara17 = "未填寫"
                If String.IsNullOrEmpty(sPara18) Then sPara18 = "未填寫"
                If String.IsNullOrEmpty(sPara19) Then sPara19 = "未填寫"
                If String.IsNullOrEmpty(sPara21) Then sPara21 = "未填寫"

                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO Small_space(" & _
                               "sma_code,com_name,eng_area,eng_add,eng_date_s,eng_date_e,pet_name,pet_date,pet_tel1,pet_mail " & _
                               ",att_name1,att_add1,att_tel1,att_name2,att_add2,att_tel2,att_name3,att_add3,att_tel3,pet_list" & _
                               ",pet_other, usr_code,ins_date,eng_road" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "',N'" & _
                               RepSql(sPara9) & "',N'" & _
                               RepSql(sPara10) & "',N'" & _
                               RepSql(sPara11) & "',N'" & _
                               RepSql(sPara12) & "',N'" & _
                               RepSql(sPara13) & "',N'" & _
                               RepSql(sPara14) & "',N'" & _
                               RepSql(sPara15) & "',N'" & _
                               RepSql(sPara16) & "',N'" & _
                               RepSql(sPara17) & "',N'" & _
                               RepSql(sPara18) & "',N'" & _
                               RepSql(sPara19) & "',N'" & _
                               RepSql(sPara20) & "',N'" & _
                               RepSql(sPara21) & "',N'" & _
                               RepSql(sPara22) & "',N'" & _
                               RepSql(sPara23) & "',N'" & _
                               RepSql(sPara24) & "')"

                    Case "MDY"
                        sSql = "UPDATE Small_space SET " & _
                               "    com_name=N'" & RepSql(sPara2) & "'," & _
                               "    eng_area=N'" & RepSql(sPara3) & "'," & _
                               "    eng_add=N'" & RepSql(sPara4) & "'," & _
                               "    eng_date_s=N'" & RepSql(sPara5) & "'," & _
                               "    eng_date_e=N'" & RepSql(sPara6) & "'," & _
                               "    pet_name=N'" & RepSql(sPara7) & "'," & _
                               "    pet_date=N'" & RepSql(sPara8) & "'," & _
                               "    pet_tel1=N'" & RepSql(sPara9) & "'," & _
                               "    pet_mail=N'" & RepSql(sPara10) & "'," & _
                               "    att_name1=N'" & RepSql(sPara11) & "'," & _
                               "    att_add1=N'" & RepSql(sPara12) & "'," & _
                               "    att_tel1=N'" & RepSql(sPara13) & "'," & _
                               "    att_name2=N'" & RepSql(sPara14) & "'," & _
                               "    att_add2=N'" & RepSql(sPara15) & "'," & _
                               "    att_tel2=N'" & RepSql(sPara16) & "'," & _
                               "    att_name3=N'" & RepSql(sPara17) & "'," & _
                               "    att_add3=N'" & RepSql(sPara18) & "'," & _
                               "    att_tel3=N'" & RepSql(sPara19) & "'," & _
                               "    pet_list=N'" & RepSql(sPara20) & "'," & _
                               "    pet_other=N'" & RepSql(sPara21) & "'," & _
                               "    eng_road=N'" & RepSql(sPara24) & "'" & _
                               " WHERE sma_code=N'" & RepSql(sPara1) & "'"
                End Select

            Case "LSI013A" '2014調整Allen08
                'If String.IsNullOrEmpty(sPara2) Then sPara2 = "未填寫"
                'If String.IsNullOrEmpty(sPara3) Then sPara3 = "未填寫"
                'If String.IsNullOrEmpty(sPara4) Then sPara4 = "未填寫"
                'If String.IsNullOrEmpty(sPara5) Then sPara5 = "未填寫"
                'If String.IsNullOrEmpty(sPara6) Then sPara6 = "未填寫"
                'If String.IsNullOrEmpty(sPara9) Then sPara9 = "未填寫"
                'If String.IsNullOrEmpty(sPara11) Then sPara11 = "未填寫"
                'If String.IsNullOrEmpty(sPara12) Then sPara12 = "未填寫"
                If String.IsNullOrEmpty(sPara13) Then sPara13 = "未填寫"
                If String.IsNullOrEmpty(sPara14) Then sPara14 = "未填寫"
                If String.IsNullOrEmpty(sPara15) Then sPara15 = "未填寫"
                'If String.IsNullOrEmpty(sPara16) Then sPara16 = "未填寫"
                'If String.IsNullOrEmpty(sPara17) Then sPara17 = "未填寫"
                If String.IsNullOrEmpty(sPara18) Then sPara18 = "未填寫"
                'If String.IsNullOrEmpty(sPara22) Then sPara22 = "未填寫"
                If String.IsNullOrEmpty(sPara23) Then sPara23 = "未填寫"
                If String.IsNullOrEmpty(sPara24) Then sPara24 = "未填寫"
                If String.IsNullOrEmpty(sPara25) Then sPara25 = "未填寫"
                If String.IsNullOrEmpty(sPara26) Then sPara26 = "未填寫"
                If String.IsNullOrEmpty(sPara27) Then sPara27 = "未填寫"
                If String.IsNullOrEmpty(sPara29) Then sPara29 = "未填寫"
                If String.IsNullOrEmpty(sPara30) Then sPara30 = "未填寫"
                If String.IsNullOrEmpty(sPara33) Then sPara33 = "未填寫"
                If String.IsNullOrEmpty(sPara34) Then sPara34 = ""
                If String.IsNullOrEmpty(sPara38) Then sPara38 = "未填寫"
                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO Fix_work(" & _
                               "fix_code,com_name,pet_tel1,pet_date,pet_time,att_add1,att_area,att_road,work_floor,work_floor_type" & _
                               ",work_date_s,work_date_e,att_cname1,att_name1,att_tel1,att_cname2,att_name2,att_tel2,fix_type, usr_code" & _
                               ",ins_date,safe_item,dan_code,att_cname3,att_name3,att_tel3,att_mob3,att_mob2,totheigh,othertype,daytype,holtype" & _
                               ",att_tel21,engtype,engchoosetype,Workplacetype,pet_man,att_common,att_email,att_chk,att_chkoth,att_man,att_tel" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "',N'" & _
                               RepSql(sPara9) & "',N'" & _
                               RepSql(sPara10) & "',N'" & _
                               RepSql(sPara11) & "',N'" & _
                               RepSql(sPara12) & "',N'" & _
                               RepSql(sPara13) & "',N'" & _
                               RepSql(sPara14) & "',N'" & _
                               RepSql(sPara15) & "',N'" & _
                               RepSql(sPara16) & "',N'" & _
                               RepSql(sPara17) & "',N'" & _
                               RepSql(sPara18) & "',N'" & _
                               RepSql(sPara19) & "',N'" & _
                               RepSql(sPara20) & "',N'" & _
                               RepSql(sPara21) & "',N'" & _
                               RepSql(sPara22) & "',N'" & _
                               RepSql(sPara23) & "',N'" & _
                               RepSql(sPara24) & "',N'" & _
                               RepSql(sPara25) & "',N'" & _
                               RepSql(sPara26) & "',N'" & _
                               RepSql(sPara27) & "',N'" & _
                               RepSql(sPara28) & "',N'" & _
                               RepSql(sPara29) & "',N'" & _
                               RepSql(sPara30) & "',N'" & _
                               RepSql(sPara31) & "',N'" & _
                               RepSql(sPara32) & "',N'" & _
                               RepSql(sPara33) & "',N'" & _
                               RepSql(sPara34) & "',N'" & _
                               RepSql(sPara35) & "',N'" & _
                               RepSql(sPara36) & "',N'" & _
                               RepSql(sPara37) & "',N'" & _
                               RepSql(sPara38) & "',N'" & _
                               RepSql(sPara39) & "',N'" & _
                               RepSql(sPara40) & "',N'" & _
                               RepSql(sPara41) & "',N'" & _
                               RepSql(sPara42) & "',N'" & _
                               RepSql(sPara43) & "')"

                    Case "MDY"
                        sSql = "UPDATE Fix_work SET " & _
                               "    com_name=N'" & RepSql(sPara2) & "'," & _
                               "    pet_tel1=N'" & RepSql(sPara3) & "'," & _
                               "    pet_date=N'" & RepSql(sPara4) & "'," & _
                               "    pet_time=N'" & RepSql(sPara5) & "'," & _
                               "    att_add1=N'" & RepSql(sPara6) & "'," & _
                               "    att_area=N'" & RepSql(sPara7) & "'," & _
                               "    att_road=N'" & RepSql(sPara8) & "'," & _
                               "    work_floor=N'" & RepSql(sPara9) & "'," & _
                               "    work_floor_type=N'" & RepSql(sPara10) & "'," & _
                               "    work_date_s=N'" & RepSql(sPara11) & "'," & _
                               "    work_date_e=N'" & RepSql(sPara12) & "'," & _
                               "    att_cname1=N'" & RepSql(sPara13) & "'," & _
                               "    att_name1=N'" & RepSql(sPara14) & "'," & _
                               "    att_tel1=N'" & RepSql(sPara15) & "'," & _
                               "    att_cname2=N'" & RepSql(sPara16) & "'," & _
                               "    att_name2=N'" & RepSql(sPara17) & "'," & _
                               "    att_tel2=N'" & RepSql(sPara18) & "'," & _
                               "    fix_type=N'" & RepSql(sPara19) & "'," & _
                               "    safe_item=N'" & RepSql(sPara22) & "'," & _
                               "    dan_code=N'" & RepSql(sPara23) & "'," & _
                               "    att_cname3=N'" & RepSql(sPara24) & "'," & _
                               "    att_name3=N'" & RepSql(sPara25) & "'," & _
                               "    att_tel3=N'" & RepSql(sPara26) & "'," & _
                               "    att_mob3=N'" & RepSql(sPara27) & "'," & _
                               "    att_mob2=N'" & RepSql(sPara28) & "'," & _
                               "    totheigh=N'" & RepSql(sPara29) & "'," & _
                               "    othertype=N'" & RepSql(sPara30) & "'," & _
                               "    daytype=N'" & RepSql(sPara31) & "'," & _
                               "    holtype=N'" & RepSql(sPara32) & "'," & _
                               "    att_tel21=N'" & RepSql(sPara33) & "'," & _
                               "    engtype=N'" & RepSql(sPara34) & "'," & _
                               "    engchoosetype=N'" & RepSql(sPara35) & "'," & _
                               "    Workplacetype=N'" & RepSql(sPara36) & "'," & _
                               "    pet_man=N'" & RepSql(sPara37) & "'," & _
                               "    att_common=N'" & RepSql(sPara38) & "'," & _
                               "    att_email=N'" & RepSql(sPara39) & "'," & _
                               "    att_chk=N'" & RepSql(sPara40) & "'," & _
                               "    att_chkoth=N'" & RepSql(sPara41) & "'," & _
                               "    att_man=N'" & RepSql(sPara42) & "'," & _
                               "    att_tel=N'" & RepSql(sPara43) & "'" & _
                               " WHERE fix_code=N'" & RepSql(sPara1) & "'"

                    Case "DanCode"
                        sSql = "UPDATE Fix_work SET " & _
                               "    dan_code=N'" & RepSql(sPara2) & "'" & _
                               " WHERE fix_code=N'" & RepSql(sPara1) & "'"

                    Case "guidance1"
                        sSql = "UPDATE guidance1 SET " & _
                               "      content=N'" & RepSql(sPara1) & "'"
                End Select
            Case "LSI020A" 'may2013/10/15

                If String.IsNullOrEmpty(sPara2) Then sPara2 = "未填寫"
                If String.IsNullOrEmpty(sPara3) Then sPara3 = "未填寫"
                If String.IsNullOrEmpty(sPara4) Then sPara4 = "未填寫"
                If String.IsNullOrEmpty(sPara5) Then sPara5 = "未填寫"
                If String.IsNullOrEmpty(sPara6) Then sPara6 = "未填寫"
                If String.IsNullOrEmpty(sPara9) Then sPara9 = "未填寫"
                If String.IsNullOrEmpty(sPara11) Then sPara11 = "未填寫"
                If String.IsNullOrEmpty(sPara12) Then sPara12 = "未填寫"
                If String.IsNullOrEmpty(sPara13) Then sPara13 = "未填寫"
                If String.IsNullOrEmpty(sPara14) Then sPara14 = "未填寫"
                If String.IsNullOrEmpty(sPara15) Then sPara15 = "未填寫"
                If String.IsNullOrEmpty(sPara17) Then sPara17 = "未填寫"
                If String.IsNullOrEmpty(sPara18) Then sPara18 = "未填寫"
                If String.IsNullOrEmpty(sPara19) Then sPara19 = "未填寫"


                Select Case sType
                    Case "ADD"
                        sSql = "INSERT INTO SCHOOL_WORK(" & _
                               "fix_code,com_name,pet_date,pet_time,work_floor,work_floor_type" & _
                               ",work_date_s,work_date_e,att_cname1,att_name1,att_tel1,att_cname2,att_name2,att_tel2,fix_type" & _
                               ",ins_date,safe_item,work_money,att_area,usr_code,att_mail" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "',N'" & _
                               RepSql(sPara9) & "',N'" & _
                               RepSql(sPara10) & "',N'" & _
                               RepSql(sPara11) & "',N'" & _
                               RepSql(sPara12) & "',N'" & _
                               RepSql(sPara13) & "',N'" & _
                               RepSql(sPara14) & "',N'" & _
                               RepSql(sPara15) & "',N'" & _
                               RepSql(sPara16) & "',N'" & _
                               RepSql(sPara17) & "',N'" & _
                               RepSql(sPara18) & "',N'" & _
                               RepSql(sPara19) & "',N'" & _
                               RepSql(sPara20) & "',N'" & _
                               RepSql(sPara21) & "')"
                    Case "MDY"
                        sSql = "UPDATE SCHOOL_WORK SET " & _
                               "    com_name=N'" & RepSql(sPara2) & "'," & _
                               "    work_floor=N'" & RepSql(sPara5) & "'," & _
                               "    work_floor_type=N'" & RepSql(sPara6) & "'," & _
                               "    work_date_s=N'" & RepSql(sPara7) & "'," & _
                               "    work_date_e=N'" & RepSql(sPara8) & "'," & _
                               "    att_cname1=N'" & RepSql(sPara9) & "'," & _
                               "    att_name1=N'" & RepSql(sPara10) & "'," & _
                               "    att_tel1=N'" & RepSql(sPara11) & "'," & _
                               "    att_cname2=N'" & RepSql(sPara12) & "'," & _
                               "    att_name2=N'" & RepSql(sPara13) & "'," & _
                               "    att_tel2=N'" & RepSql(sPara14) & "'," & _
                               "    fix_type=N'" & RepSql(sPara15) & "'," & _
                               "    safe_item=N'" & RepSql(sPara17) & "'," & _
                               "    work_money=N'" & RepSql(sPara18) & "'," & _
                               "    att_area=N'" & RepSql(sPara19) & "'," & _
                               "    att_mail=N'" & RepSql(sPara21) & "'" & _
                               " WHERE fix_code=N'" & RepSql(sPara1) & "'"
                End Select
            Case "LSI014A"
                If String.IsNullOrEmpty(sPara2) Then sPara2 = "未填寫"
                If String.IsNullOrEmpty(sPara3) Then sPara3 = "未填寫"
                If String.IsNullOrEmpty(sPara4) Then sPara4 = "未填寫"
                If String.IsNullOrEmpty(sPara5) Then sPara5 = "未填寫"
                If String.IsNullOrEmpty(sPara6) Then sPara6 = "未填寫"
                If String.IsNullOrEmpty(sPara7) Then sPara7 = "未填寫"
                If String.IsNullOrEmpty(sPara8) Then sPara8 = "未填寫"
                If String.IsNullOrEmpty(sPara9) Then sPara9 = "未填寫"
                If String.IsNullOrEmpty(sPara10) Then sPara10 = "未填寫"
                If String.IsNullOrEmpty(sPara11) Then sPara11 = "未填寫"
                If String.IsNullOrEmpty(sPara14) Then sPara14 = "未填寫"
                If String.IsNullOrEmpty(sPara15) Then sPara15 = "未填寫"
                If String.IsNullOrEmpty(sPara16) Then sPara16 = "未填寫"
                If String.IsNullOrEmpty(sPara17) Then sPara17 = "未填寫"
                If String.IsNullOrEmpty(sPara19) Then sPara19 = "未填寫"
                If String.IsNullOrEmpty(sPara20) Then sPara20 = "未填寫"
                If String.IsNullOrEmpty(sPara30) Then sPara30 = "未填寫"
                If String.IsNullOrEmpty(sPara36) Then sPara36 = "未填寫"
                If String.IsNullOrEmpty(sPara37) Then sPara37 = "未填寫"
                If String.IsNullOrEmpty(sPara39) Then sPara39 = "未填寫"
                If String.IsNullOrEmpty(sPara41) Then sPara41 = "未填寫"
                If String.IsNullOrEmpty(sPara42) Then sPara42 = "未填寫"
                If String.IsNullOrEmpty(sPara43) Then sPara43 = "未填寫"
                If String.IsNullOrEmpty(sPara46) Then sPara46 = "未填寫"

                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO Danger_work(" & _
                               "dan_code,com_name,pet_name,pet_date,pet_time,pet_tel1,att_name1,att_tel1,att_name2,att_tel2,att_add1 " & _
                               ",att_area,att_road,bulid_name,bulid_code,work_date_s,work_date_e,is_night,work_floor_s,work_floor_e " & _
                               ",work_type1,work_type2,work_type3,work_type4,work_type5,work_type6,work_type7,work_type8,work_type9 " & _
                               ",work_type_memo,use_tool1,use_tool2,use_tool3,use_tool4,use_tool5,use_tool3_memo,use_tool5_memo, " & _
                               "save_tool,save_tool_memo,dan_memo,dan_memo_memo,dan_mail,dan_mail2,usr_code,ins_date,fix_code,is_holiday" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "',N'" & _
                               RepSql(sPara9) & "',N'" & _
                               RepSql(sPara10) & "',N'" & _
                               RepSql(sPara11) & "',N'" & _
                               RepSql(sPara12) & "',N'" & _
                               RepSql(sPara13) & "',N'" & _
                               RepSql(sPara14) & "',N'" & _
                               RepSql(sPara15) & "',N'" & _
                               RepSql(sPara16) & "',N'" & _
                               RepSql(sPara17) & "',N'" & _
                               RepSql(sPara18) & "',N'" & _
                               RepSql(sPara19) & "',N'" & _
                               RepSql(sPara20) & "',N'" & _
                               RepSql(sPara21) & "',N'" & _
                               RepSql(sPara22) & "',N'" & _
                               RepSql(sPara23) & "',N'" & _
                               RepSql(sPara24) & "',N'" & _
                               RepSql(sPara25) & "',N'" & _
                               RepSql(sPara26) & "',N'" & _
                               RepSql(sPara27) & "',N'" & _
                               RepSql(sPara28) & "',N'" & _
                               RepSql(sPara29) & "',N'" & _
                               RepSql(sPara30) & "',N'" & _
                               RepSql(sPara31) & "',N'" & _
                               RepSql(sPara32) & "',N'" & _
                               RepSql(sPara33) & "',N'" & _
                               RepSql(sPara34) & "',N'" & _
                               RepSql(sPara35) & "',N'" & _
                               RepSql(sPara36) & "',N'" & _
                               RepSql(sPara37) & "',N'" & _
                               RepSql(sPara38) & "',N'" & _
                               RepSql(sPara39) & "',N'" & _
                               RepSql(sPara40) & "',N'" & _
                               RepSql(sPara41) & "',N'" & _
                               RepSql(sPara42) & "',N'" & _
                               RepSql(sPara43) & "',N'" & _
                               RepSql(sPara44) & "',N'" & _
                               RepSql(sPara45) & "',N'" & _
                               RepSql(sPara46) & "',N'" & _
                               RepSql(sPara47) & "')"

                    Case "MDY"
                        sSql = "UPDATE Danger_work SET " & _
                               "    com_name=N'" & RepSql(sPara2) & "'," & _
                               "    pet_name=N'" & RepSql(sPara3) & "'," & _
                               "    pet_date=N'" & RepSql(sPara4) & "'," & _
                               "    pet_time=N'" & RepSql(sPara5) & "'," & _
                               "    pet_tel1=N'" & RepSql(sPara6) & "'," & _
                               "    att_name1=N'" & RepSql(sPara7) & "'," & _
                               "    att_tel1=N'" & RepSql(sPara8) & "'," & _
                               "    att_name2=N'" & RepSql(sPara9) & "'," & _
                               "    att_tel2=N'" & RepSql(sPara10) & "'," & _
                               "    att_add1=N'" & RepSql(sPara11) & "'," & _
                               "    att_area=N'" & RepSql(sPara12) & "'," & _
                               "    att_road=N'" & RepSql(sPara13) & "'," & _
                               "    bulid_name=N'" & RepSql(sPara14) & "'," & _
                               "    bulid_code=N'" & RepSql(sPara15) & "'," & _
                               "    work_date_s=N'" & RepSql(sPara16) & "'," & _
                               "    work_date_e=N'" & RepSql(sPara17) & "'," & _
                               "    is_night=N'" & RepSql(sPara18) & "'," & _
                               "    work_floor_s=N'" & RepSql(sPara19) & "'," & _
                               "    work_floor_e=N'" & RepSql(sPara20) & "'," & _
                               "    work_type1=N'" & RepSql(sPara21) & "'," & _
                               "    work_type2=N'" & RepSql(sPara22) & "'," & _
                               "    work_type3=N'" & RepSql(sPara23) & "'," & _
                               "    work_type4=N'" & RepSql(sPara24) & "'," & _
                               "    work_type5=N'" & RepSql(sPara25) & "'," & _
                               "    work_type6=N'" & RepSql(sPara26) & "'," & _
                               "    work_type7=N'" & RepSql(sPara27) & "'," & _
                               "    work_type8=N'" & RepSql(sPara28) & "'," & _
                               "    work_type9=N'" & RepSql(sPara29) & "'," & _
                               "    work_type_memo=N'" & RepSql(sPara30) & "'," & _
                               "    use_tool1=N'" & RepSql(sPara31) & "'," & _
                               "    use_tool2=N'" & RepSql(sPara32) & "'," & _
                               "    use_tool3=N'" & RepSql(sPara33) & "'," & _
                               "    use_tool4=N'" & RepSql(sPara34) & "'," & _
                               "    use_tool5=N'" & RepSql(sPara35) & "'," & _
                               "    use_tool3_memo=N'" & RepSql(sPara36) & "'," & _
                               "    use_tool5_memo=N'" & RepSql(sPara37) & "'," & _
                               "    save_tool=N'" & RepSql(sPara38) & "'," & _
                               "    save_tool_memo=N'" & RepSql(sPara39) & "'," & _
                               "    dan_memo=N'" & RepSql(sPara40) & "'," & _
                               "    dan_memo_memo=N'" & RepSql(sPara41) & "'," & _
                               "    dan_mail=N'" & RepSql(sPara42) & "'," & _
                               "    dan_mail2=N'" & RepSql(sPara43) & "'," & _
                               "    fix_code=N'" & RepSql(sPara46) & "'," & _
                               "    is_holiday=N'" & RepSql(sPara47) & "'" & _
                               " WHERE dan_code=N'" & RepSql(sPara1) & "'"

                    Case "guidance"
                        sSql = "UPDATE guidance SET " & _
                               "      content=N'" & RepSql(sPara1) & "'"
                    Case "FixCode"
                        sSql = "UPDATE Danger_work SET " & _
                               "    fix_code=N'" & RepSql(sPara2) & "'" & _
                               " WHERE dan_code=N'" & RepSql(sPara1) & "'"
                End Select

            Case "LSI021A"
                If String.IsNullOrEmpty(sPara2) Then sPara2 = "未填寫"
                If String.IsNullOrEmpty(sPara3) Then sPara3 = "未填寫"
                If String.IsNullOrEmpty(sPara4) Then sPara4 = "未填寫"
                If String.IsNullOrEmpty(sPara5) Then sPara5 = "未填寫"
                If String.IsNullOrEmpty(sPara6) Then sPara6 = "未填寫"
                If String.IsNullOrEmpty(sPara7) Then sPara7 = "未填寫"
                If String.IsNullOrEmpty(sPara8) Then sPara8 = "未填寫"
                If String.IsNullOrEmpty(sPara9) Then sPara9 = "未填寫"
                If String.IsNullOrEmpty(sPara10) Then sPara10 = "未填寫"
                If String.IsNullOrEmpty(sPara11) Then sPara11 = ""
                'If String.IsNullOrEmpty(sPara12) Then sPara12 = "未填寫"
                If String.IsNullOrEmpty(sPara13) Then sPara13 = ""
                If String.IsNullOrEmpty(sPara14) Then sPara14 = ""
                If String.IsNullOrEmpty(sPara15) Then sPara15 = "未填寫"
                If String.IsNullOrEmpty(sPara16) Then sPara16 = "未填寫"
                If String.IsNullOrEmpty(sPara17) Then sPara17 = "未填寫"
                'If String.IsNullOrEmpty(sPara18) Then sPara18 = "未填寫"
                If String.IsNullOrEmpty(sPara19) Then sPara19 = ""
                If String.IsNullOrEmpty(sPara20) Then sPara20 = ""
                If String.IsNullOrEmpty(sPara21) Then sPara21 = ""
                If String.IsNullOrEmpty(sPara22) Then sPara22 = ""
                If String.IsNullOrEmpty(sPara23) Then sPara23 = "未填寫"
                If String.IsNullOrEmpty(sPara24) Then sPara24 = "未填寫"
                If String.IsNullOrEmpty(sPara25) Then sPara25 = "未填寫"
                'If String.IsNullOrEmpty(sPara26) Then sPara26 = "未填寫"
                If String.IsNullOrEmpty(sPara27) Then sPara27 = ""
                If String.IsNullOrEmpty(sPara28) Then sPara28 = ""
                If String.IsNullOrEmpty(sPara29) Then sPara29 = "未填寫"
                If String.IsNullOrEmpty(sPara30) Then sPara30 = "未填寫"
                If String.IsNullOrEmpty(sPara31) Then sPara31 = "未填寫"
                If String.IsNullOrEmpty(sPara32) Then sPara32 = ""

                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO Roof_work(" & _
                               "roof_cod,com_name,att_area,att_road,att_add1,att_name1,att_tel1,att_name2,att_tel2,roof_mail,pet_date " & _
                               ",roof_tool1,roof_sdate,roof_edate,roof_height,floor_height,roof_memo,tear_tool1,tearon_sdate,tearon_edate " & _
                               ",tearoff_sdate,tearoff_edate,tear_height,tearfloor_height,tear_memo,cage_tool1,cage_sdate,cage_edate,cageroof_height " & _
                               ",cagefloor_height,cage_memo,roof_Source" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "',N'" & _
                               RepSql(sPara9) & "',N'" & _
                               RepSql(sPara10) & "',N'" & _
                               RepSql(sPara11) & "',N'" & _
                               RepSql(sPara12) & "',N'" & _
                               RepSql(sPara13) & "',N'" & _
                               RepSql(sPara14) & "',N'" & _
                               RepSql(sPara15) & "',N'" & _
                               RepSql(sPara16) & "',N'" & _
                               RepSql(sPara17) & "',N'" & _
                               RepSql(sPara18) & "',N'" & _
                               RepSql(sPara19) & "',N'" & _
                               RepSql(sPara20) & "',N'" & _
                               RepSql(sPara21) & "',N'" & _
                               RepSql(sPara22) & "',N'" & _
                               RepSql(sPara23) & "',N'" & _
                               RepSql(sPara24) & "',N'" & _
                               RepSql(sPara25) & "',N'" & _
                               RepSql(sPara26) & "',N'" & _
                               RepSql(sPara27) & "',N'" & _
                               RepSql(sPara28) & "',N'" & _
                               RepSql(sPara29) & "',N'" & _
                               RepSql(sPara30) & "',N'" & _
                               RepSql(sPara31) & "',N'" & _
                               RepSql(sPara32) & "')"
                    Case "MDY"
                        sSql = "UPDATE Roof_work SET " & _
                               "    com_name=N'" & RepSql(sPara2) & "'," & _
                               "    att_area=N'" & RepSql(sPara3) & "'," & _
                               "    att_road=N'" & RepSql(sPara4) & "'," & _
                               "    att_add1=N'" & RepSql(sPara5) & "'," & _
                               "    att_name1=N'" & RepSql(sPara6) & "'," & _
                               "    att_tel1=N'" & RepSql(sPara7) & "'," & _
                               "    att_name2=N'" & RepSql(sPara8) & "'," & _
                               "    att_tel2=N'" & RepSql(sPara9) & "'," & _
                               "    roof_mail=N'" & RepSql(sPara10) & "'," & _
                               "    pet_date=N'" & RepSql(sPara11) & "'," & _
                               "    roof_tool1=N'" & RepSql(sPara12) & "'," & _
                               "    roof_sdate=N'" & RepSql(sPara13) & "'," & _
                               "    roof_edate=N'" & RepSql(sPara14) & "'," & _
                               "    roof_height=N'" & RepSql(sPara15) & "'," & _
                               "    floor_height=N'" & RepSql(sPara16) & "'," & _
                               "    roof_memo=N'" & RepSql(sPara17) & "'," & _
                               "    tear_tool1=N'" & RepSql(sPara18) & "'," & _
                               "    tearon_sdate=N'" & RepSql(sPara19) & "'," & _
                               "    tearon_edate=N'" & RepSql(sPara20) & "'," & _
                               "    tearoff_sdate=N'" & RepSql(sPara21) & "'," & _
                               "    tearoff_edate=N'" & RepSql(sPara22) & "'," & _
                               "    tear_height=N'" & RepSql(sPara23) & "'," & _
                               "    tearfloor_height=N'" & RepSql(sPara24) & "'," & _
                               "    tear_memo=N'" & RepSql(sPara25) & "'," & _
                               "    cage_tool1=N'" & RepSql(sPara26) & "'," & _
                               "    cage_sdate=N'" & RepSql(sPara27) & "'," & _
                               "    cage_edate=N'" & RepSql(sPara28) & "'," & _
                               "    cageroof_height=N'" & RepSql(sPara29) & "'," & _
                               "    cagefloor_height=N'" & RepSql(sPara30) & "'," & _
                               "    cage_memo=N'" & RepSql(sPara31) & "'," & _
                               "    roof_Source=N'" & RepSql(sPara32) & "'" & _
                               " WHERE roof_cod=N'" & RepSql(sPara1) & "'"

                End Select

            Case "LSI024A"
                If String.IsNullOrEmpty(sPara2) Then sPara2 = "未填寫"
                If String.IsNullOrEmpty(sPara3) Then sPara3 = "未填寫"
                If String.IsNullOrEmpty(sPara4) Then sPara4 = "未填寫"
                If String.IsNullOrEmpty(sPara5) Then sPara5 = "未填寫"
                If String.IsNullOrEmpty(sPara6) Then sPara6 = "未填寫"
                If String.IsNullOrEmpty(sPara7) Then sPara7 = "未填寫"
                If String.IsNullOrEmpty(sPara8) Then sPara8 = "未填寫"
                If String.IsNullOrEmpty(sPara9) Then sPara9 = "未填寫"
                If String.IsNullOrEmpty(sPara10) Then sPara10 = "未填寫"
                If String.IsNullOrEmpty(sPara11) Then sPara11 = ""
                If String.IsNullOrEmpty(sPara12) Then sPara12 = "未填寫"
                If String.IsNullOrEmpty(sPara13) Then sPara13 = ""
                If String.IsNullOrEmpty(sPara14) Then sPara14 = ""
                If String.IsNullOrEmpty(sPara15) Then sPara15 = "未填寫"
                If String.IsNullOrEmpty(sPara16) Then sPara16 = "未填寫"
                If String.IsNullOrEmpty(sPara17) Then sPara17 = "未填寫"
                'If String.IsNullOrEmpty(sPara18) Then sPara18 = "未填寫"
                'If String.IsNullOrEmpty(sPara19) Then sPara19 = ""
                'If String.IsNullOrEmpty(sPara20) Then sPara20 = ""
                'If String.IsNullOrEmpty(sPara21) Then sPara21 = ""
                'If String.IsNullOrEmpty(sPara22) Then sPara22 = ""
                'If String.IsNullOrEmpty(sPara23) Then sPara23 = "未填寫"
                'If String.IsNullOrEmpty(sPara24) Then sPara24 = "未填寫"
                'If String.IsNullOrEmpty(sPara25) Then sPara25 = "未填寫"
                'If String.IsNullOrEmpty(sPara26) Then sPara26 = "未填寫"
                'If String.IsNullOrEmpty(sPara27) Then sPara27 = ""
                'If String.IsNullOrEmpty(sPara28) Then sPara28 = ""
                'If String.IsNullOrEmpty(sPara29) Then sPara29 = "未填寫"
                'If String.IsNullOrEmpty(sPara30) Then sPara30 = "未填寫"
                'If String.IsNullOrEmpty(sPara31) Then sPara31 = "未填寫"
                'If String.IsNullOrEmpty(sPara32) Then sPara32 = ""

                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO Precise_Check(" & _
                               "pc_code,com_name,att_tel1,att_name1,att_mail,att_cell_phone,pc_date,eng_name,eng_code,eng_area,eng_road " & _
                               ",eng_add,eng_floor,pc_date_s,pc_date_e,pc_kind1,pc_kind2) VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "',N'" & _
                               RepSql(sPara9) & "',N'" & _
                               RepSql(sPara10) & "',N'" & _
                               RepSql(sPara11) & "',N'" & _
                               RepSql(sPara12) & "',N'" & _
                               RepSql(sPara13) & "',N'" & _
                               RepSql(sPara14) & "',N'" & _
                               RepSql(sPara15) & "',N'" & _
                               RepSql(sPara16) & "',N'" & _
                               RepSql(sPara17) & "')"
                    Case "MDY"
                        sSql = "UPDATE Precise_Check SET " & _
                               "    com_name=N'" & RepSql(sPara2) & "'," & _
                               "    att_tel1=N'" & RepSql(sPara3) & "'," & _
                               "    att_name1=N'" & RepSql(sPara4) & "'," & _
                               "    att_mail=N'" & RepSql(sPara5) & "'," & _
                               "    att_cell_phone=N'" & RepSql(sPara6) & "'," & _
                               "    pc_date=N'" & RepSql(sPara7) & "'," & _
                               "    eng_name=N'" & RepSql(sPara8) & "'," & _
                               "    eng_code=N'" & RepSql(sPara9) & "'," & _
                               "    eng_area=N'" & RepSql(sPara10) & "'," & _
                               "    eng_road=N'" & RepSql(sPara11) & "'," & _
                               "    eng_add=N'" & RepSql(sPara12) & "'," & _
                               "    eng_floor=N'" & RepSql(sPara13) & "'," & _
                               "    pc_date_s=N'" & RepSql(sPara14) & "'," & _
                               "    pc_date_e=N'" & RepSql(sPara15) & "'," & _
                               "    pc_kind1=N'" & RepSql(sPara16) & "'," & _
							   "    pc_kind2=N'" & RepSql(sPara17) & "'" & _
                               " WHERE pc_code=N'" & RepSql(sPara1) & "'"

                End Select

Case "LSI025" '20211102 Ken 新增
                If String.IsNullOrEmpty(sPara2) Then sPara2 = "未填寫"
                If String.IsNullOrEmpty(sPara3) Then sPara3 = "未填寫"
                If String.IsNullOrEmpty(sPara4) Then sPara4 = "未填寫"
                If String.IsNullOrEmpty(sPara5) Then sPara5 = "未填寫"
                If String.IsNullOrEmpty(sPara6) Then sPara6 = "未填寫"
                If String.IsNullOrEmpty(sPara7) Then sPara7 = "未填寫"
                If String.IsNullOrEmpty(sPara8) Then sPara8 = "未填寫"
                If String.IsNullOrEmpty(sPara9) Then sPara9 = "未填寫"
                If String.IsNullOrEmpty(sPara10) Then sPara10 = "未填寫"
                If String.IsNullOrEmpty(sPara11) Then sPara11 = ""
                If String.IsNullOrEmpty(sPara12) Then sPara12 = "未填寫"
                If String.IsNullOrEmpty(sPara13) Then sPara13 = ""
                If String.IsNullOrEmpty(sPara14) Then sPara14 = ""
                If String.IsNullOrEmpty(sPara15) Then sPara15 = "未填寫"
                If String.IsNullOrEmpty(sPara16) Then sPara16 = "未填寫"
                If String.IsNullOrEmpty(sPara17) Then sPara17 = "未填寫"
                'If String.IsNullOrEmpty(sPara18) Then sPara18 = "未填寫"
                'If String.IsNullOrEmpty(sPara19) Then sPara19 = ""
                'If String.IsNullOrEmpty(sPara20) Then sPara20 = ""
                'If String.IsNullOrEmpty(sPara21) Then sPara21 = ""
                'If String.IsNullOrEmpty(sPara22) Then sPara22 = ""
                'If String.IsNullOrEmpty(sPara23) Then sPara23 = "未填寫"
                'If String.IsNullOrEmpty(sPara24) Then sPara24 = "未填寫"
                'If String.IsNullOrEmpty(sPara25) Then sPara25 = "未填寫"
                'If String.IsNullOrEmpty(sPara26) Then sPara26 = "未填寫"
                'If String.IsNullOrEmpty(sPara27) Then sPara27 = ""
                'If String.IsNullOrEmpty(sPara28) Then sPara28 = ""
                'If String.IsNullOrEmpty(sPara29) Then sPara29 = "未填寫"
                'If String.IsNullOrEmpty(sPara30) Then sPara30 = "未填寫"
                'If String.IsNullOrEmpty(sPara31) Then sPara31 = "未填寫"
                'If String.IsNullOrEmpty(sPara32) Then sPara32 = ""

                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO LSI025(" & _
                               "pc_code,com_name,att_tel1,att_name1,att_mail,att_cell_phone,pc_date,eng_name,eng_code,eng_area,eng_road " & _
                               ",eng_add,eng_floor,pc_date_s,pc_date_e,pc_kind1,pc_kind2) VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "',N'" & _
                               RepSql(sPara9) & "',N'" & _
                               RepSql(sPara10) & "',N'" & _
                               RepSql(sPara11) & "',N'" & _
                               RepSql(sPara12) & "',N'" & _
                               RepSql(sPara13) & "',N'" & _
                               RepSql(sPara14) & "',N'" & _
                               RepSql(sPara15) & "',N'" & _
                               RepSql(sPara16) & "',N'" & _
                               RepSql(sPara17) & "')"
                    Case "MDY"
                        sSql = "UPDATE LSI025 SET " & _
                               "    com_name=N'" & RepSql(sPara2) & "'," & _
                               "    att_tel1=N'" & RepSql(sPara3) & "'," & _
                               "    att_name1=N'" & RepSql(sPara4) & "'," & _
                               "    att_mail=N'" & RepSql(sPara5) & "'," & _
                               "    att_cell_phone=N'" & RepSql(sPara6) & "'," & _
                               "    pc_date=N'" & RepSql(sPara7) & "'," & _
                               "    eng_name=N'" & RepSql(sPara8) & "'," & _
                               "    eng_code=N'" & RepSql(sPara9) & "'," & _
                               "    eng_area=N'" & RepSql(sPara10) & "'," & _
                               "    eng_road=N'" & RepSql(sPara11) & "'," & _
                               "    eng_add=N'" & RepSql(sPara12) & "'," & _
                               "    eng_floor=N'" & RepSql(sPara13) & "'," & _
                               "    pc_date_s=N'" & RepSql(sPara14) & "'," & _
                               "    pc_date_e=N'" & RepSql(sPara15) & "'," & _
                               "    pc_kind1=N'" & RepSql(sPara16) & "'," & _
							   "    pc_kind2=N'" & RepSql(sPara17) & "'" & _
                               " WHERE pc_code=N'" & RepSql(sPara1) & "'"

                End Select
				
				
Case "LSI026" '20211102 Ken 新增
                If String.IsNullOrEmpty(sPara2) Then sPara2 = "未填寫"
                If String.IsNullOrEmpty(sPara3) Then sPara3 = "未填寫"
                If String.IsNullOrEmpty(sPara4) Then sPara4 = "未填寫"
                If String.IsNullOrEmpty(sPara5) Then sPara5 = "未填寫"
                If String.IsNullOrEmpty(sPara6) Then sPara6 = "未填寫"
                If String.IsNullOrEmpty(sPara7) Then sPara7 = "未填寫"
                If String.IsNullOrEmpty(sPara8) Then sPara8 = "未填寫"
                If String.IsNullOrEmpty(sPara9) Then sPara9 = "未填寫"
                If String.IsNullOrEmpty(sPara10) Then sPara10 = "未填寫"
                If String.IsNullOrEmpty(sPara11) Then sPara11 = ""
                If String.IsNullOrEmpty(sPara12) Then sPara12 = "未填寫"
                If String.IsNullOrEmpty(sPara13) Then sPara13 = ""
                If String.IsNullOrEmpty(sPara14) Then sPara14 = ""
                If String.IsNullOrEmpty(sPara15) Then sPara15 = "未填寫"
                If String.IsNullOrEmpty(sPara16) Then sPara16 = "未填寫"
                If String.IsNullOrEmpty(sPara17) Then sPara17 = "未填寫"
                'If String.IsNullOrEmpty(sPara18) Then sPara18 = "未填寫"
                'If String.IsNullOrEmpty(sPara19) Then sPara19 = ""
                'If String.IsNullOrEmpty(sPara20) Then sPara20 = ""
                'If String.IsNullOrEmpty(sPara21) Then sPara21 = ""
                'If String.IsNullOrEmpty(sPara22) Then sPara22 = ""
                'If String.IsNullOrEmpty(sPara23) Then sPara23 = "未填寫"
                'If String.IsNullOrEmpty(sPara24) Then sPara24 = "未填寫"
                'If String.IsNullOrEmpty(sPara25) Then sPara25 = "未填寫"
                'If String.IsNullOrEmpty(sPara26) Then sPara26 = "未填寫"
                'If String.IsNullOrEmpty(sPara27) Then sPara27 = ""
                'If String.IsNullOrEmpty(sPara28) Then sPara28 = ""
                'If String.IsNullOrEmpty(sPara29) Then sPara29 = "未填寫"
                'If String.IsNullOrEmpty(sPara30) Then sPara30 = "未填寫"
                'If String.IsNullOrEmpty(sPara31) Then sPara31 = "未填寫"
                'If String.IsNullOrEmpty(sPara32) Then sPara32 = ""

                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO LSI026(" & _
                               "pc_code,com_name,att_tel1,att_name1,att_mail,att_cell_phone,pc_date,eng_name,eng_code,eng_area,eng_road " & _
                               ",eng_add,eng_floor,pc_date_s,pc_date_e,pc_kind1,pc_kind2) VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "',N'" & _
                               RepSql(sPara9) & "',N'" & _
                               RepSql(sPara10) & "',N'" & _
                               RepSql(sPara11) & "',N'" & _
                               RepSql(sPara12) & "',N'" & _
                               RepSql(sPara13) & "',N'" & _
                               RepSql(sPara14) & "',N'" & _
                               RepSql(sPara15) & "',N'" & _
                               RepSql(sPara16) & "',N'" & _
                               RepSql(sPara17) & "')"
                    Case "MDY"
                        sSql = "UPDATE LSI026 SET " & _
                               "    com_name=N'" & RepSql(sPara2) & "'," & _
                               "    att_tel1=N'" & RepSql(sPara3) & "'," & _
                               "    att_name1=N'" & RepSql(sPara4) & "'," & _
                               "    att_mail=N'" & RepSql(sPara5) & "'," & _
                               "    att_cell_phone=N'" & RepSql(sPara6) & "'," & _
                               "    pc_date=N'" & RepSql(sPara7) & "'," & _
                               "    eng_name=N'" & RepSql(sPara8) & "'," & _
                               "    eng_code=N'" & RepSql(sPara9) & "'," & _
                               "    eng_area=N'" & RepSql(sPara10) & "'," & _
                               "    eng_road=N'" & RepSql(sPara11) & "'," & _
                               "    eng_add=N'" & RepSql(sPara12) & "'," & _
                               "    eng_floor=N'" & RepSql(sPara13) & "'," & _
                               "    pc_date_s=N'" & RepSql(sPara14) & "'," & _
                               "    pc_date_e=N'" & RepSql(sPara15) & "'," & _
                               "    pc_kind1=N'" & RepSql(sPara16) & "'," & _
							   "    pc_kind2=N'" & RepSql(sPara17) & "'" & _
                               " WHERE pc_code=N'" & RepSql(sPara1) & "'"

                End Select

            Case "LSI015A"
                If String.IsNullOrEmpty(sPara2) Then sPara2 = "未填寫"
                If String.IsNullOrEmpty(sPara3) Then sPara3 = "未填寫"
                If String.IsNullOrEmpty(sPara4) Then sPara4 = "未填寫"
                If String.IsNullOrEmpty(sPara5) Then sPara5 = "未填寫"
                If String.IsNullOrEmpty(sPara8) Then sPara8 = "未填寫"
                If String.IsNullOrEmpty(sPara10) Then sPara10 = "未填寫"
                If String.IsNullOrEmpty(sPara11) Then sPara11 = "未填寫"
                If String.IsNullOrEmpty(sPara12) Then sPara12 = "未填寫"
                If String.IsNullOrEmpty(sPara13) Then sPara13 = "未填寫"
                If String.IsNullOrEmpty(sPara14) Then sPara14 = "未填寫"
                If String.IsNullOrEmpty(sPara15) Then sPara15 = "未填寫"
                If String.IsNullOrEmpty(sPara16) Then sPara16 = "未填寫"
                If String.IsNullOrEmpty(sPara17) Then sPara17 = "未填寫"
                If String.IsNullOrEmpty(sPara18) Then sPara18 = "未填寫"
                If String.IsNullOrEmpty(sPara19) Then sPara19 = "未填寫"
                If String.IsNullOrEmpty(sPara26) Then sPara26 = "未填寫"

                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO Danger_check(" & _
                               "danc_code,eng_name,pet_date,Pet_time,eng_add,danc_kind,danc_step,danc_div,danc_type,com_name" & _
                               ",com_idno,com_add,com_tel1,att_name1,att_name2,con_name,con_tel1,con_tel2,con_email,danc_state" & _
                               ",result,usr_code,ins_date,att_area,att_road,con_pw" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "',N'" & _
                               RepSql(sPara9) & "',N'" & _
                               RepSql(sPara10) & "',N'" & _
                               RepSql(sPara11) & "',N'" & _
                               RepSql(sPara12) & "',N'" & _
                               RepSql(sPara13) & "',N'" & _
                               RepSql(sPara14) & "',N'" & _
                               RepSql(sPara15) & "',N'" & _
                               RepSql(sPara16) & "',N'" & _
                               RepSql(sPara17) & "',N'" & _
                               RepSql(sPara18) & "',N'" & _
                               RepSql(sPara19) & "',N'" & _
                               RepSql(sPara20) & "',N'" & _
                               RepSql(sPara21) & "',N'" & _
                               RepSql(sPara22) & "',N'" & _
                               RepSql(sPara23) & "',N'" & _
                               RepSql(sPara24) & "',N'" & _
                               RepSql(sPara25) & "',N'" & _
                               RepSql(sPara26) & "')"

                    Case "MDY"
                        sSql = "UPDATE Danger_check SET " & _
                               "    eng_name=N'" & RepSql(sPara2) & "'," & _
                               "    pet_date=N'" & RepSql(sPara3) & "'," & _
                               "    Pet_time=N'" & RepSql(sPara4) & "'," & _
                               "    eng_add=N'" & RepSql(sPara5) & "'," & _
                               "    danc_kind=N'" & RepSql(sPara6) & "'," & _
                               "    danc_step=N'" & RepSql(sPara7) & "'," & _
                               "    danc_div=N'" & RepSql(sPara8) & "'," & _
                               "    danc_type=N'" & RepSql(sPara9) & "'," & _
                               "    com_name=N'" & RepSql(sPara10) & "'," & _
                               "    com_idno=N'" & RepSql(sPara11) & "'," & _
                               "    com_add=N'" & RepSql(sPara12) & "'," & _
                               "    com_tel1=N'" & RepSql(sPara13) & "'," & _
                               "    att_name1=N'" & RepSql(sPara14) & "'," & _
                               "    att_name2=N'" & RepSql(sPara15) & "'," & _
                               "    con_name=N'" & RepSql(sPara16) & "'," & _
                               "    con_tel1=N'" & RepSql(sPara17) & "'," & _
                               "    con_tel2=N'" & RepSql(sPara18) & "'," & _
                               "    con_email=N'" & RepSql(sPara19) & "'," & _
                               "    danc_state=N'" & RepSql(sPara20) & "'," & _
                               "    result=N'" & RepSql(sPara21) & "'," & _
                               "    att_area=N'" & RepSql(sPara24) & "'," & _
                               "    att_road=N'" & RepSql(sPara25) & "'," & _
                               "    con_pw=N'" & RepSql(sPara26) & "'" & _
                               " WHERE danc_code=N'" & RepSql(sPara1) & "'"
                End Select

            Case "LSI016A"
                If String.IsNullOrEmpty(sPara2) Then sPara2 = "未填寫"
                If String.IsNullOrEmpty(sPara4) Then sPara4 = "未填寫"

                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO e_news(" & _
                               "news_code,news_title,news_type,news_contect,news_pic1,usr_code,ins_date" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               Replace(sPara4, "'", "''") & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "')"

                    Case "MDY"
                        sSql = "UPDATE e_news SET " & _
                               "    news_title=N'" & RepSql(sPara2) & "'," & _
                               "    news_type=N'" & RepSql(sPara3) & "'," & _
                               "    news_contect=N'" & Replace(sPara4, "'", "''") & "'," & _
                               "    news_pic1=N'" & RepSql(sPara5) & "'," & _
                               "    usr_code=N'" & RepSql(sPara6) & "'," & _
                               "    ins_date=N'" & RepSql(sPara7) & "'" & _
                               " WHERE news_code=N'" & RepSql(sPara1) & "'"
                End Select

            Case "LSI017A"
                If String.IsNullOrEmpty(sPara5) Then sPara5 = "未填寫"

                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO APP_VERSION(" & _
                               "app_code,version,app_path,app_type,cmemo,update_date,usr_code,ins_date" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "')"

                    Case "MDY"
                        sSql = "UPDATE APP_VERSION SET " & _
                               "    version=N'" & RepSql(sPara2) & "'," & _
                               "    app_path=N'" & RepSql(sPara3) & "'," & _
                               "    app_type=N'" & RepSql(sPara4) & "'," & _
                               "    cmemo=N'" & RepSql(sPara5) & "'," & _
                               "    update_date=N'" & RepSql(sPara6) & "'" & _
                               " WHERE app_code=N'" & RepSql(sPara1) & "'"
                End Select

            Case "LSI019A"
                Select Case sType
                    Case "ADD", "COPY"
                        sSql = "INSERT INTO LABOR_INSPECT(" & _
                               "chk_date1,chk_date2,com_name,eng_name,work_staus,usr_code,ins_date" & _
                               ") VALUES (N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "',N'" & _
                               RepSql(sPara7) & "',N'" & _
                               RepSql(sPara8) & "')"

                    Case "MDY"
                        sSql = "UPDATE LABOR_INSPECT SET " & _
                               "    chk_date1=N'" & RepSql(sPara2) & "'," & _
                               "    chk_date2=N'" & RepSql(sPara3) & "'," & _
                               "    com_name=N'" & RepSql(sPara4) & "'," & _
                               "    eng_name=N'" & RepSql(sPara5) & "'," & _
                               "    work_staus=N'" & RepSql(sPara6) & "'" & _
                               " WHERE labor_inspect=" & RepSql(sPara1)

                    Case "DEL"
                        sSql = "DELETE FROM LABOR_INSPECT"
                End Select

            Case "SendNewsMail"
                Select Case sType
                    Case "Batch"
                        sSql = "INSERT INTO E_PAPER_TMP(" & _
                               "sch_code,per_mail,ins_date,ins_time" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "')"

                    Case "Result"
                        sSql = "UPDATE e_paper_sch SET " & _
                               "    send_state=N'" & RepSql(sPara2) & "'," & _
                               "    last_date=N'" & RepSql(sPara3) & "'," & _
                               "    count=count+1" & _
                               " WHERE sch_code='" & RepSql(sPara1) & "'"
                End Select

            Case "MyService"
                Select Case sType
                    Case "ADD"
                        sSql = "INSERT INTO GEO_LOCATION(" & _
                               "ms_code,latitude,longitude" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "')"

                    Case "MDY"
                        sSql = "UPDATE GEO_LOCATION SET " & _
                               "    latitude=N'" & RepSql(sPara2) & "'," & _
                               "    longitude=N'" & RepSql(sPara3) & "'" & _
                               " WHERE ms_code='" & RepSql(sPara1) & "'"
                End Select

            Case "APP_Log"
                sSql = "INSERT INTO APP_LOG(" & _
                               "apl_code,apl_type,apl_date,apl_time,app_type,usr_code" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "',N'" & _
                               RepSql(sPara3) & "',N'" & _
                               RepSql(sPara4) & "',N'" & _
                               RepSql(sPara5) & "',N'" & _
                               RepSql(sPara6) & "')"

            Case "Meeting_Date"
                sSql = "INSERT INTO Meeting_Date(" & _
                               "mrd_date,Day" & _
                               ") VALUES (N'" & _
                               RepSql(sPara1) & "',N'" & _
                               RepSql(sPara2) & "')"
            Case "UserIdCheck"
                Select Case sType
                    Case "True"
                        sSql = "UPDATE BDP080 SET " & _
                               "    count=0" & _
                               " WHERE usr_code='" & RepSql(sPara1) & "'"

                    Case "False"
                        sSql = "UPDATE BDP080 SET " & _
                               "    count=count+1" & _
                               " WHERE usr_code='" & RepSql(sPara1) & "'"

                    Case "StopUse"
                        sSql = "UPDATE BDP080 SET " & _
                               "    is_use='N'" & _
                               " WHERE is_use='Y' AND count>=5"
                End Select
        End Select

        'Try
        '    Using (SqlCon)
        '        SqlCmd.ExecuteNonQuery()
        '        SqlCon.Close()
        '    End Using
        '    Return True
        'Catch ex As Exception
        '    Return False
        'End Try
        Return SaveData(sSql)
    End Function

    ''' <summary>
    ''' '建立資料庫連線
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ConnectDB() As SqlConnection
        Dim Connection_Db As SqlConnection
        Connection_Db = New SqlConnection(Get_ConnStr("con_db"))
        Connection_Db.Open()
        ConnectDB = Connection_Db
    End Function

    ''' <summary>
    ''' 過濾sql語法
    ''' </summary>
    ''' <param name="p_str"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RepSql(ByVal p_str As String) As String
        If String.IsNullOrEmpty(p_str) = True Then Return ""

        Dim sStr As String = p_str
        sStr = Replace(Trim(sStr & ""), "'", "’")
        sStr = Replace(Trim(sStr & ""), ",", "，")
        sStr = Replace(Trim(sStr & ""), "!", "！")
        sStr = Replace(Trim(sStr & ""), "?", "？")
        sStr = Replace(Trim(sStr & ""), "~", "～")
        'sStr = Replace(Trim(sStr & ""), ":", "：")
        sStr = Replace(Trim(sStr & ""), ";", "；")
        sStr = Replace(Trim(sStr & ""), "^", "︿")
        sStr = Replace(Trim(sStr & ""), ">", "＞")
        sStr = Replace(Trim(sStr & ""), "<", "＜")
        'sStr = Replace(Trim(sStr & ""), "/", "／")
        'sStr = Replace(Trim(sStr & ""), "\", "＼")
        sStr = Replace(Trim(sStr & ""), "{", "｛")
        sStr = Replace(Trim(sStr & ""), "}", "｝")
        sStr = Replace(Trim(sStr & ""), "[", "［")
        sStr = Replace(Trim(sStr & ""), "]", "］")
        sStr = Replace(Trim(sStr & ""), "*", "＊")
        sStr = Replace(Trim(sStr & ""), "%", "％")
        sStr = Replace(Trim(sStr & ""), "$", "＄")
        'sStr = Replace(Trim(sStr & ""), "#", "＃")
        sStr = Replace(Trim(sStr & ""), "&", "＆")
        '2016/12/29 yumi
        sStr = Replace(Trim(sStr & ""), "script", "")
        sStr = Replace(Trim(sStr & ""), "XP_", "")
        sStr = Replace(Trim(sStr & ""), "SP_", "")
        sStr = Replace(Trim(sStr & ""), "\'", "")
        sStr = Replace(Trim(sStr & ""), "\'""", "")
        sStr = Replace(Trim(sStr & ""), "SP_", "")
        sStr = Replace(Trim(sStr & ""), "SP_", "")
        '過濾攻擊字元
        sStr = Regex.Replace(sStr, "\b(exec(ute)?|select|update|insert|delete|drop|create)\b|[;']|(-{2})|(/\*.*\*/)", String.Empty, RegexOptions.IgnoreCase)
        Return sStr
    End Function

    ''' <summary>
    ''' 傳入一Sql傳回一DataTable
    ''' </summary>
    ''' <param name="Csql">Sql語法</param>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    Private Function Get_DataTable(ByVal Csql As String) As DataTable
        Get_DataTable = Nothing
        Try
            If Csql <> "" Then
                '建立連線
                Dim Con_db As SqlConnection
                Con_db = Me.ConnectDB

                Using (Con_db)
                    '建立adapter
                    Dim Fun_Adpt As SqlDataAdapter
                    '建立DataTable
                    Dim Fun_Dt1 As DataTable = New DataTable

                    Fun_Adpt = New SqlDataAdapter(Csql, Con_db)
                    Fun_Adpt.Fill(Fun_Dt1)
                    Get_DataTable = Fun_Dt1
                    Con_db.Close()
                End Using
            Else
                Return New DataTable
            End If
        Catch
            'Response.Write(Csql)
            'Response.End()
        End Try
    End Function

    ''' <summary>
    ''' 取得設定程式代碼之DataTable
    ''' </summary>
    ''' <param name="i_prgcode">程式代碼</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function Get_FieldTable(ByVal i_prgcode As String) As DataTable
        '傳入一prgcode傳回一DataTable()
        Get_FieldTable = Nothing
        Dim tmpDT As New DataTable
        Dim tmpAdpt As SqlDataAdapter
        Dim SqlCmd As New SqlCommand
        Dim SqlCon As SqlConnection = New SqlConnection(Get_ConnStr(""))

        'Dim sSql As String = "SELECT * FROM BDP220 WHERE prg_code='" & i_prgcode & "' AND is_use='Y' ORDER BY scr_no"
        Dim sSql As String = "SELECT * FROM BDP220 WHERE prg_code=@prg_code AND is_use='Y' ORDER BY scr_no"
        SqlCmd = New SqlCommand(sSql, SqlCon)
        SqlCmd.Parameters.AddWithValue("@prg_code", RepSql(i_prgcode))
        Try
            tmpAdpt = New SqlDataAdapter(SqlCmd)
            tmpAdpt.Fill(tmpDT)
            Return tmpDT
            'If sSql <> "" Then
            '    '建立連線
            '    Dim Con_db As SqlConnection
            '    Con_db = Me.ConnectDB

            '    Using (Con_db)
            '        '建立adapter
            '        Dim Fun_Adpt As SqlDataAdapter
            '        '建立DataTable
            '        Dim Fun_Dt1 As DataTable = New DataTable

            '        Fun_Adpt = New SqlDataAdapter(sSql, Con_db)
            '        Fun_Adpt.Fill(Fun_Dt1)
            '        Get_FieldTable = Fun_Dt1
            '        Con_db.Close()
            '    End Using
            'Else
            '    Return New DataTable
            'End If
        Catch
            'Response.Write(sSql)
            'Response.End()
        End Try
    End Function

    ''' <summary>
    ''' 取得設定程式代碼之DataSet
    ''' </summary>
    ''' <param name="i_prgcode">程式代碼</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function Get_FieldSet(ByVal i_prgcode As String) As DataSet
        Dim tmpDS As New DataSet
        Try
            tmpDS.Tables.Add(Get_FieldTable(i_prgcode))
        Catch ex As Exception
            tmpDS.Tables.Add(New DataTable)
        End Try
        Return tmpDS
    End Function

    ''' <summary>
    ''' 取得欄位名稱
    ''' </summary>
    ''' <param name="i_prgcode">程式代碼</param>
    ''' <returns>欄位資料字串</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function Get_FieldName(ByVal i_prgcode As String) As String
        Get_FieldName = ""

        Dim Dt_Field As DataTable = Get_FieldTable(i_prgcode)

        For i As Integer = 0 To Dt_Field.Rows.Count - 1
            If i <> 0 Then Get_FieldName &= "|"
            Get_FieldName &= Dt_Field.Rows(i).Item("field_code")
        Next
    End Function

    ''' <summary>
    ''' 取得修改前資料
    ''' </summary>
    ''' <param name="Sql_str">查詢資料語法</param>
    ''' <param name="i_prgcode">程式代碼</param>
    ''' <returns>資料字串</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function Get_OldData(ByVal Sql_str As String, ByVal i_prgcode As String) As String
        Get_OldData = ""

        Dim Dt_Field As DataTable = Get_FieldTable(i_prgcode)
        Dim dt_tmp As DataTable = Get_DataTable(Sql_str)

        If dt_tmp.Rows.Count > 0 Then
            For i As Integer = 0 To Dt_Field.Rows.Count - 1
                If i <> 0 Then Get_OldData &= "_|@"
                Get_OldData &= dt_tmp.Rows(0).Item(Dt_Field.Rows(i).Item("field_code"))
            Next
        Else
            For i As Integer = 0 To Dt_Field.Rows.Count - 1
                If i <> 0 Then Get_OldData &= "_|@"
            Next
        End If
    End Function

    ''' <summary>
    ''' 取得單支程式權限字串
    ''' </summary>
    ''' <param name="usr_code">使用者代碼</param>
    ''' <param name="prg_code">程式代碼</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function Get_limit(ByVal usr_code As String, ByVal prg_code As String) As String
        '檢查參數
        If String.IsNullOrEmpty(usr_code) Or String.IsNullOrEmpty(prg_code) Then Return ""

        Dim dt_tmp As New DataTable
        Dim dt_tmp2 As New DataTable
        Dim Sql_Str As String
        Dim Fun_Str As String
        Dim Dt_fun1 As New DataTable
        Dim sLimitStr As String = ""
        Dim tmpAdpt As SqlDataAdapter
        Dim SqlCmd As New SqlCommand
        Dim SqlCon As SqlConnection = New SqlConnection(Get_ConnStr(""))

        '超級使用者直接取得最大權限
        If UCase(usr_code) = "WOHER" Then
            sLimitStr = "IMDAPCSTE"
            Return sLimitStr
        End If

        '單程式停用或超過使用期限回傳無權限
        'Sql_Str = "SELECT prg_code from BDP030 " & _
        '          " WHERE is_use = 'Y' " & _
        '          "   AND '" & Format(Now, "yyyy/MM/dd") & "' between s_date and e_date " & _
        '          "   AND prg_code='" & prg_code & "'"
        Sql_Str = "SELECT prg_code from BDP030 " & _
                  " WHERE is_use = 'Y' " & _
                  "   AND @date between s_date and e_date " & _
                  "   AND prg_code=@prg_code"
        SqlCmd = New SqlCommand(Sql_Str, SqlCon)
        SqlCmd.Parameters.AddWithValue("@date", Format(Now, "yyyy/MM/dd"))
        SqlCmd.Parameters.AddWithValue("@prg_code", prg_code)
        tmpAdpt = New SqlDataAdapter(SqlCmd)
        tmpAdpt.Fill(dt_tmp2)
        'dt_tmp2 = Get_DataTable(Sql_Str)
        If Not dt_tmp2.Rows.Count > 0 Then
            sLimitStr = ""
            Return sLimitStr
        End If
        'Else

        Dim sLimitType As String = "B" '權限類別 A:個人 B:角色 C:混合
        'sLimitType = Get_SysPara("sys_limit_type", "B")
        'Sql_Str = "SELECT Limit_type FROM BDP080 WHERE usr_code='" & usr_code & "'"
        Sql_Str = "SELECT Limit_type FROM BDP080 WHERE usr_code=@usr_code"
        SqlCmd = New SqlCommand(Sql_Str, SqlCon)
        SqlCmd.Parameters.AddWithValue("@usr_code", usr_code)
        tmpAdpt = New SqlDataAdapter(SqlCmd)
        tmpAdpt.Fill(dt_tmp)
        'dt_tmp = Get_DataTable(Sql_Str)
        If dt_tmp.Rows.Count > 0 Then sLimitType = dt_tmp.Rows(0).Item(0).ToString

        Select Case sLimitType
            Case "A" '個人
                'Fun_Str = "SELECT BDP090.limit_str " & _
                '          "  FROM BDP090 " & _
                '          "  LEFT JOIN BDP080 ON BDP090.usr_code = BDP080.usr_code " & _
                '          " WHERE BDP080.usr_code ='" & Replace(usr_code, "'", "''") & "'" & _
                '          "   AND BDP090.prg_code ='" & Replace(prg_code, "'", "''") & "'" & _
                '          "   AND BDP090.is_use ='Y'"
                Fun_Str = "SELECT BDP090.limit_str " & _
                          "  FROM BDP090 " & _
                          "  LEFT JOIN BDP080 ON BDP090.usr_code = BDP080.usr_code " & _
                          " WHERE BDP080.usr_code =@usr_code" & _
                          "   AND BDP090.prg_code =@prg_code" & _
                          "   AND BDP090.is_use ='Y'"
                SqlCmd = New SqlCommand(Fun_Str, SqlCon)
                SqlCmd.Parameters.AddWithValue("@usr_code", usr_code)
                SqlCmd.Parameters.AddWithValue("@prg_code", prg_code)
                tmpAdpt = New SqlDataAdapter(SqlCmd)
                tmpAdpt.Fill(Dt_fun1)
                'Dt_fun1 = Me.Get_DataTable(Fun_Str)
                If Dt_fun1.Rows.Count > 0 Then
                    sLimitStr = Trim(Dt_fun1.Rows(0).Item("limit_str") & "")
                Else
                    sLimitStr = ""
                End If

            Case "B" '角色
                'Fun_Str = "SELECT BDP091.limit_str " & _
                '          "  FROM BDP091 " & _
                '          "  LEFT JOIN BDP080 ON BDP091.grp_code = BDP080.grp_code " & _
                '          " WHERE BDP080.usr_code ='" & Replace(usr_code, "'", "''") & "'" & _
                '          "   AND BDP091.prg_code ='" & Replace(prg_code, "'", "''") & "'" & _
                '          "   AND BDP091.is_use ='Y'"
                Fun_Str = "SELECT BDP091.limit_str " & _
                          "  FROM BDP091 " & _
                          "  LEFT JOIN BDP080 ON BDP091.grp_code = BDP080.grp_code " & _
                          " WHERE BDP080.usr_code =@usr_code" & _
                          "   AND BDP091.prg_code =@prg_code" & _
                          "   AND BDP091.is_use ='Y'"
                SqlCmd = New SqlCommand(Fun_Str, SqlCon)
                SqlCmd.Parameters.AddWithValue("@usr_code", usr_code)
                SqlCmd.Parameters.AddWithValue("@prg_code", prg_code)
                tmpAdpt = New SqlDataAdapter(SqlCmd)
                tmpAdpt.Fill(Dt_fun1)
                'Dt_fun1 = Me.Get_DataTable(Fun_Str)
                If Dt_fun1.Rows.Count > 0 Then
                    sLimitStr = Trim(Dt_fun1.Rows(0).Item("limit_str") & "")
                Else
                    sLimitStr = ""
                End If

            Case "C" '科室
                'Fun_Str = "SELECT BDP092.limit_str " & _
                '          "  FROM BDP092 " & _
                '          "  LEFT JOIN BDP080 ON BDP092.dep_code = BDP080.dep_code " & _
                '          " WHERE BDP080.usr_code ='" & Replace(usr_code, "'", "''") & "'" & _
                '          "   AND BDP092.prg_code ='" & Replace(prg_code, "'", "''") & "'" & _
                '          "   AND BDP092.is_use ='Y'"
                Fun_Str = "SELECT BDP092.limit_str " & _
                          "  FROM BDP092 " & _
                          "  LEFT JOIN BDP080 ON BDP092.dep_code = BDP080.dep_code " & _
                          " WHERE BDP080.usr_code =@usr_code" & _
                          "   AND BDP092.prg_code =@prg_code" & _
                          "   AND BDP092.is_use ='Y'"
                SqlCmd = New SqlCommand(Fun_Str, SqlCon)
                SqlCmd.Parameters.AddWithValue("@usr_code", usr_code)
                SqlCmd.Parameters.AddWithValue("@prg_code", prg_code)
                tmpAdpt = New SqlDataAdapter(SqlCmd)
                tmpAdpt.Fill(Dt_fun1)
                'Dt_fun1 = Me.Get_DataTable(Fun_Str)
                If Dt_fun1.Rows.Count > 0 Then
                    sLimitStr = Trim(Dt_fun1.Rows(0).Item("limit_str") & "")
                Else
                    sLimitStr = ""
                End If
        End Select
        'End If

        Return sLimitStr
    End Function

    ''' <summary>
    ''' 取得程式名稱
    ''' </summary>
    ''' <param name="prg_code">程式代碼</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function Get_PrgName(ByVal prg_code As String) As String
        Dim Fun_Str As String
        Dim Dt_fun1 As New DataTable
        Dim tmpAdpt As SqlDataAdapter
        Dim SqlCmd As New SqlCommand
        Dim SqlCon As SqlConnection = New SqlConnection(Get_ConnStr(""))

        'Fun_Str = "SELECT * FROM BDP030 WHERE prg_code ='" & RepSql(prg_code) & "'"
        Fun_Str = "SELECT * FROM BDP030 WHERE prg_code =@prg_code"
        SqlCmd = New SqlCommand(Fun_Str, SqlCon)
        SqlCmd.Parameters.AddWithValue("@prg_code", prg_code)
        tmpAdpt = New SqlDataAdapter(SqlCmd)
        tmpAdpt.Fill(Dt_fun1)
        'Dt_fun1 = Me.Get_DataTable(Fun_Str)
        If Dt_fun1.Rows.Count > 0 Then
            Return Trim(Dt_fun1.Rows(0).Item("prg_name") & "")
        Else
            Return ""
        End If
    End Function

    ''' <summary>
    ''' 取得連線字串
    ''' </summary>
    ''' <param name="sConnName">連線名稱</param>
    ''' <returns>連線字串</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function Get_ConnStr(ByVal sConnName As String) As String
        Dim DS As String = WebConfigurationManager.AppSettings("DS")
        Dim IC As String = ChangeString(WebConfigurationManager.AppSettings("IC"))
        Dim ID As String = ChangeString(WebConfigurationManager.AppSettings("ID"))
        Dim PW As String = ChangeString(WebConfigurationManager.AppSettings("PW"))
       ' Dim con_str As String = "Data Source=" & DS & ";Initial Catalog=" & IC & ";User ID=" & ID & ";Password=" & PW & ";Pooling=True"
	Dim con_str As String = "Data Source=172.25.160.159,50217;Initial Catalog=lio-Servicein;User ID=startnet;Password=Pw*^))KmZ.1422;Pooling=True"
		
        Return con_str
    End Function

    ''' <summary>
    ''' 存檔模組
    ''' </summary>
    ''' <param name="Csql">sql語法</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SaveData(ByVal Csql As String) As Boolean
        Dim Con_db As SqlConnection
        Dim Sql_cmd As SqlCommand

        Con_db = Me.ConnectDB
        Try
            Using (Con_db)
                Sql_cmd = New SqlCommand(Csql, Con_db)
                Sql_cmd.ExecuteNonQuery()
                SaveData = True
                Con_db.Close()
            End Using
        Catch
            'HttpContext.Current.Response.Write(Csql)
            'HttpContext.Current.Response.End()
            Return False
        End Try

        Return True
    End Function

    ''' <summary>
    ''' 取得系統參數
    ''' </summary>
    ''' <param name="ParaName">參數代碼</param>
    ''' <param name="Default_value">預設參數</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function Get_SysPara(ByVal ParaName As String, ByVal Default_value As String) As String
        Dim Fun_Str As String
        Dim Dt_fun1 As New DataTable
        Dim tmpAdpt As SqlDataAdapter
        Dim SqlCmd As New SqlCommand
        Dim SqlCon As SqlConnection = New SqlConnection(Get_ConnStr(""))

        'Fun_Str = "SELECT par_value FROM BDP000 WHERE par_name='" & ParaName & "'"
        Fun_Str = "SELECT par_value FROM BDP000 WHERE par_name=@ParaName"
        SqlCmd = New SqlCommand(Fun_Str, SqlCon)
        SqlCmd.Parameters.AddWithValue("@ParaName", ParaName)
        tmpAdpt = New SqlDataAdapter(SqlCmd)
        tmpAdpt.Fill(Dt_fun1)
        'Dt_fun1 = Me.Get_DataTable(Fun_Str)
        If Dt_fun1.Rows.Count > 0 Then
            Get_SysPara = Trim(Dt_fun1.Rows(0).Item(0) & "")
        Else
            Get_SysPara = Default_value
        End If
    End Function

    <WebMethod()> _
    Public Function Get_Guidance() As String
        Dim Fun_Str As String
        Dim Dt_fun1 As DataTable

        Fun_Str = "SELECT content FROM guidance"
        Dt_fun1 = Me.Get_DataTable(Fun_Str)
        If Dt_fun1.Rows.Count > 0 Then
            Return Trim(Dt_fun1.Rows(0).Item(0) & "")
        Else
            Return ""
        End If
    End Function

    <WebMethod()> _
    Public Function Get_Guidance1() As String
        Dim Fun_Str As String
        Dim Dt_fun1 As DataTable

        Fun_Str = "SELECT content FROM guidance1"
        Dt_fun1 = Me.Get_DataTable(Fun_Str)
        If Dt_fun1.Rows.Count > 0 Then
            Return Trim(Dt_fun1.Rows(0).Item(0) & "")
        Else
            Return ""
        End If
    End Function

    <WebMethod()> _
    Public Function Get_YComplain() As String
        Dim Fun_Str As String
        Dim Dt_fun1 As DataTable

        Fun_Str = "SELECT content FROM YComplain"
        Dt_fun1 = Me.Get_DataTable(Fun_Str)
        If Dt_fun1.Rows.Count > 0 Then
            Return Trim(Dt_fun1.Rows(0).Item(0) & "")
        Else
            Return ""
        End If
    End Function

    <WebMethod()> _
    Public Function Get_NComplain() As String
        Dim Fun_Str As String
        Dim Dt_fun1 As DataTable

        Fun_Str = "SELECT content FROM NComplain"
        Dt_fun1 = Me.Get_DataTable(Fun_Str)
        If Dt_fun1.Rows.Count > 0 Then
            Return Trim(Dt_fun1.Rows(0).Item(0) & "")
        Else
            Return ""
        End If
    End Function

    <WebMethod()> _
    Public Function Get_Complain() As String
        Dim Fun_Str As String
        Dim Dt_fun1 As DataTable

        Fun_Str = "SELECT content FROM Complain"
        Dt_fun1 = Me.Get_DataTable(Fun_Str)
        If Dt_fun1.Rows.Count > 0 Then
            Return Trim(Dt_fun1.Rows(0).Item(0) & "")
        Else
            Return ""
        End If
    End Function

    ''' <summary>
    ''' 取得自動編號欄位目前最大值(可設條件)
    ''' </summary>
    ''' <param name="TableName">TABLE</param>
    ''' <param name="FieldName">欄位名稱</param>
    ''' <param name="sWhere">where子句</param>
    ''' <returns>最大值</returns>
    ''' <remarks>by linda</remarks>
    <WebMethod()> _
    Public Function Get_AutoIntMax(ByVal TableName As String, ByVal FieldName As String, ByVal sWhere As String) As String
        '取得自動編號欄位目前最大值
        Dim Fun_Str As String
        Dim Dt_fun1 As DataTable

        Fun_Str = "SELECT MAX(" & FieldName & ") AS Max_no FROM " & TableName & sWhere & " GROUP BY " & FieldName & " ORDER BY " & FieldName & " DESC"
        Dt_fun1 = Me.Get_DataTable(Fun_Str)
        If Dt_fun1.Rows.Count > 0 Then
            Get_AutoIntMax = Trim(Dt_fun1.Rows(0).Item("Max_no") & "")
        Else
            Get_AutoIntMax = 0
        End If
    End Function

    ''' <summary>
    ''' 取得代碼關聯名稱
    ''' </summary>
    ''' <param name="sCode">代碼</param>
    ''' <param name="sTable">Table名稱</param>
    ''' <param name="sCodeField">代碼欄位</param>
    ''' <param name="sNameField">關聯欄位名稱</param>
    ''' <returns>關聯名稱</returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function QueryData(ByVal sCode As String, ByVal sTable As String, ByVal sCodeField As String, ByVal sNameField As String) As String
        QueryData = ""
        Dim sSql As String = ""
        Dim code_list As String() = Split(sCode, ",")
        Dim DT_tmp As New DataTable
        Dim tmpAdpt As SqlDataAdapter
        Dim SqlCmd As New SqlCommand
        Dim SqlCon As SqlConnection = New SqlConnection(Get_ConnStr(""))

        For i As Integer = 0 To code_list.Length - 1
            If code_list(i) <> "" Then
                'sSql = "SELECT " & sNameField & " FROM " & sTable & " WHERE " & sCodeField & "='" & code_list(i) & "'"
                sSql = "SELECT " & sNameField & " FROM " & sTable & " WHERE " & sCodeField & "=@code_list"
                SqlCmd = New SqlCommand(sSql, SqlCon)
                SqlCmd.Parameters.AddWithValue("@code_list", code_list(i))
                tmpAdpt = New SqlDataAdapter(SqlCmd)
                tmpAdpt.Fill(DT_tmp)
                'DT_tmp = Get_DataTable(sSql)
                If i <> 0 Then QueryData &= ","
                If DT_tmp.Rows.Count > 0 Then
                    QueryData &= DT_tmp.Rows(0).Item(sNameField).ToString
                Else
                    QueryData &= "[查無資料]"
                End If
            End If
        Next
    End Function

    ''' <summary>
    ''' 產生輔助視窗連結
    ''' </summary>
    ''' <param name="fUrl">網頁路徑</param>
    ''' <param name="sAdd_code">輔助視窗代號</param>
    ''' <param name="FieldList">欄位索引|號分隔</param>
    ''' <param name="TextBoxList">要傳回的textbox名稱以|號分隔 陣列數目要與fieldlist相同</param>
    ''' <param name="Query_where"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function Set_AddForm(ByVal fUrl As String, ByVal sAdd_code As String, ByVal FieldList As String, ByVal TextBoxList As String, ByVal Query_where As String) As String
        Dim Tmp_str As String
        Dim Tmp_script As String

        If Query_where <> "" Then
            Tmp_str = fUrl & "?add_code=" & sAdd_code & "&fieldlist=" & FieldList & "&textboxlist=" & TextBoxList & "&subwhere=" & Query_where
        Else
            Tmp_str = fUrl & "?add_code=" & sAdd_code & "&fieldlist=" & FieldList & "&textboxlist=" & TextBoxList
        End If
        Tmp_script = "subwindow=window.open('" & Tmp_str & "','AddForm','height=550,width=600,top=100,left=200,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no','');subwindow.focus();return false;"

        Set_AddForm = Tmp_script
    End Function

    ''' <summary>
    ''' 產生輔助視窗連結(CheckBoxList)
    ''' </summary>
    ''' <param name="fUrl">網頁路徑</param>
    ''' <param name="sAdd_code">輔助視窗代號</param>
    ''' <param name="TextBoxList">要傳回的TextBox(名稱|代號值)</param>
    ''' <param name="Query_where"></param>
    ''' <param name="Checked">已選取項目(代號值)</param>
    ''' <param name="ChkName">已選取項目(名稱)</param>
    ''' <param name="Column">CheckBoxList顯示欄位數</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function Set_AddList(ByVal fUrl As String, ByVal sAdd_code As String, ByVal TextBoxList As String, ByVal Query_where As String, ByVal Checked As String, ByVal ChkName As String, ByVal Column As Integer) As String
        Dim Tmp_str As String
        Tmp_str = fUrl & "?add_code=" & sAdd_code & "&textbox=" & TextBoxList
        If Query_where <> "" Then Tmp_str &= "&subwhere=" & Query_where
        If Checked <> "" Then Tmp_str &= "&checked=" & Checked
        If ChkName <> "" Then Tmp_str &= "&ChkName=" & Server.UrlEncode(ChkName)
        If Column <> 0 Then Tmp_str &= "&column=" & Column.ToString

        Set_AddList = "subwindow=window.open('" & Tmp_str & "','AddList','height=550,width=440,top=100,left=200,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no','');subwindow.focus();return false;"
    End Function

    ''' <summary>
    ''' 產生日曆視窗連結
    ''' </summary>
    ''' <param name="fUrl">網頁路徑</param>
    ''' <param name="txtID">TextBox ID</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function Set_Calendar(ByVal fUrl As String, ByVal txtID As String) As String
        Dim Tmp_str As String
        Dim Tmp_script As String

        Tmp_str = fUrl & "?TextBoxId=" & txtID
        Tmp_script = "subwindow=window.open('" & Tmp_str & "','Calendar','height=230,width=220,top=100,left=200,status=no,toolbar=no,menubar=no,location=no','');subwindow.focus();return false;"

        Set_Calendar = Tmp_script
    End Function

    ''' <summary>
    ''' 寫入使用記錄(包含操作內容)
    ''' </summary>
    ''' <param name="Cprg_memo">執行內容</param>
    ''' <param name="Cusr_type">修改類型</param>
    ''' <param name="Cprg_code">程式代號</param>
    ''' <param name="Cusr_code">使用者代號</param>
    ''' <param name="Cusr_ip">使用者IP</param>
    ''' <param name="sFieldName">欄位名稱</param>
    ''' <param name="sOldData">修改前資料</param>
    ''' <param name="sNewData">修改後資料</param>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Sub InsUsrRec(ByVal Cprg_memo As String, ByVal Cusr_type As String, ByVal Cprg_code As String, ByVal Cusr_code As String, ByVal Cusr_ip As String, ByVal sFieldName As String, ByVal sOldData As String, ByVal sNewData As String)
        Dim Sql_str As String
        Dim nDate As String
        Dim nTime As String
        Dim SqlCmd As New SqlCommand
        Dim SqlCon As SqlConnection = Me.ConnectDB

        nDate = Year(Now) & "/" & Right("00" & Month(Now), 2) & "/" & Right("00" & Day(Now), 2)
        nTime = Right("00" & Hour(Now), 2) & ":" & Right("00" & Minute(Now), 2) & ":" & Right("00" & Second(Now), 2)

        'Sql_str = "INSERT INTO BDP200 (" & _
        '          "            usr_code,prg_code,usr_ip,usr_date,usr_time," & _
        '          "            usr_type,cmemo,field_names,data_old,data_new)" & _
        '          " VALUES ('" & _
        '          RepSql(Cusr_code) & "','" & _
        '          Cprg_code & "','" & _
        '          RepSql(Cusr_ip) & "','" & _
        '          nDate & "','" & _
        '          nTime & "','" & _
        '          RepSql(Cusr_type) & "','" & _
        '          RepSql(Cprg_memo) & "',N'" & _
        '          RepSql(sFieldName) & "',N'" & _
        '          RepSql(sOldData) & "',N'" & _
        '          RepSql(sNewData) & "')"
        Sql_str = "INSERT INTO BDP200 (" & _
                  "            usr_code,prg_code,usr_ip,usr_date,usr_time," & _
                  "            usr_type,cmemo,field_names,data_old,data_new)" & _
                  " VALUES (@Cusr_code,@Cprg_code,@Cusr_ip,@nDate,@nTime," & _
                  "@Cusr_type,@Cprg_memo,@sFieldName,@sOldData,@sNewData)"
        SqlCmd = New SqlCommand(Sql_str, SqlCon)
        SqlCmd.Parameters.AddWithValue("@Cusr_code", RepSql(Cusr_code))
        SqlCmd.Parameters.AddWithValue("@Cprg_code", Cprg_code)
        SqlCmd.Parameters.AddWithValue("@Cusr_ip", RepSql(Cusr_ip))
        SqlCmd.Parameters.AddWithValue("@nDate", nDate)
        SqlCmd.Parameters.AddWithValue("@nTime", nTime)
        SqlCmd.Parameters.AddWithValue("@Cusr_type", RepSql(Cusr_type))
        SqlCmd.Parameters.AddWithValue("@Cprg_memo", RepSql(Cprg_memo))
        SqlCmd.Parameters.AddWithValue("@sFieldName", RepSql(sFieldName))
        SqlCmd.Parameters.AddWithValue("@sOldData", RepSql(sOldData))
        SqlCmd.Parameters.AddWithValue("@sNewData", RepSql(sNewData))
        Try
            Using (SqlCon)
                SqlCmd.ExecuteNonQuery()
                SqlCon.Close()
            End Using
        Catch ex As Exception

        End Try
        'SaveData(Sql_str)
    End Sub

    ''' <summary>
    ''' 字串加密、解密
    ''' </summary>
    ''' <param name="Ctext">加解密字串</param>
    ''' <param name="Ctype1">E:加密、D:解密</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <WebMethod()> _
    Public Function EncodeString(ByVal Ctext As String, ByVal Ctype1 As String) As String
        Dim Ilop As Integer
        Dim Tmp_num As Double
        Dim Base_num As Double
        Dim Ctrx() As String

        EncodeString = ""
        Tmp_num = Second(Format(Now, "hh:mm:ss"))

        If Ctext = "" Then
            '空白則不處理直接 return
            Exit Function
        End If
        Try


            Select Case UCase(Ctype1)
                Case "E" '加碼
                    For Ilop = 1 To Len(Ctext)
                        Randomize()
                        Tmp_num = Int((9 * Rnd()) + 1)
                        If Tmp_num Mod 3 = 0 Then
                            EncodeString = EncodeString & Format(Tmp_num, "0") & CStr(Asc(Mid(Ctext, Ilop, 1)) + (Val(Format(Tmp_num, "0")) * Val(Format(Len(Ctext), "0")))) & "－"
                        Else
                            If Tmp_num Mod 2 = 0 Then
                                EncodeString = EncodeString & Format(Tmp_num, "0") & CStr(Asc(Mid(Ctext, Ilop, 1)) + (Val(Format(Tmp_num, "0")) + Val(Format(Len(Ctext), "0")))) & "－"
                            Else
                                EncodeString = EncodeString & Format(Tmp_num, "0") & CStr(Asc(Mid(Ctext, Ilop, 1)) + (Math.Abs(Val(Format(Tmp_num, "0")) - Val(Format(Len(Ctext), "0"))))) & "－"
                            End If
                        End If
                    Next
                    If Len(EncodeString) Mod 3 = 0 Then
                        EncodeString = EncodeString & CStr(Int((1000 * Rnd()) + 1)) & "－" & CStr(Int((100 * Rnd()) + 1)) & "3"
                    Else
                        If Len(EncodeString) Mod 2 = 0 Then
                            EncodeString = EncodeString & CStr(Int((1000 * Rnd()) + 1)) & "－" & CStr(Int((1000 * Rnd()) + 1)) & "－" & CStr(Int((100 * Rnd()) + 1)) & "2"
                        Else
                            EncodeString = EncodeString & CStr(Int((100 * Rnd()) + 1)) & "1"
                        End If
                    End If
                Case "D" '解碼
                    If Mid(Right(Ctext, 1), 1, 1) = "3" Then Base_num = 2
                    If Mid(Right(Ctext, 1), 1, 1) = "2" Then Base_num = 3
                    If Mid(Right(Ctext, 1), 1, 1) = "1" Then Base_num = 1
                    Ctrx = Split(Ctext & "－", "－")
                    For Ilop = 0 To UBound(Ctrx) - (Base_num + 1)
                        If Val(Mid(Ctrx(Ilop), 1, 1)) Mod 3 = 0 Then
                            EncodeString = EncodeString & Chr(Val(Mid(Ctrx(Ilop), 2, 10)) - (Val(Mid(Ctrx(Ilop), 1, 1)) * (UBound(Ctrx) - Base_num)))
                        Else
                            If Val(Mid(Ctrx(Ilop), 1, 1)) Mod 2 = 0 Then
                                EncodeString = EncodeString & Chr(Val(Mid(Ctrx(Ilop), 2, 10)) - (Val(Mid(Ctrx(Ilop), 1, 1)) + (UBound(Ctrx) - Base_num)))
                            Else
                                EncodeString = EncodeString & Chr(Val(Mid(Ctrx(Ilop), 2, 10)) - (Math.Abs(Val(Mid(Ctrx(Ilop), 1, 1)) - (UBound(Ctrx) - Base_num))))
                            End If
                        End If
                    Next
            End Select

        Catch ex As Exception
            Return "身分證解碼錯誤"
        End Try
    End Function

    <WebMethod()> _
    Public Function Get_PaperHtml(ByVal sPaperCode As String) As String
        Dim sSql As String
        Dim dtTmp As New DataTable
        Dim tmpAdpt As SqlDataAdapter
        Dim SqlCmd As New SqlCommand
        Dim SqlCon As SqlConnection = New SqlConnection(Get_ConnStr(""))

        'sSql = "SELECT paper_html FROM e_paper WHERE paper_code='" & sPaperCode & "'"
        sSql = "SELECT paper_html FROM e_paper WHERE paper_code=@sPaperCode"
        SqlCmd = New SqlCommand(sSql, SqlCon)
        SqlCmd.Parameters.AddWithValue("@sPaperCode", sPaperCode)
        tmpAdpt = New SqlDataAdapter(SqlCmd)
        tmpAdpt.Fill(dtTmp)
        'dtTmp = Me.Get_DataTable(sSql)
        If dtTmp.Rows.Count > 0 Then
            Return Trim(dtTmp.Rows(0).Item(0) & "")
        End If

        Return ""
    End Function

    ''' <summary>
    ''' 將錯誤訊息寫入TXT文字檔寫
    ''' </summary>
    ''' <param name="sFunName">執行程式名稱</param>
    ''' <param name="sErrMsg">錯誤訊息字串</param>
    ''' <remarks></remarks>
    Private Sub ErrorLog(ByVal sFunName As String, ByVal sErrMsg As String)
        '將LOG匯出至TXT檔
        Dim sLogFilePath As String = Server.MapPath("~/")
        Dim sLogFileName As String = "webservice.log"
        Dim myStreamWriter As System.IO.StreamWriter = Nothing
        Try
            myStreamWriter = New System.IO.StreamWriter(sLogFilePath & sLogFileName, True, System.Text.Encoding.GetEncoding("BIG5"))
            myStreamWriter.Write(Format(Now, "yyyy/MM/dd HH:mm:ss") & " [" & sFunName & "] " & sErrMsg & vbCrLf)
        Catch ex As Exception

        Finally
            If myStreamWriter IsNot Nothing Then
                myStreamWriter.Close()
            End If
        End Try
    End Sub

    Private Function ChangeString(ByVal C_Str As String) As String
        Dim str As String = ""

        For Ilop As Integer = 1 To Len(C_Str)
            If (Ilop Mod 2) = 0 Then
                str &= ChangeChar2(Mid(C_Str, Ilop, 1))
            Else
                str &= ChangeChar4(Mid(C_Str, Ilop, 1))
            End If
        Next
        Return str
    End Function

    Private Function ChangeChar4(ByVal C_bit As String) As String
        Select Case C_bit
            Case "n" : ChangeChar4 = "!"
            Case "F" : ChangeChar4 = "#"
            Case "p" : ChangeChar4 = "$"
            Case "A" : ChangeChar4 = "%"
            Case "w" : ChangeChar4 = "&"
            Case "1" : ChangeChar4 = "("
            Case "]" : ChangeChar4 = ")"
            Case "f" : ChangeChar4 = "*"
            Case "E" : ChangeChar4 = "+"
            Case "5" : ChangeChar4 = ","
            Case "L" : ChangeChar4 = "-"
            Case "r" : ChangeChar4 = "."
            Case "T" : ChangeChar4 = "/"
            Case "i" : ChangeChar4 = "0"
            Case "H" : ChangeChar4 = "1"
            Case "<" : ChangeChar4 = "2"
            Case "@" : ChangeChar4 = "3"
            Case "{" : ChangeChar4 = "4"
            Case "c" : ChangeChar4 = "5"
            Case "B" : ChangeChar4 = "6"
            Case "k" : ChangeChar4 = "7"
            Case "^" : ChangeChar4 = "8"
            Case "y" : ChangeChar4 = "9"
            Case "W" : ChangeChar4 = ":"
            Case "+" : ChangeChar4 = ";"
            Case "6" : ChangeChar4 = "<"
            Case "u" : ChangeChar4 = "="
            Case "J" : ChangeChar4 = ">"
            Case "C" : ChangeChar4 = "?"
            Case "P" : ChangeChar4 = "@"
            Case "t" : ChangeChar4 = "A"
            Case "}" : ChangeChar4 = "B"
            Case "[" : ChangeChar4 = "C"
            Case "K" : ChangeChar4 = "D"
            Case "?" : ChangeChar4 = "E"
            Case "e" : ChangeChar4 = "F"
            Case "j" : ChangeChar4 = "G"
            Case "U" : ChangeChar4 = "H"
            Case "2" : ChangeChar4 = "I"
            Case ">" : ChangeChar4 = "J"
            Case "a" : ChangeChar4 = "K"
            Case "=" : ChangeChar4 = "L"
            Case "o" : ChangeChar4 = "M"
            Case "9" : ChangeChar4 = "N"
            Case "." : ChangeChar4 = "O"
            Case "(" : ChangeChar4 = "P"
            Case "_" : ChangeChar4 = "Q"
            Case "v" : ChangeChar4 = "R"
            Case "3" : ChangeChar4 = "S"
            Case "!" : ChangeChar4 = "T"
            Case "h" : ChangeChar4 = "U"
            Case "~" : ChangeChar4 = "V"
            Case "b" : ChangeChar4 = "W"
            Case "`" : ChangeChar4 = "X"
            Case "7" : ChangeChar4 = "Y"
            Case "*" : ChangeChar4 = "Z"
            Case "G" : ChangeChar4 = "["
            Case "l" : ChangeChar4 = "\"
            Case "R" : ChangeChar4 = "]"
            Case "m" : ChangeChar4 = "^"
            Case "s" : ChangeChar4 = "_"
            Case "0" : ChangeChar4 = "`"
            Case "-" : ChangeChar4 = "a"
            Case "\" : ChangeChar4 = "b"
            Case "Z" : ChangeChar4 = "c"
            Case ":" : ChangeChar4 = "d"
            Case "#" : ChangeChar4 = "e"
            Case "I" : ChangeChar4 = "f"
            Case "D" : ChangeChar4 = "g"
            Case "," : ChangeChar4 = "h"
            Case "N" : ChangeChar4 = "i"
            Case "O" : ChangeChar4 = "j"
            Case ";" : ChangeChar4 = "k"
            Case "%" : ChangeChar4 = "l"
            Case "Q" : ChangeChar4 = "m"
            Case "z" : ChangeChar4 = "n"
            Case "&" : ChangeChar4 = "o"
            Case "x" : ChangeChar4 = "p"
            Case "S" : ChangeChar4 = "q"
            Case ")" : ChangeChar4 = "r"
            Case "Y" : ChangeChar4 = "s"
            Case "4" : ChangeChar4 = "t"
            Case "M" : ChangeChar4 = "u"
            Case "/" : ChangeChar4 = "v"
            Case "8" : ChangeChar4 = "w"
            Case "g" : ChangeChar4 = "x"
            Case "X" : ChangeChar4 = "y"
            Case "V" : ChangeChar4 = "z"
            Case "d" : ChangeChar4 = "{"
            Case "q" : ChangeChar4 = "}"
            Case " " : ChangeChar4 = "~"
            Case "$" : ChangeChar4 = " "
            Case Else
                ChangeChar4 = C_bit
        End Select
    End Function

    Private Function ChangeChar2(ByVal C_bit As String) As String
        Select Case C_bit
            Case "a" : ChangeChar2 = "!"
            Case "1" : ChangeChar2 = "#"
            Case "m" : ChangeChar2 = "$"
            Case "H" : ChangeChar2 = "%"
            Case "8" : ChangeChar2 = "&"
            Case "g" : ChangeChar2 = "("
            Case "o" : ChangeChar2 = ")"
            Case "0" : ChangeChar2 = "*"
            Case "O" : ChangeChar2 = "+"
            Case "k" : ChangeChar2 = ","
            Case "K" : ChangeChar2 = "-"
            Case "d" : ChangeChar2 = "."
            Case "x" : ChangeChar2 = "/"
            Case "R" : ChangeChar2 = "0"
            Case "{" : ChangeChar2 = "1"
            Case "," : ChangeChar2 = "2"
            Case "<" : ChangeChar2 = "3"
            Case "Y" : ChangeChar2 = "4"
            Case "q" : ChangeChar2 = "5"
            Case "A" : ChangeChar2 = "6"
            Case "(" : ChangeChar2 = "7"
            Case "/" : ChangeChar2 = "8"
            Case "&" : ChangeChar2 = "9"
            Case "I" : ChangeChar2 = ":"
            Case "J" : ChangeChar2 = ";"
            Case "e" : ChangeChar2 = "<"
            Case "u" : ChangeChar2 = "="
            Case "6" : ChangeChar2 = ">"
            Case "w" : ChangeChar2 = "?"
            Case "D" : ChangeChar2 = "@"
            Case "#" : ChangeChar2 = "A"
            Case "c" : ChangeChar2 = "B"
            Case "j" : ChangeChar2 = "C"
            Case "!" : ChangeChar2 = "D"
            Case ">" : ChangeChar2 = "E"
            Case "y" : ChangeChar2 = "F"
            Case "T" : ChangeChar2 = "G"
            Case "~" : ChangeChar2 = "H"
            Case "i" : ChangeChar2 = "I"
            Case "F" : ChangeChar2 = "J"
            Case ")" : ChangeChar2 = "K"
            Case "b" : ChangeChar2 = "L"
            Case "h" : ChangeChar2 = "M"
            Case "%" : ChangeChar2 = "N"
            Case "-" : ChangeChar2 = "O"
            Case "t" : ChangeChar2 = "P"
            Case "s" : ChangeChar2 = "Q"
            Case "X" : ChangeChar2 = "R"
            Case "f" : ChangeChar2 = "S"
            Case "*" : ChangeChar2 = "T"
            Case "l" : ChangeChar2 = "U"
            Case "_" : ChangeChar2 = "V"
            Case "^" : ChangeChar2 = "W"
            Case "n" : ChangeChar2 = "X"
            Case "+" : ChangeChar2 = "Y"
            Case "P" : ChangeChar2 = "Z"
            Case "U" : ChangeChar2 = "["
            Case "W" : ChangeChar2 = "\"
            Case "r" : ChangeChar2 = "]"
            Case "z" : ChangeChar2 = "^"
            Case "N" : ChangeChar2 = "_"
            Case "p" : ChangeChar2 = "`"
            Case "`" : ChangeChar2 = "a"
            Case "$" : ChangeChar2 = "b"
            Case "M" : ChangeChar2 = "c"
            Case "2" : ChangeChar2 = "d"
            Case "S" : ChangeChar2 = "e"
            Case "v" : ChangeChar2 = "f"
            Case "4" : ChangeChar2 = "g"
            Case "." : ChangeChar2 = "h"
            Case ";" : ChangeChar2 = "i"
            Case " " : ChangeChar2 = "j"
            Case "9" : ChangeChar2 = "k"
            Case "C" : ChangeChar2 = "l"
            Case "B" : ChangeChar2 = "m"
            Case "}" : ChangeChar2 = "n"
            Case "\" : ChangeChar2 = "o"
            Case "G" : ChangeChar2 = "p"
            Case "V" : ChangeChar2 = "q"
            Case "E" : ChangeChar2 = "r"
            Case "L" : ChangeChar2 = "s"
            Case "3" : ChangeChar2 = "t"
            Case "?" : ChangeChar2 = "u"
            Case ":" : ChangeChar2 = "v"
            Case "@" : ChangeChar2 = "w"
            Case "5" : ChangeChar2 = "x"
            Case "]" : ChangeChar2 = "y"
            Case "7" : ChangeChar2 = "z"
            Case "[" : ChangeChar2 = "{"
            Case "=" : ChangeChar2 = "}"
            Case "Q" : ChangeChar2 = "~"
            Case "Z" : ChangeChar2 = " "
            Case Else
                ChangeChar2 = C_bit
        End Select
    End Function
End Class
