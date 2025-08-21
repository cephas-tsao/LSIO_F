Imports System.Data
Imports System.Xml
Imports System.Text.RegularExpressions
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.IO
Imports System.Web.Configuration

Partial Class login_check
    '繼承pagebase
    Inherits PageBase

	
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
	    '強制使用COKKIE 安全性
		
		
		If Session("confirm") Is Nothing Then
			'mErr &= "驗證碼無法確認!\n"
			'Response.Redirect("Default.aspx?mes_str=驗證碼無法確認!")
			response.Write ("<script>alert('驗證碼無法確認!');history.back();</script>")
	     response.end
		ElseIf Request("tb_confirm") <> Session("confirm").ToString() Then
	response.Write ("<script>alert('驗證碼輸入錯誤!');history.back();</script>")
	     response.end		
	'Response.Redirect("Default.aspx?mes_str=驗證碼輸入錯誤!")
			'mErr &= "驗證碼輸入錯誤!\n"
		End If
		
		
        For Each sCookie As String In Response.Cookies
		Response.Cookies(sCookie).HttpOnly = True
        Response.Cookies(sCookie).Secure = True
            
        Next
	
	
	
        '開放巨匠登入
        ' If Request("EncodeStartNet") = "QAZXSW" Then
            ' Session("usr_code") = "GEESTOTHER"
            ' Session("usr_name") = "巨匠遷入"
            ' Session("login_time") = Format(Now(), "HH:mm")
            ' Session("grp_code") = "G004"
            ' Response.Redirect("blank.aspx")
            ' Exit Sub
        ' End If

        Dim ticket As String = ""
        Try
            ticket = Request("EncodedAPREQB64")
        Catch ex As Exception
        End Try

        '若有接收到EncodedAPREQB64參數則使用SSO驗證
        If String.IsNullOrEmpty(ticket) = False Then
            Dim SSO_ws As New SSO_Service.ap
            Dim Xml As XmlNode = SSO_ws.VerifyTicket("LSIOServiceTest", ticket)
            'Dim SSO_ws As New SSOWebService.uIAMSoapClient
            'Dim Xml As XmlNode = SSO_ws.VerifyTicket("LSIOServiceTest", "45tlgj)r", ticket)

            If Xml.HasChildNodes Then
                Dim result As String = ""
                Dim apid As String = ""
                Dim ip As String = ""
                Dim description As String = ""
                Dim userDN As String = ""
                Dim sAMAccountName As String = ""
                Dim givenName As String = ""
                Dim userPrincipalName As String = ""
                Dim IDN As String = ""
                Dim orgID As String = ""
                Dim depID As String = ""

                For i As Integer = 0 To Xml.ChildNodes.Count - 1
                    Select Case Xml.ChildNodes(i).Name
                        Case "result"
                            result = Trim(Xml.ChildNodes(i).InnerText)
                        Case "apid"
                            apid = Trim(Xml.ChildNodes(i).InnerText)
                        Case "ip"
                            ip = Trim(Xml.ChildNodes(i).InnerText)
                        Case "description"
                            description = Trim(Xml.ChildNodes(i).InnerText)
                        Case "userDN"
                            userDN = Trim(Xml.ChildNodes(i).InnerText)
                        Case "sAMAccountName"
                            sAMAccountName = Trim(Xml.ChildNodes(i).InnerText)
                        Case "givenName"
                            givenName = Trim(Xml.ChildNodes(i).InnerText)
                        Case "userPrincipalName"
                            userPrincipalName = Trim(Xml.ChildNodes(i).InnerText)
                        Case "IDN"
                            IDN = Trim(Xml.ChildNodes(i).InnerText)
                        Case "orgID"
                            orgID = Trim(Xml.ChildNodes(i).InnerText)
                        Case "depID"
                            depID = Trim(Xml.ChildNodes(i).InnerText)
                    End Select
                Next
                'Response.Write("result:" & result & "<br>")
                'Response.Write("apid:" & apid & "<br>")
                'Response.Write("ip:" & ip & "<br>")
                'Response.Write("description:" & description & "<br>")
                'Response.Write("userDN:" & userDN & "<br>")
                'Response.Write("sAMAccountName:" & sAMAccountName & "<br>")
                'Response.Write("givenName:" & givenName & "<br>")
                'Response.Write("userPrincipalName:" & userPrincipalName & "<br>")
                'Response.Write("IDN:" & IDN & "<br>")
                'Response.Write("orgID:" & orgID & "<br>")
                'Response.Write("depID:" & depID & "<br>")
                'Dim MailList As String() = Split(userPrincipalName, "@")
                'Dim pass As String = Right(IDN, 4)
                'Response.Write("usercode:" & MailList(0).ToUpper & "<br>")
                'Response.Write("pass:" & pass & "<br>")

                '若result不等於true則跳回Default.aspx
                If result <> "true" Or String.IsNullOrEmpty(userPrincipalName) Then
                    'Response.Redirect("Default.aspx?mes_str=Login false")
					Response.Redirect("Default.aspx")
                Else

                    Dim Dt_db As DataTable = Get_DataSet("Login", "Mail", userPrincipalName).Tables(0)

                    '資料庫無資料須新建
                    If Dt_db.Rows.Count <= 0 Then
                        Dim ws As New WebReference.LSIO_WebService
                        Dim MailList As String() = Split(userPrincipalName, "@")
                        Dim pass As String = "a1234567@"  'Right(IDN, 4)

                        If SaveEditData("BDP080A", "ADD", MailList(0).ToUpper, givenName, ws.EncodeString(pass, "E"), "G003", _
                                     "A", userPrincipalName, "Y", "SSO建立帳號", "A001", "", ws.EncodeString(IDN, "E")) Then

                            Session("usr_code") = MailList(0).ToUpper
                            Session("usr_name") = givenName
                            Session("login_time") = Format(Now(), "HH:mm")
                            Session("grp_code") = "G003"

                            Response.Redirect("blank.aspx")
                        Else
                            'Response.Redirect("Default.aspx?mes_str=帳號建立失敗")
							Response.Redirect("Default.aspx")
                        End If
                    Else
                        '停用帳號回到登入頁
                        If Trim(Dt_db.Rows(0).Item("is_use") & "") <> "Y" Then Response.Redirect("Default.aspx")
						'Response.Redirect("Default.aspx?mes_str=此帳號停用")

                        Session("usr_code") = Trim(Dt_db.Rows(0).Item("usr_code") & "")
                        Session("usr_name") = Trim(Dt_db.Rows(0).Item("usr_name") & "")
                        Session("login_time") = Format(Now(), "HH:mm")
                        Session("grp_code") = Trim(Dt_db.Rows(0).Item("grp_code") & "")

                        Response.Redirect("blank.aspx")
                    End If
                End If
            Else
                Response.Redirect("Default.aspx")
            End If

        Else '若沒接收到參數則使用一般登入
            '定義變數
            Dim ws As New WebReference.LSIO_WebService
            Dim i_usrcode As String
            Dim i_usrpass As String
            Dim Dt_db As DataTable
            Dim Tmp_str As String

            '建立snfun1物件
            'Dim Other_obj As Object
            'Other_obj = Server.CreateObject("SnFun1.OtherFun")

            '接收參數
            i_usrcode = "" & Request("usr_code")
            i_usrpass = "" & Request("usr_pass")

            '避免 Sql injection
            i_usrcode = Replace(i_usrcode, "'", "")
            i_usrpass = Replace(i_usrpass, "'", "")

            '登入檢查
			if ( ( i_usrcode ="" )  or  ( i_usrpass ="" )      )then
			Response.Redirect("Default.aspx")
			end if 
			
			if  NOt( Regex.Match( i_usrcode , "[^0-9a-zA-z.-{7}+]").Success ) then
			
			else
			 Response.Redirect("Default.aspx")

		    end if 
			
			
			
			if NOt( Regex.Match(i_usrpass , "[^0-9a-zA-z!@#${8}+]").Success ) then
			
			else
			 Response.Redirect("Default.aspx")

		    end if 

			
			'登入檢查
            If UCase(i_usrcode) = "WOHER" Then
                'supervisor的檢查
                Tmp_str = CStr(Month(Now)) + CStr(Day(Now))
                If i_usrpass = "julia" & Tmp_str Then
                    Session("usr_code") = "WOHER"
                    Session("usr_name") = "supervisor"
                    Session("login_time") = Format(Now(), "HH:mm")
                    'Session("Dpname") = ""
                    'Session("Updpno") = ""
                    'Session("dpnos") = ""
                    Session("grp_code") = "G001"
                    Response.Redirect("blank.aspx")
                Else
                    Session("usr_code") = ""
                    Session("usr_name") = ""
                    Session("login_time") = ""
                    Session("grp_code") = ""
                    'Session("Dpname") = ""
                    'Session("dpnos") = ""
                    'Session("Updpno") = ""
                    'Response.Redirect("Default.aspx?mes_str=Login error")
					Response.Redirect("Default.aspx")
                End If
            Else
                Dt_db = Get_DataSet("Login", "Check", i_usrcode).Tables(0)
                If Dt_db.Rows.Count = 0 Then
                    Session("usr_code") = ""
                    Session("usr_name") = ""
                    Session("login_time") = ""
                    Session("grp_code") = ""
                    'Session("Dpname") = ""
                    'Session("dpnos") = ""
                    'Session("Updpno") = ""
                   ' Response.Redirect("Default.aspx?mes_str=無效的帳號名稱")
				   Response.Redirect("Default.aspx")
                    Exit Sub
                End If

                Dim usrpassE As String = ws.EncodeString(Dt_db.Rows(0).Item("usr_pass") & "", "D")

                If i_usrpass = usrpassE Then

                    'If Dt_db.Rows.Count > 0 Then
                    Session("usr_code") = UCase(i_usrcode)
                    Session("usr_name") = Dt_db.Rows(0).Item("usr_name")
                    Session("login_time") = Format(Now(), "HH:mm")
                    Session("grp_code") = Dt_db.Rows(0).Item("grp_code")
                    'Session("usr_type") = Dt_db.Rows(0).item("usr_type")
                    'Session("dpnos") = Dt_db.Rows(0).item("dpnos")
                    'If IsDBNull(Dt_db.rows(0).item("Dpname")) Then
                    '    Session("Dpname") = ""
                    '    Session("Updpno") = ""
                    'Else
                    '    Session("Dpname") = Dt_db.rows(0).item("Dpname") 'add by ann
                    '    If Trim(Dt_db.rows(0).item("Updpno")) = "" Then
                    '        Session("Updpno") = Dt_db.rows(0).item("dpno")
                    '    Else
                    '        Session("Updpno") = Dt_db.rows(0).item("Updpno")
                    '    End If
                    'End If

                    Dt_db = Get_DataSet("Login", "ch_date", Session("usr_code")).Tables(0)
                    If Dt_db.Rows.Count > 0 Then
                        '偵測提醒修改密碼天數
                        'Dt_db.rows(0).itme("usr_date")
                        Dim date_str As String = DateDiff("d", Dt_db.Rows(0).Item("usr_date"), Format(Now, "yyyy/MM/dd"))
                        If Convert.ToDouble(date_str) > Convert.ToDouble(ws.Get_SysPara("ch_date", "30")) Then
                            Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('離上次修改密碼已超過" & date_str & "日');location.href='blank.aspx';" & "</" & "Script>"))
                        Else
                            Response.Redirect("blank.aspx")
                        End If
                    Else
                        Me.Page.Form.Controls.Add(New LiteralControl("<script>alert('您尚未修改過密碼，請記得修改密碼');location.href='blank.aspx';" & "</" & "Script>"))
                    End If
                    '若停止javascript則無法自動跳頁至blank.aspx，因此加入下列強制轉換頁面
                    '若需要呈現跳出訊息則需註解掉此行程式 by schneider
                    Response.Redirect("blank.aspx")
                Else
	
	
	Dim Conn As SqlConnection = New SqlConnection
      Conn.ConnectionString = WebConfigurationManager.ConnectionStrings("AppSysConnectionString").ConnectionString	
Dim Conn1 As SqlConnection = New SqlConnection
      Conn1.ConnectionString = WebConfigurationManager.ConnectionStrings("AppSysConnectionString").ConnectionString	
											

								
Conn1.open()

Dim sqlstr1  = "select * From BDP080 where usr_code='"& Request("usr_code") &"'"
	Dim cmd1 As SQLCommand = New SQLCommand(sqlstr1 ,Conn1)
	Dim dr1 As SQLDataReader = cmd1.ExecuteReader() 
	  Dim CC
      if  dr1.Read()  then 
	   Dim usr_name = dr1.Item("usr_name") 
	   Dim login_time =Format(Now(), "yyyy-MM-dd HH:mm:ss")
		
		if dr1.Item("count") >= 6 then
			 CC = 0
		else 
			 CC = dr1.Item("count")+1
		end if
		
		Conn.open()
		Dim sqlstr2  = "update BDP080 set count='"& CC &"' where usr_code='"& Request("usr_code") &"'"
	Dim cmd2 As SQLCommand = New SQLCommand(sqlstr2 ,Conn)
	Dim dr2 As SQLDataReader = cmd2.ExecuteReader() 
	  
    if  dr2.Read()  then 
	'response.Write (dr2.Item("usr_code"))
	end if
		conn.close()
		
		
		if dr1.Item("count") = 5 then
		
			dim em = new System.Net.Mail.MailMessage()
     em.From = new System.Net.Mail.MailAddress("lio.epaperWaring@mail.taipei.gov.tw", "臺北市勞動檢查處系統", System.Text.Encoding.UTF8)
    em.To.Add(new System.Net.Mail.MailAddress(""))  '收件者
	em.To.Add(new System.Net.Mail.MailAddress("lio.119@mail.taipei.gov.tw"))  '收件者
			em.Subject = "帳號異動登入通知-"& usr_name &"-"& Request("usr_code") &""   '信件主題 
     em.SubjectEncoding = System.Text.Encoding.UTF8
			em.Body = "管理員您好：<br>目前系統"& usr_name &"-"& Request("usr_code") &"，登入異常超過5次，請留意資訊安全措施。" '內容 
    em.BodyEncoding = System.Text.Encoding.UTF8
em.IsBodyHtml = true	  '信件內容是否使用HTML格式
dim smtp = new System.Net.Mail.SmtpClient()
smtp.Port = 25
 smtp.UseDefaultCredentials = False 
'smtp.Host = "smtp.gov.taipei"  'SMTP伺服器
smtp.Host = "163.29.37.56"  'SMTP伺服器
smtp.Send(em)	
		
		end if
		response.Write ("<script>alert('第"& CC &"次-帳密錯誤');history.back();</script>")
		response.end
			
	        ' Session("usr_code") = dr1.Item("usr_code")
             'Session("usr_name") = dr1.Item("usr_name")
            ' Session("login_time") = Format(Now(), "HH:mm")
            ' Session("grp_code") = dr1.Item("grp_code")
			
			
		end if
					
							
'dim em = new System.Net.Mail.MailMessage()
'     em.From = new System.Net.Mail.MailAddress("ken@sc-top.org.tw", "台北市勞動檢查處系統", System.Text.Encoding.UTF8)
'     em.To.Add(new System.Net.Mail.MailAddress(""& dr1.Item("usr_mail") &""))  '收件者
'									em.Subject = "會議異動通知-"& dr2.Item("mrd_date") &""& dr2.Item("mrd_name") &""   '信件主題 
'     em.SubjectEncoding = System.Text.Encoding.UTF8
'     em.Body = ""& Request("com_name") &"您好：<br>您登記的會議"& dr2.Item("mrd_date") &"「"& dr2.Item("mrd_name") &"」有異動，請至系統查詢" '內容 
'     em.BodyEncoding = System.Text.Encoding.UTF8
'     em.IsBodyHtml = true	  '信件內容是否使用HTML格式
'     dim smtp = new System.Net.Mail.SmtpClient()
'     smtp.Credentials = new System.Net.NetworkCredential("ken", "sctop1qaz2wsx")
'     smtp.Port = 25
'     smtp.Host = "mail.sc-top.org.tw"  'SMTP伺服器
'     smtp.Send(em)		
'							
'	Show_Message("已將修改資訊寄給通知原登記者知悉-"& dr1.Item("usr_mail") &"")
	
	
	Conn1.close()			
				

	
                    Session("usr_code") = ""
                    Session("usr_name") = ""
                    Session("login_time") = ""
                    Session("grp_code") = ""
                    'Session("Dpname") = ""
                    'Session("dpnos") = ""

	
                    Response.Redirect("Default.aspx?mes_str=輸入帳號/密碼錯誤")
					Response.Redirect("Default.aspx")
	
	
                End If
            End If
            ws.Dispose()
        End If
    End Sub
End Class
