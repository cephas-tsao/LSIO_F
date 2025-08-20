<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<%
response.Buffer = true
session.Codepage =65001
response.Charset = "utf-8"
%>

<%@ Import Namespace = "System.Data" %>
<%@ Import Namespace = "System.Data.SqlClient" %>
<%@ Import Namespace = "System.Data.OleDb" %>
<%@ Import NameSpace = "System.IO"%>
<%@ Import NameSpace = "System.Web.Configuration" %>

	
<%	
Dim Con As SqlConnection = New SqlConnection
      Con.ConnectionString = WebConfigurationManager.ConnectionStrings("AppSysConnectionString").ConnectionString	  
Dim Conn As SqlConnection = New SqlConnection
      Conn.ConnectionString = WebConfigurationManager.ConnectionStrings("AppSysConnectionString").ConnectionString
		
Dim filetxt  As String	
 
Dim a1 As String = Request("a1")
Dim a2 As String = Request("a2")
Dim a3 As String = Request("a3")
Dim a4 As String = Request("a4")
Dim a5 As String = Request("a5")
Dim a6 As String = Request("a6")
Dim a7 As String = Request("a7")
Dim active As String = Request("active")
Dim ckcode As String = Request("ckcode")
Dim reckcode As String = Request("reckcode")
Dim pass As String = Request("pass")
Dim chkpass As String = Request("chkpass")
  
  
if ckcode <> reckcode then
   'response.Write ("<script>alert('驗證碼錯誤');history.back();</script>")
   'response.end
  end if
  

   
' Conn.open()
'
'Dim OU4
'Dim Y as Integer
'Dim sqlstr2  = "Select * From Ad_Member where adb_email='"& a3 &"'"
'
'	  Dim cmd2 As SQLCommand = New SQLCommand(sqlstr2 ,Conn)
'	  Dim dr2 As SQLDataReader = cmd2.ExecuteReader() 
	  
      ' if  dr2.Read()  then 
   
    dim em = new System.Net.Mail.MailMessage()
     em.From = new System.Net.Mail.MailAddress("lio.epaper2@gov.taipei", "臺北市勞動檢查處系統", System.Text.Encoding.UTF8)
     em.To.Add(new System.Net.Mail.MailAddress(""& Request("a3") &""))  '收件者
	 em.BCC.Add(new System.Net.Mail.MailAddress("jyunyulin@gmail.com"))  '收件者
     em.Subject = "塔式起重機填報作業通報-"& Request("stitle") &""   '信件主題 
     em.SubjectEncoding = System.Text.Encoding.UTF8
     em.Body = ""& Request("com_name") &"您好：<br>"& Request("textarea") &"" '內容 
     em.BodyEncoding = System.Text.Encoding.UTF8
     em.IsBodyHtml = true	  '信件內容是否使用HTML格式
     dim smtp = new System.Net.Mail.SmtpClient()
     smtp.Port = 25
	 smtp.UseDefaultCredentials = False
     'smtp.Host = "smtp.gov.taipei"  'SMTP伺服器
	 smtp.Host = "163.29.37.56"  'SMTP伺服器
     smtp.Send(em)							
   
  
' dim em = new System.Net.Mail.MailMessage()
'     em.From = new System.Net.Mail.MailAddress("lio.epaper@mail.taipei.gov.tw", "勞檢處安心電子報", System.Text.Encoding.UTF8)
'     em.To.Add(new System.Net.Mail.MailAddress(""& Request("a3") &""))  '收件者
'	 em.BCC.Add(new System.Net.Mail.MailAddress("lio.epaper@mail.taipei.gov.tw"))    
'     em.Subject = "取消勞檢處安心電子報確認"   '信件主題 
'     em.SubjectEncoding = System.Text.Encoding.UTF8
'     em.Body = "<p>"& Request("a1") &"您好：</p><p>謝謝您訂閱台北市勞檢處平安電子報，如確定 <a href='https://172.25.157.135/epaper/post3.aspx?SN="& dr2.Item("chkSN") &"'>請點我取消訂閱</a>。</p><p>敬祝：平安健康</p><p>台北市勞稽處敬上</p><p>&nbsp;</p>" 
'     em.BodyEncoding = System.Text.Encoding.UTF8
'     em.IsBodyHtml = true	  '信件內容是否使用HTML格式
'     dim smtp = new System.Net.Mail.SmtpClient()
'     smtp.Port = 25
'	 smtp.UseDefaultCredentials = False
'     'smtp.EnableSsl = true   '啟動SSL 
'     smtp.Host = "mailrelay.taipei.gov.tw"  'SMTP伺服器
'     smtp.Send(em)
 
   
  
   response.Write ("<script>alert('已通報完成');window.location.href='Default-1-2.aspx?pc_code="& Request("pc_code") &"';</script>")
	response.end
   
   'else
   
      '  response.Write ("<script>alert('查無該MAIL');history.back();</script>")
      '  response.end
       
       
      ' end if

	'' dr2.close  
'Conn.close()
   

  

	
	 
	
   
  

 %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>TEST</title>
</head>

<body>

</body>
</html>
