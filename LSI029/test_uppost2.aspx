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
'Dim webString As String
'webString = "Provider=SQLOLEDB;Data Source=.;Database=2M;User ID=test;Password=test"
	
	 Dim Con As SqlConnection = New SqlConnection
      Con.ConnectionString = WebConfigurationManager.ConnectionStrings("AppSysConnectionString").ConnectionString
	  
	   Dim Conn As SqlConnection = New SqlConnection
      Conn.ConnectionString = WebConfigurationManager.ConnectionStrings("AppSysConnectionString").ConnectionString
	

'Dim Con As OleDbConnection = new OleDbConnection(webString)	
Dim filetxt  As String	
	
 
Dim ID As String = Request("ID")
Dim IDname As String = Request("IDname")
Dim c1m As String = Request("c1m")  
Dim AID As String = Request("AID")
Dim a0 As String = Request("a0")
Dim a1 As String = Request("a1")
Dim a2 As String = Request("a2")
Dim a3 As String = Request("a3")
Dim a4 As String = Request("a4")
Dim a5 As String = Request("a5")
Dim a6 As String = Request("a6")
Dim a7 As String = Request("a7")
Dim a8 As String = Request("a8")
Dim a9 As String = Request("a9")
Dim a10 As String = Request("a10")
Dim a11 As String = Request("a11")
Dim a12 As String = Request("a12")
Dim a13 As String = Request("a13")
Dim a14 As String = Request("a14")
Dim a15 As String = Request("a15")
Dim a16 As String = Request("a16")
Dim a17 As String = Request("a17")
Dim a18 As String = Request("a18")
Dim a19 As String = Request("a19")
Dim a20 As String = Request("a20")
Dim a21 As String = Request("a21")
Dim a22 As String = Request("a22")
Dim a23 As String = Request("a23")
Dim a24 As String = Request("a24")
Dim a25 As String = Request("a25")
Dim a26 As String = Request("a26")
Dim a27 As String = Request("a27")
Dim a28 As String = Request("a28")
Dim a29 As String = Request("a29")
Dim a30 As String = Request("a30")
Dim a31 As String = Request("a31")
Dim a32 As String = Request("a32")
Dim a33 As String = Request("a33")
Dim a34 As String = Request("a34")
Dim a35 As String = Request("a35")
Dim a36 As String = Request("a36")
Dim a37 As String = Request("a37")
Dim a38 As String = Request("a38")
Dim a39 As String = Request("a39")
Dim a40 As String = Request("a40")
Dim a41 As String = Request("a41")
Dim a42 As String = Request("a42")
Dim a43 As String = Request("a43")
Dim a44 As String = Request("a44")
Dim a45 As String = Request("a45")
Dim a46 As String = Request("a46")
Dim a47 As String = Request("a47")
Dim a48 As String = Request("a48")
Dim a49 As String = Request("a49")
Dim a50 As String = Request("a50")
Dim active As String = Request("active")
Dim Dindx
Dim indx = Request("indx")
   
   Dim qa1 = Request.Form("qa1")
   Dim qa2 = Request.Form("qa2")
   Dim qa3 = Request.Form("qa3")
   Dim qa4 = Request.Form("qa4")
   Dim qa5 = Request.Form("qa5")
   Dim qa6 = Request.Form("qa6")
   Dim qa7 = Request.Form("qa7")
   Dim qa8 = Request.Form("qa8")
   Dim SN = Request.Form("SN")
   Dim d1 = Request.Form("d1")
   Dim d2 = Request.Form("d2")
   Dim d3 = Request.Form("d3")
   Dim d4 = Request.Form("d4")
   Dim d5 = Request.Form("d5")
   Dim d6 = Request.Form("d6")
   Dim d7 = Request.Form("d7")
   Dim d8 = Request.Form("d8")
   Dim d9 = Request.Form("d9")
   Dim d10 = Request.Form("d10")  
   Dim d11 = Request.Form("d11")
   Dim d12 = Request.Form("d12")
   Dim Q1txt = Request.Form("Q1txt")
   Dim Q1chk = Request.Form("Q1chk")
   
   Dim Q1txtA = Request.Form("Q1txtA")
   Dim Q2txtA = Request.Form("Q2txtA")
   Dim s2T = Request.Form("s2T")
   Dim s3T = Request.Form("s3T")
   
   
   
  
   
  ' Dim Q1txt = HttpUtility.UrlDecode(Request.Form("Q1txt"))
  
  
  if active = 11 then 
  
  
 
elseif active = 22 then	
 
 
 Con.Open()
 Dim cmd1 As new SqlCommand
 cmd1.Connection = Con
 
	cmd1.CommandText="update CheckTable set namea=@namea, valuea=@valuea ,chka=@chka where indx='"& Request("indx") &"' "
    cmd1.Parameters.Add("@namea", SqlDbType.nvarchar, 50).Value = ""& Request("name") &""
	cmd1.Parameters.Add("@valuea", SqlDbType.nvarchar, 50).Value = ""& Request("value") &""
	cmd1.Parameters.Add("@chka", SqlDbType.nvarchar, 50).Value = ""& Request("chka") &""
	cmd1.ExecuteNonQuery()
	Con.close()
Con.Dispose()
	
	response.Write ("<script>alert('更新成功');window.location.href='Default.aspx';</script>")

elseif active = 88 then	
 
 
Conn.Open()
 Dim cmd3 As new SqlCommand
 cmd3.Connection = Conn	
 
 
  cmd3.CommandText="update intoExamtest set s2T='"& s2T &"' ,s3T='"& s3T &"' ,Q1txtA='"& Q1txtA &"' ,Q2txtA='"& Q2txtA &"' where indx='"& Request("indx") &"'" 

  cmd3.ExecuteNonQuery()

	Conn.close()
    Conn.Dispose()
	
	response.Write ("<script>alert('送出成功');window.location.href='TA-list.aspx?AID="& Request("AID") &"&ID="& Request("ID") &"&indx="& Request("indx") &"';</script>")
	
	
	
elseif active = 3 then	
 
 

    Dim SNA As String = Format(Now(), "yyyyMMddhhmmssfff")            
	Dim pictxt As String = Request.Form("pictxt")
			
	    Dim savePath As String = "C:\tTTL\homework\"
        Dim FileUpload1 As HttpPostedFile= Request.Files(0)
			
            '==================================================(Start)
            Dim fileName As String = FileUpload1.FileName    '-- User上傳的檔名（不包含 Client端的路徑！）

            Dim pathToCheck As String = savePath & fileName
            Dim tempfileName As String = Nothing

            If (System.IO.File.Exists(pathToCheck)) Then
                Dim my_counter As Integer = 2
                While (System.IO.File.Exists(pathToCheck))
                    ' --檔名相同的話，目前上傳的檔名（改成 tempfileName），前面會用數字來代替。
                    tempfileName = my_counter.ToString() & "_" & fileName
                    pathToCheck = savePath & tempfileName
                    my_counter = my_counter + 1
                End While

                fileName = tempfileName
                'Label1.Text = "抱歉，您上傳的檔名發生衝突，檔名修改如下" & "<br>" & fileName
            End If

            '–完成檔案上傳的動作。		
            '-- 重點！！必須包含 Server端的「目錄」與「檔名」，才能使用 .SaveAs()方法！
		
		   if fileName <> "" then
		    
		'	fileName = SNA &".docx"
            savePath = savePath & fileName
		    FileUpload1.SaveAs(savePath)
			
			
			Conn.Open()
 Dim cmd3 As new SqlCommand
 cmd3.Connection = Conn	
 
 'if pictxt = "pictxt" then
   cmd3.CommandText="update intoExamtest set Q2file='"& fileName &"' where ID='"& Request("ID") &"' and AID='"& Request("AID") &"' and c1m='"& Request("c1m") &"' " 
 'elseif  pictxt = "pictxt1" then
  ' cmd3.CommandText="update intoExamtest set Q2file='"& fileName &"' where ID='"& Request("ID") &"' and AID='"& Request("AID") &"' and c1m='"& Request("c1m") &"' "
 'end if 
 
  cmd3.ExecuteNonQuery()

	Conn.close()
    Conn.Dispose()
	
	response.Write ("<script>alert('送出成功');window.location.href='testend.aspx?AID="& Request("AID") &"&ID="& Request("ID") &"&c1m="& Request("c1m") &"';</script>")
	

	
	else
	
	response.Write ("<script>alert('送出成功');window.location.href='testend.aspx?AID="& Request("AID") &"&ID="& Request("ID") &"&c1m="& Request("c1m") &"';</script>")
			
			
     end if	
 
 
 
	 
	

elseif active = 6 then	
 
 
 Con.Open()
 Dim cmd1 As new SqlCommand
 cmd1.Connection = Con
 
	cmd1.CommandText="update treasure set a2 =@a2 where indxTB='"& Request("indx") &"' "
    cmd1.Parameters.Add("@a2", SqlDbType.nvarchar, 500).Value = ""& Request("a2") &""
	response.Write ("<script>alert('已存檔');window.location.href='treasure-up.aspx?indx="& Request("indx") &"';</script>")
	 
	
    cmd1.ExecuteNonQuery()
	
Con.close()
Con.Dispose()

elseif active = 7 then	
 
 
 Con.Open()
 Dim cmd1 As new SqlCommand
 cmd1.Connection = Con
 
	cmd1.CommandText="update treasure set a3=@a3,a4=@a4 where indxTB='"& Request("indx") &"' "
    cmd1.Parameters.Add("@a3", SqlDbType.nvarchar, 500).Value = ""& Request("a3") &""
	cmd1.Parameters.Add("@a4", SqlDbType.nvarchar, 500).Value = ""& Request("a4") &""
	
	response.Write ("<script>alert('已存檔');window.location.href='treasure-up.aspx?indx="& Request("indx") &"';</script>")
	 
	
    cmd1.ExecuteNonQuery()
	
Con.close()
Con.Dispose()

elseif active = 8 then	
 
 
 Con.Open()
 Dim cmd1 As new SqlCommand
 cmd1.Connection = Con
 
	cmd1.CommandText="update treasure_chker set a3 =@a3 where indx='"& Request("Dindx") &"' "
    cmd1.Parameters.Add("@a3", SqlDbType.nvarchar, 500).Value = ""& Request("a3") &""
	response.Write ("<script>alert('已修改');window.location.href='treasure-up.aspx?indx="& Request("indx") &"';</script>")
	 
    cmd1.ExecuteNonQuery()
	
Con.close()
Con.Dispose()	
	
	
	elseif active = 9 then	
 
 
 Con.Open()
 Dim cmd1 As new SqlCommand
 cmd1.Connection = Con
 
	cmd1.CommandText="update accounting set d2 =@d2,d3 =@d3,d5 =@d5,d6 =@d6,d7 =@d7,d8 =@d8,d10 =@d10,d11 =@d11 where indx='"& Request("indx") &"' "
    cmd1.Parameters.Add("@d2", SqlDbType.nvarchar, 500).Value = ""& Request("d2") &""
	cmd1.Parameters.Add("@d3", SqlDbType.nvarchar, 500).Value = ""& Request("d3") &""
	cmd1.Parameters.Add("@d5", SqlDbType.nvarchar, 500).Value = ""& Request("d5") &""
	cmd1.Parameters.Add("@d6", SqlDbType.nvarchar, 500).Value = ""& Request("d6") &""
	cmd1.Parameters.Add("@d7", SqlDbType.nvarchar, 500).Value = ""& Request("d7") &""
	cmd1.Parameters.Add("@d8", SqlDbType.nvarchar, 500).Value = ""& Request("d8") &""
	cmd1.Parameters.Add("@d10", SqlDbType.nvarchar, 500).Value = ""& Request("d10") &""
	cmd1.Parameters.Add("@d11", SqlDbType.nvarchar, 500).Value = ""& Request("d11") &""
	
	
	
	
	response.Write ("<script>alert('已修改');window.location.href='5-3-web.aspx?indx="& Request("indx") &"';</script>")
	 
    cmd1.ExecuteNonQuery()
	
Con.close()
Con.Dispose()
	
 end if  
   
   
   
   
	response.Write ("<script>alert('修改成功');window.location.href='inter.aspx?a"& indx &"';</script>")
	
	response.end

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
