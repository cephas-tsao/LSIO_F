<%@ Page Language="VB" ContentType="application/json" ResponseEncoding="utf-8" %><%
response.Buffer = true
session.Codepage =65001
response.Charset = "utf-8"
%><%@ Import Namespace = "System.Data" %>
<%@ Import Namespace = "System.Data.SqlClient" %>
<%@ Import NameSpace = "System.IO"%>
<%@ Import NameSpace = "System.Web.Configuration" %>
<%
Dim Conns As SqlConnection = New SqlConnection
Conns.ConnectionString = WebConfigurationManager.ConnectionStrings("AppSysConnectionString").ConnectionString
%>{"tableNU": "A211022","tableName":"輕質屋頂與施工架及吊籠作業通報1","contents":[<%
 dim R = 0
   Dim cmd6 As SqlCommand
	  
	  ' cmd6 = New SqlCommand("select * from CheckTable  where tableNU='"& Request("tableNU") &"' order by indx asc", Conns)
	cmd6 = New SqlCommand("select * from CheckTable  where tableNU='"& Request("tableNU") &"' order by indx asc", Conns)
	   Conns.Open()
	   Dim dr6 As SQLDataReader = cmd6.ExecuteReader() 
	 While dr6.Read() 
%>{"name": "<%=dr6.Item("namea")%>", "value": "<%=dr6.Item("valuea")%>" ,"chka":"<%=dr6.Item("chka")%>"},<%
     End While
  dr6.close()
  Conns.close()
%>{"name": "endd", "value": "00","chka":"true"}]}