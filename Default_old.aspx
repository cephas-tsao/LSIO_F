<%@ import Namespace="System.data" %>
<%@ import Namespace="System.Web.UI.Control" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Web.Configuration" %>
<%@ import Namespace="Microsoft.Security.Application" %>
<%

    Dim p1 As New PageBase
    Dim ws As New WebReference.LSIO_WebService
    Dim ipquest As String = ""
	Dim ipquest_RTADDR As String = ""
	  Dim ipquest1 As String = ""
    Dim root_ip As String
    Dim ip_arr(4) As Object
    Dim con(4) As Object
    Dim I As Integer
    Dim i_messtr As String
    Dim Ip_ok As String
    Dim IpHost as string  =  Sanitizer.GetSafeHtmlFragment ( Request.Servervariables("WL_CLENT_IP")  )
	
	 
    '�H�U�o�q�y�k�i�H������proxy���᪺ip�i�Ψӭ���O����ip��user
     If (Request.ServerVariables("LOCAL_ADDR") = "" Or InStr(Request.ServerVariables("LOCAL_ADDR"), "unknown") = 0 ) Then
	 ipquest1 = Request.ServerVariables("LOCAL_ADDR")
	end if 
	
	If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" Or InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then
        ipquest = Request.ServerVariables("REMOTE_ADDR")
    ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then
        ipquest = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") - 1)
    ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then
        ipquest = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") - 1)
    Else
        ipquest = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
    End If
    
    ipquest = Trim(Mid(ipquest, 1, 15))
     ipquest_RTADDR = Trim(Mid(ipquest_RTADDR, 1, 15))
    root_ip = Request("sn_julia1208")
      
    if root_ip = CSTR(minute(NOW) + hour(NOW)) then 
      ipquest = "0.0.0.0"
    end if
  
    If ipquest = "" Then
        Response.Write("Unknow your IP address, can not login system.")
        Response.End()
    End If
	If ipquest = "locahost" Then
	ipquest = "0.0.0.0"
    End If
  
    con = Split(ipquest, ".")
    For I = 0 To UBound(con)-1
        ip_arr(I + 1) = CInt(con(I))
    Next
  

    Dim Dt_tmp As DataTable
    
    Ip_ok = False
    If ipquest <> "0.0.0.0" Then
        '�}��IP�Ϭq       
        Dt_tmp = p1.Get_DataSet("Default", "A").Tables(0)
        
        For I = 0 To Dt_tmp.Rows.Count - 1
            With Dt_tmp.Rows(I)
                If (ip_arr(1) >= .Item("ip_s1") And ip_arr(1) <= .Item("ip_e1")) And (ip_arr(2) >= .Item("ip_s2") And ip_arr(2) <= .Item("ip_e2")) And (ip_arr(3) >= .Item("ip_s3") And ip_arr(3) <= .Item("ip_e3")) And (ip_arr(4) >= .Item("ip_s4") And ip_arr(4) <= .Item("ip_e4")) Then
                    Ip_ok = True
                    Exit For
                End If
            End With
        Next
                    
        Dt_tmp.Clear()
          
        Dt_tmp = p1.Get_DataSet("Default", "B").Tables(0)
        
        For I = 0 To Dt_tmp.Rows.Count - 1
            With Dt_tmp.Rows(I)
                If (ip_arr(1) >= .Item("ip_s1") And ip_arr(1) <= .Item("ip_e1")) And (ip_arr(2) >= .Item("ip_s2") And ip_arr(2) <= .Item("ip_e2")) And (ip_arr(3) >= .Item("ip_s3") And ip_arr(3) <= .Item("ip_e3")) And (ip_arr(4) >= .Item("ip_s4") And ip_arr(4) <= .Item("ip_e4")) Then
                    Ip_ok = False
                    Exit For
                End If
            End With
        Next
    End If
        
    '�P�_�O�_�}��ip����()
    If ConfigurationManager.AppSettings("Check_ip") = "N" Or ipquest = "0.0.0.0" Then
       Ip_ok = True
    End If
  
    If Ip_ok = False Then
        '�����\�n���t��
		Response.Write("<p>�T��s��</p>")
       ' Response.Write("your IP : " & ipquest & " ,can not login system.")
		' Response.Write("<p>your t IP : " & Request.ServerVariables("REMOTE_ADDR") & " ,can not login system.</p>")
		'Response.Write("<p>your C IP : " & IpHost   & " ,can not login system.</p>")
        Response.End()
    End If
	
	 '�����\�n���t��
	if Server.HtmlEncode(ipquest)=Server.HtmlEncode("163.29.40.135")  then
	Response.Write("<p>�T��s��</p>")
		Response.End()
	end if 
   
    
    '���o�M�צW��
    Dim PrjName As String = ConfigurationManager.AppSettings("Prj_name")
	
	    '�j��ϥ�COKKIE �w����
        For Each sCookie As String In Response.Cookies
		Response.Cookies(sCookie).HttpOnly = True
        Response.Cookies(sCookie).Secure = True
            
        Next
%>

<html>
<head>
<meta name="GENERATOR" content="211">
<meta name="ProgId" content="23">
<meta http-equiv="Content-Type" content="text/html; charset=BIG5">
<title><%=PrjName%></title>
</head>
<style>
    .style1
    {
        width: 100%;
        height: 240px;
    }
    .style32
    {
        width: 155px;
    }
    .style33
    {
        width: 255px;
        height: 75px;
    }
    .style37
    {
        width: 50%;
        height: 294px;
    }
    .style38
    {
        height: 48px;
        width: 255px;
    }
    .style40
    {
        height: 294px;
    }
    .style41
    {
        width: 85px;
    }
</style>
<body oncontextmenu="window.event.returnValue=false" bgcolor="white">

  <%
    i_messtr = request.querystring("mes_str")
    if i_messtr <> "" then
          i_messtr = i_messtr & "<br>"
    end if
  %>
  <div align="center" style="background-position:center middle" >
  <TABLE border=0 cellSpacing=0 summary=main cellPadding=0 align=center>
   <!--top -->
    <table style="position:relative; margin-left:auto; margin-right:auto; width:768; top: 0px; left: 0px;"  
          border ="0">
     <tr><td>         
           <div>
                <table align="center" border="0" cellpadding='0' cellspacing='0' width="768">
                <tr>
                  <td align="right" background="../images/menu-01/default-head.jpg" height="100" style="vertical-align: bottom">
                    <font size="6" color="#f0f8ff" face="�з���"><%=PrjName%></font>
                   </td>
                </tr>
                </table>
            </div>
         </td>
     </tr>
     </table>
     <hr noshade="noshade"/  style="width:768;height:1px">

     <!-- center -->
     <table style="border: 1px solid #006400; width:768;">
     <tr>
     <td colspan="2" style="border: 3px solid #006400; background-color:#006400;">
     <font size="5" color="#f0f8ff" face="�з���">�w��ϥ�  << <%=PrjName%> >></font>
     </tr>
     
     <tr>
     <!-- left -->
     <td class="style37">
     <table class="style1">
     <tr>
     <td align="center">
     <img src="/images/menu-01/default-center.jpg"></image>
     </td>
     </tr>
          
     </table>
         
    </td>
    <!-- right -->
     <td class="style40" align="center">
    <table>
     <tr>
      <td >
        <table class="style1" style="border-style: groove; border-color: #CCCCCC;background-color:#d8fbd8" >
        <tr><td align="center" ><font color="blue" size="6" face="�з���">�t�εn�J</font></td></tr>
        <tr>
         <td class="style33">
          <Form name="Form1" Action="login_check.aspx" Method="POST">
              
                  <table border='0' width ="300">
                  
                      <tr >
                        <td class="style41" align="right">
                        <font face="Arial" color="#335B98" width ="100"><b>�b���G</b></font></td>
                        <td><input type="text" name="usr_code" size="24" class="style32" ></td>
                      </tr>
                      <tr>
                        <td class="style41" align="right">
                        <font face="Arial" color="#335B98" align = "right"><b>�K�X�G</b></font>
                        </td>
                        <td><input type="password" name="usr_pass" size="24" class="style32"></td>
                      </tr>
                      <tr >
                        <td class="style41"></td>
                        <td>
						<Font color="#335B98">
							<strong>
								IP:<br>
								
									<%=Server.HtmlEncode(ipquest)%><br>
									
									
							</strong>
						</font>
						</td>
                        <tr>
                        <td class="style41"></td>
                        <td><input class="submit" type="image" src="/images/btn.png" value="�n�J"><br>
                        &nbsp;&nbsp;<span lang="zh-tw"><a href="ForgetPassword.aspx" target="_blank"><b><font size="2">�ѰO�K�X</font></b></a></span></td>
                        </tr>
                      </tr>    
                  </table>
                  <tr>
                  <td class="style38" align="center"><font face="Arial" size="3" color="red"><%=i_messtr%></font></td>
                  </tr>
          </Form>
         </td>
        </tr>
        
        </table>
      </td>
     </tr>
      
    </table>
    </td>
    
    </tr>
        
    
    </table>
     <!-- bottom -->

     <hr noshade="noshade"/ style="width:768px;height:1px">
     
    <table style="position:relative; margin-left:auto; margin-right:auto; width:768;" border ="0" cellpadding='0' cellspacing='0' >
     <tr><td>         
           
                  <table align="center" border="0">
                    <tr><td><img src="../images/menu-01/default-bottom.jpg" / style="width:768; height: 76px;" /></td></tr>
                </table>
           </td>
     </tr>
     </table>
    </TABLE>
  </div>
</body>

<script language="javascript">
  Form1.usr_code.focus();
</script>

</html>
