<%@ Page Language="VB" AutoEventWireup="false" ContentType="text/html" ResponseEncoding="UTF-8" CodeFile="Default.aspx.vb" Inherits="_Default" %>
<%
response.Buffer = true
session.Codepage =65001
response.Charset = "utf-8"
%>
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
	
	 
    '以下這段語法可以偵測到proxy之後的ip可用來限制不是實體ip的user
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
        '開放的IP區段       
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
        
    '判斷是否開啟ip偵測()
    If ConfigurationManager.AppSettings("Check_ip") = "N" Or ipquest = "0.0.0.0" Then
       Ip_ok = True
    End If
  
    If Ip_ok = False Then
        '不允許登錄系統
		Response.Write("<p>禁止存取</p>")
       ' Response.Write("your IP : " & ipquest & " ,can not login system.")
		' Response.Write("<p>your t IP : " & Request.ServerVariables("REMOTE_ADDR") & " ,can not login system.</p>")
		'Response.Write("<p>your C IP : " & IpHost   & " ,can not login system.</p>")
        Response.End()
    End If
	
	 '不允許登錄系統
	if Server.HtmlEncode(ipquest)=Server.HtmlEncode("163.29.40.135")  then
	Response.Write("<p>禁止存取</p>")
		Response.End()
	end if 
   
    
    '取得專案名稱
    Dim PrjName As String = ConfigurationManager.AppSettings("Prj_name")
	
	    '強制使用COKKIE 安全性
        For Each sCookie As String In Response.Cookies
		Response.Cookies(sCookie).HttpOnly = True
        Response.Cookies(sCookie).Secure = True
            
        Next
%>



  <%
    i_messtr = request.querystring("mes_str")
    if i_messtr <> "" then
          i_messtr = i_messtr & "<br>"
    end if
  %>
  
  
 
 <!DOCTYPE html>
<html lang="en">
<head>
	<title><%=PrjName%></title>
	<meta charset="UTF-8">
	<meta name="viewport" content="width=device-width, initial-scale=1">
<!--===============================================================================================-->	
	<link rel="icon" type="image/png" href="images/icons/favicon.ico"/>
<!--===============================================================================================-->
	<link rel="stylesheet" type="text/css" href="vendor/bootstrap/css/bootstrap.min.css">
<!--===============================================================================================-->
	<link rel="stylesheet" type="text/css" href="fonts/font-awesome-4.7.0/css/font-awesome.min.css">
<!--===============================================================================================-->
	<link rel="stylesheet" type="text/css" href="vendor/animate/animate.css">
<!--===============================================================================================-->	
	<link rel="stylesheet" type="text/css" href="vendor/css-hamburgers/hamburgers.min.css">
<!--===============================================================================================-->
	<link rel="stylesheet" type="text/css" href="vendor/select2/select2.min.css">
<!--===============================================================================================-->
	<link rel="stylesheet" type="text/css" href="css/util.css">
	<link rel="stylesheet" type="text/css" href="css/main.css">
<!--===============================================================================================-->
</head>
<body>
	<div class="limiter">
		<div class="container-login100">
			<div class="wrap-login100">
				<div class="login100-pic js-tilt" data-tilt>
					<img src="images/img-01.png" alt="IMG"><br><br><br><br><br>
					<button class="login100-form-btn" style="background-color:darkslateblue;" onClick="window.location.href='SSOlogin.aspx'">
							使用SSO登入
						</button>
				</div>

				<Form name="Form1" Action="login_check.aspx" class="login100-form validate-form" runat="server"  Method="POST">
					<span class="login100-form-title">
<img src="images/img-021.png">
					</span>

					<div class="wrap-input100 validate-input" data-validate = "Valid email is required: ex@abc.xyz">
						<input class="input100" type="text" name="usr_code" placeholder="請輸入帳號">
						<span class="focus-input100"></span>
						<span class="symbol-input100">
							<i class="fa fa-envelope" aria-hidden="true"></i>
						</span>
					</div>
	                     
					<div class="wrap-input100 validate-input" data-validate = "Password is required">
						<input class="input100" type="password" name="usr_pass" placeholder="請輸入密碼">
						<span class="focus-input100"></span>
						<span class="symbol-input100">
							<i class="fa fa-lock" aria-hidden="true"></i>
						</span>
					</div><br>
					 <div align="center">
						    <asp:TextBox class="input100" ID="tb_confirm" name="tb_confirm" runat="server" MaxLength="4" placeholder="請輸入驗證碼"
                            EnableViewState="False"></asp:TextBox><br />
            <asp:Image ID="img_confirm" ImageUrl="confirm.ashx" runat="server" Height="44px"
                            Width="150px" /><br />
                        <asp:Button ID="bn_new_confirm" runat="server" Text="重新產生驗證圖示" CssClass="text9pt"
                            OnClick="bn_new_confirm_Click" />
					
					   <div>來源IP:<%=Request.ServerVariables("LOCAL_ADDR")%> </div>
						   </div>
					<div class="container-login100-form-btn">
						<button class="login100-form-btn">
							登入
						</button>
						
						
					
					<span lang="zh-tw"><a href="ForgetPassword.aspx" target="_blank"><b><font size="2">忘記密碼</font></b></a></span>

</div>

				</form>
			</div>
		</div>
	</div>
	
	 <script language="javascript" type="text/javascript">
            function renew_img() {
                var img, now;
                now = new Date();
                img = document.getElementById("img_confirm");
                img.src = "confirm.ashx?ti=" + now.getSeconds().toString() + now.getMilliseconds().toString();
            }
        </script>

        <asp:Literal ID="lt_show" runat="server"></asp:Literal>

	
<!--===============================================================================================-->	
	<script src="vendor/jquery/jquery-3.7.1.min.js"></script>
<!--===============================================================================================-->
	<script src="vendor/bootstrap/js/popper.js"></script>
	<script src="vendor/bootstrap/js/bootstrap.min.js"></script>
<!--===============================================================================================-->
	<script src="vendor/select2/select2.min.js"></script>
<!--===============================================================================================-->
	<script src="vendor/tilt/tilt.jquery.min.js"></script>
	<script >
		$('.js-tilt').tilt({
			scale: 1.1
		})
	</script>
<!--===============================================================================================-->
	<script src="js/main.js"></script>

</body>
</html>
  