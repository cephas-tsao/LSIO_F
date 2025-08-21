<%@ Page Language="VB" AutoEventWireup="false"  CodeFile="SSOlogin.aspx.vb"  Inherits="SSOlogin"  %>

<%@ import Namespace="System.data" %>
<%@ import Namespace="System.Web.UI.Control" %>
<%@ import Namespace="System.Data.SqlClient" %>
<%@ import Namespace="System.Web.Configuration" %>
<%@ import Namespace="Microsoft.Security.Application" %>



  
 <!DOCTYPE html>
<html lang="en">
<head>
	<title>SSO登入</title>
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
					<img src="images/img-01.png" alt="IMG"><br><br><br>
					<button class="login100-form-btn" onClick="window.location.href='Default.aspx'">
							一般登入
					</button>
				</div>

				<Form name="Form1" id="Form1" Action="SSOlogin_check.aspx" class="login100-form validate-form"  Method="POST">
                       <div id="anticfrs" runat="server"></div>
					<span class="login100-form-title">
<img src="images/img-021.png">
					</span>
<div align="center">SSO「帳號」請輸入員工mail帳號</div>
					<div class="wrap-input100 validate-input" data-validate = "Valid email is required: ex@abc.xyz">
						<input class="input100" type="text" name="usr_code" id="usr_code" value="lio.119@mail.taipei.gov.tw" placeholder="請輸入員工mail帳號">
						<span class="focus-input100"></span>
						<span class="symbol-input100">
							<i class="fa fa-envelope" aria-hidden="true"></i>
						</span>
					</div>
	                     
					<div class="wrap-input100 validate-input" data-validate = "Password is required">
						<input class="input100" type="password" name="usr_pass" id="usr_pass"  placeholder="請輸入密碼">
						<span class="focus-input100"></span>
						<span class="symbol-input100">
							<i class="fa fa-lock" aria-hidden="true"></i>
						</span>
					</div><br>
					 <div align="center">
						   
					
					   <div>來源IP:<%=Request.ServerVariables("LOCAL_ADDR")%> </div>
						   </div>
					<div class="container-login100-form-btn">
						
						<button type="button" class="login100-form-btn" style="background-color:darkslateblue;" onClick="SSOlogin()">
							SSO登入
						</button>
						
						
				<!--	
					<span lang="zh-tw"><a href="ForgetPassword.aspx" target="_blank"><b><font size="2">忘記密碼</font></b></a></span>
-->

</div>

				</Form>
			</div>
		</div>
	</div>
	
	 <script language="javascript" type="text/javascript">
		 

		 
var t = 5;
		 function token()
{
    t -= 1;
   // document.getElementById('div1').innerHTML= t;
    
    if(t==3)
    {
       $("#info").text("判讀token中......");
    }
	
	 if(t==0)
    {
       Form1.submit();
    }
    
    //每秒執行一次,showTime()
    setTimeout("token()",1000);
}
		 
		 
function SSOlogin(x){  

//var int = x ;
//
//var adm_title = $("#adm_title").val();
//var adm_name = $('select[name="adm_name"]').val();  
//var adm_type = $('select[name="adm_type"]').val();  
//var adm_fmail = $("#adm_fmail").val(); 
//var classtxt = CKEDITOR.instances["Text-Display"].getData();
//var dateuser ="<%=Format(Now(), "yyyyMMdd")%>"; 
//var indx="<%=Request("indx")%>";
//var active="up1";
//var adm_fname = $("#adm_fname").val(); 
	
  $.blockUI({
	message:'<img src=loading.gif class=loading><span id="info">提取SSO資訊中.....</span>',
	css:{border:'3px solid white',background:'#CCCCCC',padding:'10px',color:'#000000' }
	
	});
		
	 
token();
			 
	$.ajax({ 
		type: "post",  			
                url: 'msgweb_post10.aspx',
                cache: true,
				data: {
					active:active,
					indx:indx,
					adm_title:adm_title,
					adm_name:adm_name,
					adm_type:adm_type,
					classtxt:escape(classtxt),
					adm_fname:adm_fname,
					adm_fmail:adm_fmail		
				},
				success: function(data) { 
				    $.unblockUI();
				    console.log(data);
					alert('更新成功');	
					//window.location.href='6-1-web.aspx'           
				} ,
				error: function(data){
					alert('網路錯誤暫時無法存檔');	
					//alert(data);
					$.unblockUI();	
					
				}
            }); 

  }



   </script>
		 
		 
		 
		 
<%--            function renew_img() {
                var img, now;
                now = new Date();
                img = document.getElementById("img_confirm");
                img.src = "confirm.ashx?ti=" + now.getSeconds().toString() + now.getMilliseconds().toString();
            }
        </script>--%>

        <asp:Literal ID="lt_show" runat="server"></asp:Literal>

	
<!--===============================================================================================-->	
	<script src="vendor/jquery/jquery-3.7.1.min.js"></script>
	<script src="js/jquery.blockUI.js" type="text/javascript" ></script>
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
  