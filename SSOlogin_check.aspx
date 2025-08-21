<%@ Page Language="VB" AutoEventWireup="false"  validateRequest="False" ContentType="text/html" ResponseEncoding="utf-8" %>

<%@ Import Namespace="System.Diagnostics" %>

<%@ Import Namespace = "System.Data" %>
<%@ Import Namespace = "System.Data.SqlClient" %>
<%@ Import NameSpace = "System.IO"%>
<%@ Import NameSpace = "System.Web.Configuration" %>
<%@ Import NameSpace = "System.Threading.Tasks" %>
<%@ Import NameSpace = "System.Threading" %>

<%
Dim Conn1 As SqlConnection = New SqlConnection
      Conn1.ConnectionString = WebConfigurationManager.ConnectionStrings("AppSysConnectionString").ConnectionString  
%>

<!DOCTYPE html>

<html>

<head>

    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="description" content="">
    <meta name="author" content="">

    <title>SSO驗證</title>

</head>

<body style="font-size: 10pt;">

           
<p>

	
	
  
      
  <br>
</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p align="center"><img src="images.png" width="259" height="194" alt=""/></p>
<p align="center">SSO 驗證結果</p>
	
	
	<%  
Conn1.open()
   Dim strWhere As String = ""
        strWhere &= "usr_mail = @usr_mail"
        Dim sqlstr1 As String = "Select * From BDP080 where  " & strWhere & ""
        Dim cmd1 As SqlCommand = New SqlCommand(sqlstr1, Conn1)
        Dim queryStr As String = Request("usr_code")
        Dim queryPass As String = Request("usr_code")
        cmd1.Parameters.Add("@usr_mail", SqlDbType.NVarChar).Value = queryStr
        cmd1.Parameters.Add("@usr_pass", SqlDbType.NVarChar).Value = queryPass
        Dim dr1 As SQLDataReader = cmd1.ExecuteReader()

	  
	 if  dr1.Read()  then 
	   Dim usr_name = dr1.Item("usr_name") 
	    Dim login_time =Format(Now(), "yyyy-MM-dd HH:mm:ss")
	   
	         Session("usr_code") = dr1.Item("usr_code")
             Session("usr_name") = dr1.Item("usr_name")
             Session("login_time") = Format(Now(), "HH:mm")
	        Session("grp_code") = dr1.Item("grp_code")
	        Response.Redirect("blank.aspx")
	     %>
	    
	<table width="411" border="1" align="center">
  <tbody>
    <tr>
      <td width="41" align="center">員工姓名</td>
      <td width="143" align="center">
	  <%
	  if usr_name = "admin" then
	  usr_name = "蘇鵬楠"
	  end if
	  
	  
	  %>
	  
	  
	  <%=usr_name%></td>
    </tr>
    <tr>
      <td align="center">登入時間</td>
      <td align="center" valign="middle"><%=login_time%></td>
    </tr>
  </tbody>
</table>
<p>&nbsp;</p>
                
     
 <div align="center">

  <button type="button"  onClick="window.location.href='blank.aspx'" >登入系統</button>&emsp;
    <button type="button"  onClick="window.location.href='SSOlogin.aspx'">結束離開</button> 
 
	</div>
			
			
            
	   <%else%>
		   
		    <div align="center">
				<span style="color: red; font-size: 16px;">驗證失敗</span>
				
				<br><br><br><br>

    <button type="button"  onClick="window.location.href='SSOlogin.aspx'">結束離開</button> 
           
	</div>
			
	   
	    <%
	 end if 
	 dr1.close  
	    Conn1.Close()

	 
%>


    <!-- jQuery -->
    <script src="../vendor/jquery/jquery.min.js"></script>  
    <script src="../js/jquery.blockUI.js" type="text/javascript" ></script>
    <script type="text/javascript" src="../js/jquery-ui-1.12.1.custom/jquery-ui.min.js"></script>
    <script type="text/javascript" src="ckeditor/ckeditor.js"></script>

<link href="../js/jquery-ui-1.12.1.custom/jquery-ui.min.css"rel="stylesheet" type="text/css"/>
    
    
    <script>
	
	
  
function msgre_post(x){  

var int = x ;

var adm_title = $("#adm_title").val();
var adm_name = $('select[name="adm_name"]').val();  
var adm_type = $('select[name="adm_type"]').val();  
var adm_fmail = $("#adm_fmail").val(); 
var classtxt = CKEDITOR.instances["Text-Display"].getData();
var dateuser ="<%=Format(Now(), "yyyyMMdd")%>"; 
var indx="<%=Request("indx")%>";
var active="up1";
var adm_fname = $("#adm_fname").val(); 
	
  
  $.blockUI({
	message:'<img src=loading.gif class=loading>記錄中.....',
	css:{border:'3px solid white',background:'#CCCCCC',padding:'10px',color:'#000000' }
	
	});

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
   
   <script>
   CKEDITOR.replace("Text-Display");
</script>
    
</body>

</html>
