<%@ Page Language="VB" AutoEventWireup="false"  validateRequest="False" ContentType="text/html" ResponseEncoding="utf-8"  Debug="true" %>
<%@ Register src="../../usercontrol/WebUserMenu.ascx" tagname="WebUserMenu" tagprefix="uc1" %>
<%@ Register src="../../usercontrol/WebSubMenu.ascx" tagname="WebSubMenu" tagprefix="uc2" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
<title><%=ConfigurationManager.AppSettings("Prj_name")%></title>
<link rel="stylesheet" href="https://pro.fontawesome.com/releases/v5.10.0/css/all.css" integrity="sha384-AYmEC3Yw5cVb3ZcuHtOA93w35dYTsvhLPVnYs9eStHfGJvOvKxVfELGroGkvsg+p" crossorigin="anonymous"/>
<script type="text/css">
.gvStyle th
{
    background-color: #E2EAF2;
    font-weight: lighter;
    border: 1px solid #ccc;
    height:25px;
    text-align:center;
}
</script>
<script language="javascript">
function Chk_All_1() {
    for (var i = 0; i < document.form1.length; i++) {
        if (document.form1.elements[i].type == 'checkbox') {
            if (document.form1.elements[i].name.charAt(0) == 'C') {
                if (document.forms[0].elements["CBL_Field1_0"].checked == true) {
                    document.form1.elements[i].checked = true;
                }
                else {
                    document.form1.elements[i].checked = false;
                }
            }
        }
    }
}
</SCRIPT>
</head>
<body >
    <form id="form1" runat="server">
			
     <table id="customers" style="position:relative; margin-left:auto; margin-right:auto; width:990px;"  border ="0">
       <tr>
         <td width="100%">
           <div>
                <uc1:WebUserMenu ID="WebUserMenu" runat="server" />
           </div>
         </td>
     </tr>
    </table>
	</form>
   
	<div align="center" style="width: 990px;border:1px #000000 solid;margin:0px auto;">
																					
<p><input type="button"  onclick="window.location.href='#'" value="各季統計">&nbsp;&nbsp;<input type="button"  onclick="window.location.href='#'" value="調查匯出">&nbsp;&nbsp;<input type="button"  onclick="window.location.href='#'" value="報表匯出"><br />
</p>
<table width="95%"  border="0" align="center" cellpadding="1" cellspacing="1" >
  <tbody>
    <tr>
      <td height="30" bgcolor="#39A96B">&nbsp;</td>
      <td bgcolor="#39A96B">&nbsp;</td>
      <td bgcolor="#39A96B">&nbsp;</td>
      <td bgcolor="#39A96B">&nbsp;</td>
      <td bgcolor="#39A96B">&nbsp;</td>
      <td bgcolor="#39A96B">&nbsp;</td>
      <td bgcolor="#39A96B">&nbsp;</td>
      <td bgcolor="#39A96B">&nbsp;</td>
    </tr>
    <tr>
      <td height="33" bgcolor="#9AEB65">&nbsp;</td>
      <td bgcolor="#9AEB65">&nbsp;</td>
      <td bgcolor="#9AEB65">&nbsp;</td>
      <td bgcolor="#9AEB65">&nbsp;</td>
      <td bgcolor="#9AEB65">&nbsp;</td>
      <td bgcolor="#9AEB65">&nbsp;</td>
      <td bgcolor="#9AEB65">&nbsp;</td>
      <td bgcolor="#9AEB65">&nbsp;</td>
    </tr>
    <tr>
      <td height="37" bgcolor="#D9F86B">&nbsp;</td>
      <td bgcolor="#D9F86B">&nbsp;</td>
      <td bgcolor="#D9F86B">&nbsp;</td>
      <td bgcolor="#D9F86B">&nbsp;</td>
      <td bgcolor="#D9F86B">&nbsp;</td>
      <td bgcolor="#D9F86B">&nbsp;</td>
      <td bgcolor="#D9F86B">&nbsp;</td>
      <td bgcolor="#D9F86B">&nbsp;</td>
    </tr>
  </tbody>
</table>
<p><br />
</p>
	
																						
</div>
	
	
</body>
</html>
