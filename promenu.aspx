<%@ Page Language="VB" AutoEventWireup="false" CodeFile="promenu.aspx.vb" Inherits="promenu" %>

<%@ Register src="usercontrol/WebUserMenu.ascx" tagname="WebUserMenu" tagprefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>農藥延伸使用查詢系統</title>
</head>
<frameset cols="235,*" framespacing="1" border="1" frameborder="1">
  <frameset rows="100%,0">
    <frame name="left" id = "left" target="main" src="menulist.aspx" scrolling="yes">
  </frameset>
  <frame name="right" src="blank.aspx" target="_self" scrolling="auto">


<form id="form1" runat="server">
     <table style="position:relative; margin-left:auto; margin-right:auto; width:990px; "  border ="0">
     <tr>
         <td BackGround="../../images/default-head2.jpg" Height="90"  style="">
           <div>
                <uc1:WebUserMenu ID="WebUserMenu1" runat="server" />
            </div>
         </td>
     </tr>
     </table>
</form>
</html>
