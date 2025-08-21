<%@ Page Language="VB" AutoEventWireup="false" CodeFile="blank.aspx.vb" Inherits="blank" %>

<%@ Register src="usercontrol/WebUserMenu.ascx" tagname="WebUserMenu" tagprefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title><%Response.Write(ConfigurationManager.AppSettings("Prj_name"))%></title>
    <link rel='stylesheet' href ="/../Setting/style01.css" type = 'text/css'/>
    <style type="text/css">
        .style1
        {
            width: 990px;
        }
        .style3
        {
            width: 778px;
        }
        .style4
        {
            width: 205px;
        }
    </style>
</head>
<body>
    <form id="form2" runat="server">
    
    <table style="width:990px;"  border ="0" align=center>
    <tr>
        <td colspan="2" >
           <div>
           <uc1:WebUserMenu ID="WebUserMenu1" runat="server" />
           </div>
        </td>
    </tr>
    <tr>
      <td colspan="2" valign=top align=center style="width:50%;">
      <table><tr><td align=left>
      本網站可於Internet Explorer 8.0、Mozilla Firefox 8.0、<br />
      Apple Safari 5.1等瀏覽器完整使用系統功能，建議使用<br />
      相同或更高版本瀏覽器以得到最好的效果。<br />
      </td></tr></table><br />
    您使用的瀏覽器為：<asp:Label ID="BrowserName" runat="server" ForeColor="Blue" 
              Font-Bold="True" />　版本為：<asp:Label 
              ID="BrowserVersion" runat="server" ForeColor="Red" Font-Bold="True" />
      </td>
    </tr>
    </table>    
    <asp:TextBox ID="txt_group_code" runat="server" Visible="False"></asp:TextBox>
    <asp:TextBox ID="txt_group_name" runat="server" Visible="False"></asp:TextBox>
    </form>
</body>
</html>
