<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="basic_LSI008A_Default" MaintainScrollPositionOnPostback="true" EnableEventValidation = "false" %>
<%@ Register src="../../usercontrol/WebUserMenu.ascx" tagname="WebUserMenu" tagprefix="uc1" %>
<%@ Register src="../../usercontrol/MeetingOrder.ascx" tagname="MeetingOrder" tagprefix="uc2" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=ConfigurationManager.AppSettings("Prj_name")%></title>
<link rel="stylesheet" href="https://pro.fontawesome.com/releases/v5.10.0/css/all.css" integrity="sha384-AYmEC3Yw5cVb3ZcuHtOA93w35dYTsvhLPVnYs9eStHfGJvOvKxVfELGroGkvsg+p" crossorigin="anonymous"/><link rel='stylesheet' href ="/../Setting/style<%=Get_SysPara("css_sys", "01")%>.css" type = 'text/css'/>
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
</head>
<body>
    <form id="form1" runat="server">
      <table id="customers" style="position:relative; margin-left:auto; margin-right:auto; width:990px;"  border ="0">
        <tr>
          <td width="100%">
            <div>
              <uc1:WebUserMenu ID="WebUserMenu" runat="server" />
            </div>
          </td>
        </tr>
        <tr>
          <td><div style="text-align: center;">
            <asp:Label ID="Label1" runat="server" style="text-align: center" Font-Bold="True" />
          </div>
          <table width="100%" border ="1">
            <tr class="style8">
              <td colspan="3">
                  <asp:Label ID="l_yyyymm" runat="server" Text="起始月份：" Font-Bold="True" />
                  <asp:DropDownList ID="ddl_yyyymm" runat="server" AutoPostBack="True" />
                  <asp:DropDownList ID="ddl_room_code" runat="server" AutoPostBack="True" />
                  <asp:Button ID="bt_1" runat="server" Text="1個月" /> 
                  <asp:Button ID="bt_2" runat="server" Text="2個月" /> 
                  <asp:Button ID="bt_3" runat="server" Text="3個月" /> 
              </td>
            </tr>



              <tr class="style8">

                  <asp:Panel ID="Panel1_1" runat="server">

                      <td width="33%" style="text-align: center">
                          <asp:Label ID="l_MeetingOrder1" runat="server" Font-Bold="True" />
                      </td>
                  </asp:Panel>
                  <asp:Panel ID="Panel2_1" runat="server">

                      <td width="33%" style="text-align: center">
                          <asp:Label ID="l_MeetingOrder2" runat="server" Font-Bold="True" />
                      </td>
                  </asp:Panel>
                  <asp:Panel ID="Panel3_1" runat="server">

                      <td width="33%" style="text-align: center">
                          <asp:Label ID="l_MeetingOrder3" runat="server" Font-Bold="True" />
                      </td>
                  </asp:Panel>
                  </asp:Panel>

              </tr>
              <tr valign="top">
                  <asp:Panel ID="Panel1" runat="server">
                      <td>
                          <uc2:MeetingOrder ID="MeetingOrder1" runat="server" />
                      </td>
                  </asp:Panel>
                  <asp:Panel ID="Panel2" runat="server">
                      <td>
                          <uc2:MeetingOrder ID="MeetingOrder2" runat="server" />
                      </td>
                  </asp:Panel>
                  <asp:Panel ID="Panel3" runat="server">

                      <td>
                          <uc2:MeetingOrder ID="MeetingOrder3" runat="server" />
                      </td>
                  </asp:Panel>

              </tr>
          </table>
        </tr>
      </table>
    </form>
</body>
</html>
