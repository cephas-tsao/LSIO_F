<%@ Page Language="VB" AutoEventWireup="false" CodeFile="edit.aspx.vb" Inherits="management_BDP080A_edit" aspcompat="true" MaintainScrollPositionOnPostback="true" %>
<%@ Register src="../../usercontrol/WebUserMenu.ascx" tagname="WebUserMenu" tagprefix="uc1" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title><%=ConfigurationManager.AppSettings("Prj_name")%></title>
    <script language="javascript" id="js1" src="../../public/checkjs.js"></script> 
    <link rel="stylesheet" href="https://pro.fontawesome.com/releases/v5.10.0/css/all.css" integrity="sha384-AYmEC3Yw5cVb3ZcuHtOA93w35dYTsvhLPVnYs9eStHfGJvOvKxVfELGroGkvsg+p" crossorigin="anonymous"/><link rel='stylesheet' href ="/../Setting/style<%=Get_SysPara("css_sys", "01")%>.css" type = 'text/css'/>
    <style type="text/css">
        .style1
        {
            width: 16%;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <table style="position:relative; margin-left:auto; margin-right:auto; width:990px; "  border ="0">
        <tr>
           <td width="100%">
             <div>
                <uc1:WebUserMenu ID="WebUserMenu" runat="server" />
             </div>
           </td>
        </tr>
        <tr>        
           <td>
    <div>
        <table width="100%" border="1" id="customers" cellspacing='0.1'>
            <tr>
                <td class="fail" colspan="2" style="text-align: center">
                    <asp:Label ID="Label1" runat="server" Font-Bold="True" Style="text-align: center"></asp:Label></td>
            </tr>
            <tr>
                <td class="style1">
                    <asp:Label ID="Label2" runat="server" Text="使用者ID" Width="90px"></asp:Label>
                    <asp:Image ID="Image4" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_usr_code" runat="server" MaxLength="20" Width="190px"></asp:TextBox>
                    </td>
            </tr>
            <tr>
                <td class="style1">
                    <asp:Label ID="Label5" runat="server" Text="使用者名稱" Width="90px"></asp:Label>
                    <asp:Image ID="Image5" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_usr_name" runat="server" MaxLength="32" Width="290px"></asp:TextBox></td>
            </tr>
            <%--<tr>
                <td class="style1">
                    <asp:Label ID="Label9" runat="server" Text="身份證字號" Width="90px"></asp:Label>
                    <asp:Image ID="Image8" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_usr_idno" runat="server" MaxLength="10" Width="100px"></asp:TextBox></td>
            </tr>--%>
            <tr>
                <td class="style1">
                    <asp:Label ID="Label3" runat="server" Text="密碼" Width="90px"></asp:Label>
                    <asp:Image ID="Image3" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_usr_pass" runat="server" MaxLength="32" Width="228px" TextMode="Password"></asp:TextBox>
                    </td>
            </tr>
            <tr>
                <td class="style1">
                    <asp:Label ID="Label4" runat="server" Text="角色" Width="90px"></asp:Label>
                    <asp:Image ID="Image1" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:DropDownList ID="ddl_grp_code" runat="server">
                    </asp:DropDownList>
                    </td>
            </tr>
            <tr>
                <td class="style1">
                    <asp:Label ID="Label7" runat="server" Text="所屬科室" Width="90px"></asp:Label>
                    <asp:Image ID="Image2" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:DropDownList ID="ddl_dep_code" runat="server">
                    </asp:DropDownList>
                    </td>
            </tr>
            <!--
            <tr>
                <td class="style1">
                    <asp:Label ID="Label8" runat="server" Text="層級" Width="90px"></asp:Label>
                    <asp:Image ID="Image7" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:DropDownList ID="ddl_limit_level" runat="server">
                    </asp:DropDownList>
                    </td>
            </tr>
            -->
            <tr>
                <td class="style1">
                    <asp:Label ID="Label6" runat="server" Text="權限類別" Width="90px"></asp:Label>
                    <asp:Image ID="Image6" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:DropDownList ID="ddl_Limit_type" runat="server">
                    </asp:DropDownList>
                    </td>
            </tr>
            
            <tr>
                <td class="style1">
                    Mail</td>
                <td class="style2">
                    <asp:TextBox ID="txt_usr_mail" runat="server" MaxLength="100" Width="328px"></asp:TextBox>
                    </td>
            </tr>
            <tr>
                <td class="style1">
                    是否使用</td>
                <td class="style2">
                    <asp:DropDownList ID="ddl_is_use" runat="server">
                    </asp:DropDownList></td>
            </tr>
            <tr>
                <td class="style1">
                    備註</td>
                <td class="style2">
                    <asp:TextBox ID="txt_usr_memo" runat="server" MaxLength="100" Width="627px"></asp:TextBox>
                    </td>
            </tr>              
        </table>
    </div>
            <asp:Button ID="Button1" runat="server" Text="存檔" />
            <asp:Button ID="Bt_BackUp" runat="server" Text="離開" />
            <asp:TextBox ID="txtPgmType" runat="server" Visible="False" />
            <asp:TextBox ID="txtPrimary" runat="server" Visible="False" />
            <asp:TextBox ID="txtSubWhere" runat="server" Visible="False" />
            <asp:TextBox ID="txtOrder" runat="server" Visible="False" />
            <asp:TextBox ID="txtUsrIdno" runat="server" Visible="False" />
            <br />
            <asp:Label ID="l_ErrMsg" runat="server" Font-Bold="True" ForeColor="Red" />
            </td>
        </tr>
      </table>    
    </form>
</body>
</html>
