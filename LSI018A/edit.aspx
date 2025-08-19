<%@ Page Language="VB" AutoEventWireup="false" CodeFile="edit.aspx.vb" Inherits="management_BDP220A_edit" aspcompat="true" MaintainScrollPositionOnPostback="true" %>
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
            width: 13%;
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
                <td class="fail" colspan="6" style="text-align: center">
                    <asp:Label ID="Label1" runat="server" Font-Bold="True" Style="text-align: center"></asp:Label></td>
            </tr>
            <tr>
                <td class="style1" style="width: 13%">
                    程式編號<asp:Image ID="Image2" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2" style="width: 100%">
                    <asp:TextBox ID="txt_prg_code" runat="server" MaxLength="14" Width="150px"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="style1" style="width: 13%">
                    欄位代碼<asp:Image ID="Image1" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2" style="width: 35%">
                    <asp:TextBox ID="txt_field_code" runat="server" MaxLength="20" Width="200px"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="style1" style="width: 13%">
                    欄位名稱</td>
                <td class="style2" style="width: 35%">
                    <asp:TextBox ID="txt_field_name" runat="server" MaxLength="20" Width="200px"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="style1" style="width: 13%">
                    資料來源<asp:Image ID="Image4" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2" style="width: 35%">
                    <asp:TextBox ID="txt_data_source" runat="server" MaxLength="50" Width="500px"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="style1" style="width: 13%">
                    資料類型</td>
                <td class="style2" style="width: 35%">
                    <asp:DropDownList ID="ddl_data_type" runat="server">
                    </asp:DropDownList></td>
            </tr>
            <tr>
                <td class="style1" style="width: 13%">
                    順序<asp:Image ID="Image5" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2" style="width: 35%">
                    <asp:TextBox ID="txt_scr_no" runat="server" MaxLength="2" Width="40px"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="style1" style="width: 13%">
                    是否使用</td>
                <td class="style2" style="width: 35%">
                    <asp:DropDownList ID="ddl_is_use" runat="server">
                    </asp:DropDownList></td>
            </tr>          
        </table>
    </div>
            <asp:Button ID="Button1" runat="server" Text="存檔" />
            <asp:Button ID="Bt_BackUp" runat="server" Text="離開" />
            <asp:TextBox ID="txtPgmType" runat="server" Visible="False" />
            <asp:TextBox ID="txtPrimary" runat="server" Visible="False" />
            <asp:TextBox ID="txtSubWhere" runat="server" Visible="False" />
            <asp:TextBox ID="txtOrder" runat="server" Visible="False" />
            <br />
            <asp:Label ID="l_ErrMsg" runat="server" Font-Bold="True" ForeColor="Red" />
            </td>
        </tr>
      </table>    
    </form>
</body>
</html>
