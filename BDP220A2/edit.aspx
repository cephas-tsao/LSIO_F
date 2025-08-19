<%@ Page Language="VB" AutoEventWireup="false" CodeFile="edit.aspx.vb" Inherits="management_BDP220A_edit" MaintainScrollPositionOnPostback="true" %>
<%@ Register src="../../usercontrol/WebUserMenu.ascx" tagname="WebUserMenu" tagprefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title><%Response.Write(ConfigurationManager.AppSettings("Prj_name"))%></title>
    <script language="javascript" id="js1" src="../../public/checkjs.js"></script> 
    <link rel='stylesheet' href ="/../Setting/style<%response.write(Get_SysPara("css_sys", "01"))%>.css" type = 'text/css'/>
</head>
<body>
    <form id="form1" runat="server">
        <table style="position:relative; margin-left:auto; margin-right:auto; width:990px; "  border ="0">
        <tr>
         <td>
           <div>
                <uc1:WebUserMenu ID="WebUserMenu2" runat="server" />
            </div>
         </td>
        </tr>
        <tr>
            <td><asp:LinkButton ID="Bt_BackUp" runat="server">回上頁</asp:LinkButton></td>
        </tr>
        <tr>
        <td>
        
    <div>
        <table width="100%" border="1" cellspacing='0.1'>
            <tr>
                <td class="fail" colspan="6" style="text-align: center">
                    <asp:Label ID="Label1" runat="server" Font-Bold="True" Style="text-align: center"></asp:Label></td>
            </tr>
            <tr>
                <td class="style1" style="width: 13%">
                    程式編號<asp:Image ID="Image2" runat="server" ImageUrl="~/images/tick.ico" /></td>
                <td class="style2" style="width: 100%">
                    <asp:TextBox ID="txt_prg_code" runat="server" MaxLength="14" Width="150px"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="style1" style="width: 13%">
                    欄位代碼<asp:Image ID="Image1" runat="server" ImageUrl="~/images/tick.ico" /></td>
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
                    資料來源<asp:Image ID="Image4" runat="server" ImageUrl="~/images/tick.ico" /></td>
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
                    順序<asp:Image ID="Image5" runat="server" ImageUrl="~/images/tick.ico" /></td>
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
            <asp:TextBox ID="txt_pgm_type" runat="server" Visible="False"></asp:TextBox>
            <asp:TextBox ID="txtPrimary" runat="server" Visible="False"></asp:TextBox>
            <asp:TextBox ID="txtSubWhere" runat="server" Visible="False"></asp:TextBox>
            <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="True"
            ShowSummary="False" />    
                    </td>
            </tr>
          </table> 
          

    </form>
</body>
</html>
