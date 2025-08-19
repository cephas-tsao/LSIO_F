<%@ Page Language="VB" AutoEventWireup="false" CodeFile="edit.aspx.vb" Inherits="basic_LSI019A_edit" aspcompat="true" MaintainScrollPositionOnPostback="true" %>
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
                <td class="fail" colspan="2" style="text-align: center">
                    <asp:Label ID="Label1" runat="server" Font-Bold="True" Style="text-align: center"></asp:Label></td>
            </tr>
            <tr>
                <td class="style1">
                    第1次檢查日期</td>
                <td class="style2">
                    <asp:TextBox ID="txt_chk_date1" runat="server" MaxLength="10" Width="100px"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="style1">
                    第2次檢查日期</td>
                <td class="style2">
                    <asp:TextBox ID="txt_chk_date2" runat="server" MaxLength="10" Width="100px"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="style1">
                    事業單位名稱</td>
                <td class="style2" colspan="2">
                    <asp:TextBox ID="txt_com_name" runat="server" MaxLength="50" Width="500px"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="style1">
                    工程名稱</td>
                <td class="style2" colspan="2">
                    <asp:TextBox ID="txt_eng_name" runat="server" MaxLength="50" Width="500px"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="style1">
                    工地狀態</td>
                <td class="style2" colspan="2">
                    <asp:TextBox ID="txt_work_staus" runat="server" MaxLength="20" Width="220px"></asp:TextBox></td>
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
