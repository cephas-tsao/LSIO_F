<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="management_BDP006A_Default" aspcompat="true" validateRequest="false" %>
<%@ Register src="../../usercontrol/WebUserMenu.ascx" tagname="WebUserMenu" tagprefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>使用者基本資料確認頁面</title>
    <link rel="stylesheet" href="https://pro.fontawesome.com/releases/v5.10.0/css/all.css" integrity="sha384-AYmEC3Yw5cVb3ZcuHtOA93w35dYTsvhLPVnYs9eStHfGJvOvKxVfELGroGkvsg+p" crossorigin="anonymous"/><link rel='stylesheet' href ="/../Setting/style<%=Get_SysPara("css_sys", "01")%>.css" type = 'text/css'/>
     <script type="text/jscript">   
        function DisableButton()   
        {
           var btn = document.getElementById('Button1');
           btn.disabled = false;
        }
     </script>  
    <style type="text/css">
        .style1
        {
            width: 223px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server" defaultbutton="Button1">
    <table style="position:relative; margin-left:auto; margin-right:auto; width:990px; "  border ="0">
     <tr>
         <td>
           <div>
                <uc1:WebUserMenu ID="WebUserMenu" runat="server" />
            </div>
         </td>
     </tr>
     <tr>
        <td >
    <div align="center">
    <table style="width:50%;" border ="0" >
        <tr>
            <td colspan="2" align="center" bgcolor='#000090'>
                <asp:Label ID="Label1" runat="server" Text="Label"></asp:Label>
            </td>
        </tr>
        <tr>
            <td align="right" bgcolor='#99CCFF' class="alter01">
                請輸入舊密碼</td>
                <td bgcolor='#CCCCCC' align="left">
                    <asp:TextBox ID="txt_UsrOldPW" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td align="right" bgcolor='#99CCFF' class="alter01">
                    <asp:Label ID="lab_UsrNewPW" runat="server" Text="請輸入新密碼"></asp:Label>
                </td>
                <td bgcolor='#CCCCCC' align="left">
                    <asp:TextBox ID="txt_UsrNewPW" runat="server" MaxLength="32"></asp:TextBox>
                </td>
        </tr>
        <tr>
            <td align="right" bgcolor='#99CCFF' class="alter01">
                <span lang="zh-tw">
                    <asp:Label ID="lab_UsrNewPW2" runat="server" Text="請再輸入一次新密碼"></asp:Label>
                </span>
                </td>
                <td bgcolor='#CCCCCC' align="left">
                    <asp:TextBox ID="txt_UsrNewPW2" runat="server" MaxLength="32"></asp:TextBox>
                    </td>
        </tr>
         <tr>
                <td  align="left">
                    <asp:Button ID="Button1" runat="server" Text="確認儲存" />
                    <asp:TextBox ID="txt_UsrPW" runat="server" Visible="false"></asp:TextBox>
                    <asp:TextBox ID="txt_UsrCode" runat="server" Visible="false"></asp:TextBox>
                </td>
            </tr>
        </table>
        </div>
       </td>
       </tr>       
    </table>
    <br/>
    </form>
</body>
</html>