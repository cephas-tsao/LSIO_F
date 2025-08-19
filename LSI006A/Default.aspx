<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="basic_LSI006A_Default" aspcompat="true" validateRequest="false" %>
<%@ Register src="../../usercontrol/WebUserMenu.ascx" tagname="WebUserMenu" tagprefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>時間參數設定</title>
    <link rel="stylesheet" href="https://pro.fontawesome.com/releases/v5.10.0/css/all.css" integrity="sha384-AYmEC3Yw5cVb3ZcuHtOA93w35dYTsvhLPVnYs9eStHfGJvOvKxVfELGroGkvsg+p" crossorigin="anonymous"/><link rel='stylesheet' href ="/../Setting/style<%=Get_SysPara("css_sys", "01")%>.css" type = 'text/css'/>
     <script type="text/jscript">   
        function DisableButton()   
        {
           var btn = document.getElementById('Button1');
           btn.disabled = false;
        }
     </script>  
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
        <td>
        <div align="center">
        <table  style="width:50%;" border ="0" >
        <tr>
            <td  class="fail" colspan="2" align="center" bgcolor='#000090'>
                <asp:Label ID="Label1" runat="server" Text="Label"></asp:Label>
            </td>
        </tr>
        <tr>
            <td align="right" bgcolor='#99CCFF' class="alter01">
                寄件者名稱</td>
                <td bgcolor='#CCCCCC' align="left">
                    <asp:TextBox ID="txt_sender" runat="server" MaxLength="500" Width="200px"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td align="right" bgcolor='#99CCFF' class="alter01">
                郵件信箱</td>
                <td bgcolor='#CCCCCC' align="left">
                    <asp:TextBox ID="txt_sender_mail" runat="server" MaxLength="500" Width="200px"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td align="right" bgcolor='#99CCFF' class="alter01">
                回信信箱</td>
                <td bgcolor='#CCCCCC' align="left">
                    <asp:TextBox ID="txt_resender_mail" runat="server" MaxLength="500" 
                        Width="200px"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td  align="center" colspan="2" >
                <asp:Button ID="Button1" runat="server" Text="確認儲存" />
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