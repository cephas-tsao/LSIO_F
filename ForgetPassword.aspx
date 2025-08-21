<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ForgetPassword.aspx.vb" Inherits="ForgetPassword" aspcompat="true" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>忘記密碼？</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
    </div>
    <table style="width:100%;" border="1" frame="border">
        <tr bgcolor="#000090"  style="width:100px;">
            <td colspan="2">
                <center><b><font color="white" size="4">忘記密碼</font></b></center>
            </td>
        </tr>
        <tr bgcolor="#507CD1">
            <td colspan="2">
                <span lang="zh-tw"><b><font color="white">請填寫以下資訊以便檢核，若檢核通過將自動發送帳號密碼至您所填寫的信箱</font></b></span></td>
        </tr>
        <tr bgcolor="#EFF3FB">
            <td width="15%">
                您的帳號</td>
            <td width="85%">
                <asp:TextBox ID="txt_id" runat="server" Width="120px"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td>
                <span lang="zh-tw">電子信箱</span></td>
            <td>
                <asp:TextBox ID="txt_mail" runat="server" Width="300px"></asp:TextBox>
            </td>
        </tr>
    </table>
    <span lang="zh-tw">&nbsp;</span><br />
    <asp:Button ID="Button1" runat="server" Text="確認送出" />
    </form>
    </body>
</html>
