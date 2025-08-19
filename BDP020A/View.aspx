<%@ Page Language="VB" AutoEventWireup="false" CodeFile="View.aspx.vb" Inherits="management_BDP019A_View" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title><%Response.Write(ConfigurationManager.AppSettings("Prj_name"))%></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table width="100%" border="1">
            <tr>
                <td colspan="2" style="text-align: center">
                    <asp:Label ID="Label1" runat="server" Font-Bold="True" Style="text-align: center"></asp:Label></td>
            </tr>
            <tr>
                <td width="15%">
                    <span lang="zh-tw">公告主旨</span></td>
                <td>
                    <asp:Label ID="Label3" runat="server" Text="Label" Width="95%"></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                    <span lang="zh-tw">公告人</span></td>
                <td>
                    <asp:Label ID="Label4" runat="server" Text="Label" Width="95%"></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                    <span lang="zh-tw">公告內文</span></td>
                <td>
                    <asp:Label ID="Label5" runat="server" Text="Label" Width="95%"></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                    <span lang="zh-tw">附加檔案</span></td>
                <td>
                    <asp:Label ID="Label2" runat="server" Text="無" Visible="False"></asp:Label>
                    <asp:ImageButton ID="ImageButton1" runat="server" 
                        ImageUrl="~/images/icon/download.gif" Visible="False" />
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <span lang="zh-tw">
                    <asp:Label ID="Label6" runat="server" Text="Label" Width="95%"></asp:Label>
                    </span></td>
            </tr>
            <tr>
                <td colspan="2">
                    <asp:Button ID="Button1" runat="server" Text="回上一頁" 
                        PostBackUrl="~/management/BDP019A/Default.aspx" />
                    <span lang="zh-tw">&nbsp;
                    </span>
                    <asp:TextBox ID="txt_pgm_type" runat="server" Visible="False"></asp:TextBox>
                    <asp:TextBox ID="txtSubWhere" runat="server" Visible="False"></asp:TextBox>
                    <asp:TextBox ID="TextBox1" runat="server" Visible="False"></asp:TextBox>
                </td>
            </tr>            
          
            </table>
    </div>
        
    </form>
</body>
</html>