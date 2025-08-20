<%@ Page Language="VB" AutoEventWireup="false" CodeFile="addForm.aspx.vb" Inherits="public_addForm" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>資料檢視頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div style="text-align: center"><asp:Label ID="Label1" runat="server" style="text-align: center" Font-Bold="True"></asp:Label></div>
    <div>
        <table width="100%" border ="1">
            <tr style="background-color:#d3d3d3">
                <td style="width: 70%">
                    <asp:Image ID="Image1" runat="server" ImageUrl="~/images/bs_search.gif" />搜尋<asp:DropDownList
                        ID="ddlField" runat="server" DataSourceID="SqlDS2" DataTextField="field_name"
                        DataValueField="sel_field">
                    </asp:DropDownList>
                    <asp:TextBox ID="TxtFind" runat="server"></asp:TextBox>
                    <asp:Button ID="Button1" runat="server" Text="查詢" />&nbsp;
                </td>
                <td style="width: 30%; text-align: right;">
                    <asp:Label ID="Label2" runat="server" Font-Bold="False" Style="text-align: center" Font-Size="Medium"></asp:Label>
                </td>
            </tr>
            <tr>
                <td colspan="2">
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" AllowPaging="True" BackColor="LightGoldenrodYellow" BorderColor="Tan" BorderWidth="1px" CellPadding="2" DataSourceID="SqlDS1" ForeColor="Black" GridLines="None">
            <FooterStyle BackColor="Tan" />
            <SelectedRowStyle BackColor="DarkSlateBlue" ForeColor="GhostWhite" />
            <PagerStyle BackColor="PaleGoldenrod" ForeColor="DarkSlateBlue" HorizontalAlign="Center" />
            <HeaderStyle BackColor="Tan" Font-Bold="True" />
            <AlternatingRowStyle BackColor="FloralWhite" />
            <Columns>
                <asp:CommandField SelectText="取回" ShowSelectButton="True" />
            </Columns>
        </asp:GridView>
                </td>
            </tr>
            <tr>
                <td style="width: 70%">
                    <asp:TextBox ID="txtSubWhere" runat="server" Visible="False"></asp:TextBox>
                    <asp:TextBox ID="AddCode" runat="server" Visible="False"></asp:TextBox>
                    <asp:TextBox ID="FieldList" runat="server" Visible="False"></asp:TextBox>
                    <asp:TextBox ID="TextBoxList" runat="server" Visible="False"></asp:TextBox>
                    <asp:TextBox ID="txtSubWhere2" runat="server" Visible="False"></asp:TextBox></td>
                <td style="width: 30%">
                    </td>
            </tr>
        </table>
        &nbsp;&nbsp;
        <asp:SqlDataSource ID="SqlDS1" runat="server"></asp:SqlDataSource><asp:SqlDataSource ID="SqlDS2" runat="server"></asp:SqlDataSource>
        &nbsp;
    
    </div>
    </form>
</body>
</html>
