<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="management_BDP003_A_Default" MaintainScrollPositionOnPostback="true" %>
<%@ Register src="../../usercontrol/WebUserMenu.ascx" tagname="WebUserMenu" tagprefix="uc1" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title><%Response.Write(ConfigurationManager.AppSettings("Prj_name"))%></title>
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
        <td>
        
    <div style="text-align: center"><asp:Label ID="Label1" runat="server" style="text-align: center" Font-Bold="True"></asp:Label></div>
    <div>
        <table width="100%" border ="1">
            <tr style="background-color:#d3d3d3">
                <td style="width: 70%">
                    <asp:Image ID="Image1" runat="server" ImageUrl="~/images/bs_search.gif" />搜尋<asp:DropDownList
                        ID="ddlField" runat="server" DataSourceID="SqlDS2" DataTextField="field_name"
                        DataValueField="field_code">
                    </asp:DropDownList>
                    <asp:TextBox ID="TxtFind" runat="server"></asp:TextBox>
                    &nbsp;<asp:ImageButton ID="Button1" runat="server" ImageUrl="~/images/icon/find2.png"
                        ToolTip="查詢" /><asp:ImageButton ID="Button4" runat="server" CommandArgument="ADD"
                            ImageUrl="~/images/icon/add.png" PostBackUrl="edit.aspx" ToolTip="新增" />
                    <asp:Button ID="Button2" runat="server" Text="進階" Visible="False" /></td>
                <td style="width: 30%; text-align: right;">
                    <asp:Label ID="Label2" runat="server" Font-Bold="False" Style="text-align: center" Font-Size="Medium"></asp:Label>
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AutoGenerateColumns="False"
                        BackColor="LightGoldenrodYellow" BorderColor="Tan" BorderWidth="1px" CellPadding="2"
                        DataSourceID="SqlDS1" ForeColor="Black" GridLines="None">
                        <Columns>
                            <asp:TemplateField ShowHeader="False">
                                <ItemTemplate>
                                    <asp:ImageButton ID="ImageButton2" runat="server" CausesValidation="false" CommandName="delete"
                                        ImageUrl="~/images/icon/delete.gif" OnClientClick="return confirm('確定要刪除資料?');"
                                        Text="按鈕" ToolTip="刪除" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>
                            <asp:TemplateField Visible="False">
                                <ItemTemplate>
                                    <asp:ImageButton ID="ImageButton3" runat="server" CausesValidation="false" CommandArgument="MDY"
                                        CommandName="select" ImageUrl="~/images/icon/edit.gif" PostBackUrl="edit.aspx"
                                        Text="按鈕" ToolTip="修改" />&nbsp;
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:ImageButton ID="ImageButton4" runat="server" CausesValidation="false" CommandArgument="COPY"
                                        CommandName="select" ImageUrl="~/images/icon/copy.png" PostBackUrl="edit.aspx"
                                        Text="按鈕" ToolTip="複製" Visible="False" />
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>
                        </Columns>
                        <FooterStyle BackColor="Tan" />
                        <SelectedRowStyle BackColor="DarkSlateBlue" ForeColor="GhostWhite" />
                        <PagerStyle BackColor="PaleGoldenrod" ForeColor="DarkSlateBlue" HorizontalAlign="Center" />
                        <HeaderStyle BackColor="Tan" Font-Bold="True" />
                        <AlternatingRowStyle BackColor="FloralWhite" />
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td style="width: 70%">
                    <asp:TextBox ID="txtSubWhere" runat="server" Visible="False"></asp:TextBox>
                    <asp:TextBox ID="txtPrimary" runat="server" Visible="False"></asp:TextBox>
                    <asp:TextBox ID="txtPgmType" runat="server" Visible="False"></asp:TextBox></td>
                <td style="width: 30%">
                    </td>
            </tr>
        </table>
        &nbsp;&nbsp;
        <asp:SqlDataSource ID="SqlDS1" runat="server"></asp:SqlDataSource><asp:SqlDataSource ID="SqlDS2" runat="server"></asp:SqlDataSource>
        &nbsp;
    
    </div>
    
        </td>
    </tr>
    </table>
    
    </form>
</body>
</html>
