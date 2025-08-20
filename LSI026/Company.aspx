<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Company.aspx.vb" Inherits="basic_LSI026_Company" %>

<%@ Register Src="../../usercontrol/WebUserMenu.ascx" TagName="WebUserMenu" TagPrefix="uc1" %>
<%@ Register Src="../../usercontrol/WebSubMenu.ascx" TagName="WebSubMenu" TagPrefix="uc2" %>
<!DOCTYPE html>


<script type="text/css">
.gvStyle th
{
    background-color: #E2EAF2;
    font-weight: lighter;
    border: 1px solid #ccc;
    height:25px;
    text-align:center;
}
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
         <div id="anticfrs" runat="server"></div>
        <table id="customers" style="position: relative; margin-left: auto; margin-right: auto; width: 990px;" border="0">
            <tr>
                <td width="100%">
                    <div>
                        <uc1:WebUserMenu ID="WebUserMenu" runat="server" />
                    </div>
                </td>
            </tr>
            <tr>
                <td>
                    <div style="text-align: center;">
                        <asp:Label ID="Label1" runat="server" Style="text-align: center" Font-Bold="True" />
                    </div>
                    <uc2:WebSubMenu ID="WebSubMenu" runat="server" />

                    <table width="100%" border="1">
                        <tr class="style8">
                            <td style="width: 42%">請選擇單位：
                    <asp:DropDownList
                        ID="ddlField" runat="server"
                        DataValueField="field_code">
                        <asp:ListItem Text="一般科"></asp:ListItem>
                        <asp:ListItem Text="勞條科"></asp:ListItem>
                    </asp:DropDownList>
                                <%--   <asp:TextBox ID="TxtFind" runat="server" CssClass="findtb"></asp:TextBox>
                    <asp:Button ID="Button1" runat="server" ToolTip="查詢" Text="查詢" PostBackUrl="Default.aspx" />--%>
                                <asp:Button ID="Bt_srch" runat="server" ToolTip="查詢" Text="查詢" OnClick="Bt_srch_Click" />
                            </td>
                            <td style="text-align: center">

                                 <asp:ImageButton ID="ImageButton1" runat="server" Width="40px"  ImageAlign="Left"
                                    ImageUrl="~/images/icons/favicon.ico" ToolTip="回上頁" PostBackUrl="~/basic/LSI026/Default.aspx" />
                                <asp:ImageButton ID="ImageButton3" runat="server" Width="40px"  ImageAlign="Left"
                                    ImageUrl="~/images/icon/preview.png" ToolTip="公司申報表" PostBackUrl="~/basic/LSI026/com_check.aspx" />
                                
                                <asp:ImageButton ID="IB_import" runat="server" Width="40px"  ImageAlign="Left"
                                    ImageUrl="~/images/icon/insert.png" ToolTip="匯入Excel" />
                            </td>
                            <td style="width: 28%; text-align: right;">
                                <asp:Label ID="Label2" runat="server" Font-Bold="False" Style="text-align: center" Font-Size="Medium" />
                            </td>
                        </tr>
                        <tr class="style8">
                            <td colspan="3">
                                <asp:ImageButton ID="IB_GV1_First1" runat="server" ImageUrl="~/images/First.gif" PostBackUrl="company.aspx" ToolTip="第一頁" />
                                <asp:ImageButton ID="IB_GV1_Previous1" runat="server" ImageUrl="~/images/Previous.gif" PostBackUrl="company.aspx" ToolTip="上一頁" />
                                <asp:ImageButton ID="IB_GV1_Next1" runat="server" ImageUrl="~/images/Next.gif" PostBackUrl="company.aspx" ToolTip="下一頁" />
                                <asp:ImageButton ID="IB_GV1_Last1" runat="server" ImageUrl="~/images/Last.gif" PostBackUrl="company.aspx" ToolTip="最後頁" />
                                為確保查詢效能，預設只顯示當年度資料。
                    <asp:GridView ID="GridView1" runat="server" AllowPaging="True"
                        AutoGenerateColumns="False"
                        DataSourceID="SqlDS1" BackColor="White" BorderColor="#CCCCCC"
                        BorderStyle="None" BorderWidth="1px" CellPadding="3">
                        <RowStyle CssClass="grid4" ForeColor="#000066" />
                        <Columns>
                         
                                <asp:TemplateField ShowHeader="False" ItemStyle-Width="10%">
                                    <itemtemplate>
                                    <asp:ImageButton ID="ImageButton2" runat="server" CausesValidation="false" CommandName="delete"
                                        ImageUrl="~/images/icon/delete.gif" OnClientClick="return confirm('確定要刪除資料?');"
                                        Text="按鈕" ToolTip="刪除" PostBackUrl="Default.aspx" />
                                </itemtemplate>
                                    <%--  <HeaderStyle CssClass="grid1" />--%>
                                    <itemstyle horizontalalign="Center" />
                                </asp:TemplateField>
                                <%--                      <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:ImageButton ID="ImageButton3" runat="server" CausesValidation="false" CommandArgument="MDY"
                                        CommandName="select" ImageUrl="~/images/icon/edit.gif"
                                        Text="按鈕" ToolTip="修改" />
                                </ItemTemplate>
                                <HeaderStyle CssClass="grid1" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>
             <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:ImageButton ID="ImageButton4" runat="server" CausesValidation="false" CommandArgument="COPY"
                                        CommandName="select" ImageUrl="~/images/icon/copy.png"
                                        Text="按鈕" ToolTip="複製" Visible="False" />
                                </ItemTemplate>
                                <HeaderStyle CssClass="grid1" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>--%>
                        </Columns>
                        <FooterStyle CssClass="grid1" ForeColor="#000066" BackColor="White" />
                        <SelectedRowStyle CssClass="grid8" BackColor="#669999" Font-Bold="True"
                            ForeColor="White" />
                        <HeaderStyle CssClass="grid3" Font-Bold="True" BackColor="#006699"
                            ForeColor="White" />
                        <PagerStyle CssClass="grid5" HorizontalAlign="Left" BackColor="White"
                            ForeColor="#000066" />
                        <EditRowStyle CssClass="grid6" />
                        <AlternatingRowStyle CssClass="grid7" />
                    </asp:GridView>
                                <asp:ImageButton ID="IB_GV1_First2" runat="server" ImageUrl="~/images/First.gif" PostBackUrl="company.aspx" ToolTip="第一頁" />
                                <asp:ImageButton ID="IB_GV1_Previous2" runat="server" ImageUrl="~/images/Previous.gif" PostBackUrl="company.aspx" ToolTip="上一頁" />
                                <asp:ImageButton ID="IB_GV1_Next2" runat="server" ImageUrl="~/images/Next.gif" PostBackUrl="company.aspx" ToolTip="下一頁" />
                                <asp:ImageButton ID="IB_GV1_Last2" runat="server" ImageUrl="~/images/Last.gif" PostBackUrl="company.aspx" ToolTip="最後頁" />
                                <asp:Label ID="l_PageMsg" runat="server" />
                            </td>
                        </tr>
                    </table>

                    <asp:TextBox ID="txtSubWhere" runat="server" Visible="False" />
                    <asp:TextBox ID="txtOrder" runat="server" Visible="False" />
                    <asp:TextBox ID="txtPrimary" runat="server" Visible="False" />
                    <asp:TextBox ID="txtPgmType" runat="server" Visible="False" />
                    <asp:SqlDataSource ID="SqlDS1" runat="server" />
                    <asp:SqlDataSource ID="SqlDS2" runat="server" />
                    <%--</div>--%>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
