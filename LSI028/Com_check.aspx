<%@ Page Language="VB" AutoEventWireup="true" CodeFile="Com_check.aspx.vb" EnableEventValidation = "false" Inherits="basic_LSI026_Com_check" %>

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
                            <td style="width: 42%">請選擇年度單位：
                                 <asp:DropDownList
                        ID="DDLYear" runat="server"
                        DataValueField="field_year">
                        <asp:ListItem Text="110" Value="110"></asp:ListItem>
                        <asp:ListItem Text="111" Value="111" Selected="True" ></asp:ListItem>
                        <asp:ListItem Text="112" Value="112"></asp:ListItem>
                    </asp:DropDownList>
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
                                <asp:ImageButton ID="IB_export" runat="server"  OnClick="IB_export_Click"
                        ImageUrl="~/images/icon/export.png" ToolTip="匯出Excel"  />

                                <asp:ImageButton ID="IB_import" runat="server"
                                    ImageUrl="~/images/icon/insert.png" ToolTip="匯入Excel" />
                            </td>
                            <td style="width: 28%; text-align: right;">
                                <asp:Label ID="Label2" runat="server" Font-Bold="False" Style="text-align: center" Font-Size="Medium" />
                            </td>
                        </tr>
                        <tr class="style8">
                            <td colspan="3">
                                <asp:ImageButton ID="IB_GV1_First1" runat="server" ImageUrl="~/images/First.gif" PostBackUrl="com_check.aspx" ToolTip="第一頁" />
                                <asp:ImageButton ID="IB_GV1_Previous1" runat="server" ImageUrl="~/images/Previous.gif" PostBackUrl="com_check.aspx" ToolTip="上一頁" />
                                <asp:ImageButton ID="IB_GV1_Next1" runat="server" ImageUrl="~/images/Next.gif" PostBackUrl="com_check.aspx" ToolTip="下一頁" />
                                <asp:ImageButton ID="IB_GV1_Last1" runat="server" ImageUrl="~/images/Last.gif" PostBackUrl="com_check.aspx" ToolTip="最後頁" />
                                為確保查詢效能，預設只顯示當年度資料。
                    <asp:GridView ID="GridView1" runat="server" AllowPaging="True"
                        AutoGenerateColumns="False" PageSize="20"
                        BackColor="White" BorderColor="#CCCCCC"
                        BorderStyle="None" BorderWidth="1px" CellPadding="3">
                        <RowStyle CssClass="grid4" ForeColor="#000066" />
                        <Columns>
                            <%--   <asp:TemplateField ShowHeader="False" ItemStyle-Width="10%">
                                <ItemTemplate >
                                    <asp:ImageButton ID="ImageButton2" runat="server" CausesValidation="false" CommandName="delete"
                                        ImageUrl="~/images/icon/delete.gif" OnClientClick="return confirm('確定要刪除資料?');"
                                        Text="按鈕" ToolTip="刪除" PostBackUrl="Default.aspx" />
                                </ItemTemplate>
                                <HeaderStyle CssClass="grid1" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>--%>
                            
                              <asp:TemplateField HeaderText="事業單位名稱" SortExpression="com_name">
                                <ItemTemplate>
                                    <asp:label ID="com_name" runat="server" Text='<%# Bind("com_name")%>'></asp:label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="第一季" SortExpression="sea1">
                                <ItemTemplate>
                                    <asp:TextBox ID="sea1" runat="server" Text='<%# Bind("sea1")%>' ></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField>
                             <asp:TemplateField HeaderText="第二季" SortExpression="sea2">
                                <ItemTemplate>
                                    <asp:TextBox ID="sea2" runat="server" Text='<%# Bind("sea2")%>' ></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField> <asp:TemplateField HeaderText="第三季" SortExpression="sea3">
                                <ItemTemplate>
                                    <asp:TextBox ID="sea3" runat="server" Text='<%# Bind("sea3")%>' ></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField> <asp:TemplateField HeaderText="第四季" SortExpression="sea4">
                                <ItemTemplate>
                                    <asp:TextBox ID="sea4" runat="server" Text='<%# Bind("sea4")%>' ></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateField>
                             <asp:TemplateField  Visible="false">
                                <ItemTemplate>
                                    <asp:label ID="cid" runat="server" Text='<%# Bind("cid")%>'></asp:label>
                                </ItemTemplate>
                            </asp:TemplateField>
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
                                <asp:ImageButton ID="IB_GV1_First2" runat="server" ImageUrl="~/images/First.gif" PostBackUrl="com_check.aspx" ToolTip="第一頁" />
                                <asp:ImageButton ID="IB_GV1_Previous2" runat="server" ImageUrl="~/images/Previous.gif" PostBackUrl="com_check.aspx" ToolTip="上一頁" />
                                <asp:ImageButton ID="IB_GV1_Next2" runat="server" ImageUrl="~/images/Next.gif" PostBackUrl="com_check.aspx" ToolTip="下一頁" />
                                <asp:ImageButton ID="IB_GV1_Last2" runat="server" ImageUrl="~/images/Last.gif" PostBackUrl="com_check.aspx" ToolTip="最後頁" />
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

                </td>
            </tr>
        </table>
    </form>
</body>
</html>
