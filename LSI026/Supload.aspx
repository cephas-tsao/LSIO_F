<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Supload.aspx.vb" Inherits="basic_LSI026_Supload" %>

<%@ Register Src="../../usercontrol/WebUserMenu.ascx" TagName="WebUserMenu" TagPrefix="uc1" %>
<%@ Register Src="../../usercontrol/WebSubMenu.ascx" TagName="WebSubMenu" TagPrefix="uc2" %>
<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title><%=ConfigurationManager.AppSettings("Prj_name")%></title>
    <link rel="stylesheet" href="https://pro.fontawesome.com/releases/v5.10.0/css/all.css" integrity="sha384-AYmEC3Yw5cVb3ZcuHtOA93w35dYTsvhLPVnYs9eStHfGJvOvKxVfELGroGkvsg+p" crossorigin="anonymous" />
    <link rel='stylesheet' href="/../Setting/style<%=Get_SysPara("css_sys", "01")%>.css" type='text/css' />
</head>
<body>
    <form id="form1" runat="server">
        	 <div id="anticfrs" runat="server"></div>
        <table id="customers" style="position: relative; margin-left: auto; margin-right: auto; width: 990px;" border="0">
            <tr>
                <td width="100%">
                    <div>
                        <uc1:WebUserMenu ID="WebUserMenu" runat="server" />
                        <asp:Panel ID="Panel1" runat="server" HorizontalAlign="Left" Visible="False">
                            <table width="100%" border="1">
                                <tr>
                                    <td class="fail" colspan="2" style="text-align: center"><font color="blue">新增督導事項：</font></td>
                                </tr>
                                <tr>
                                    <td class="style1">項目</td>
                                    <td class="style2">
                                        <asp:DropDownList runat="server" ID="group_name" RepeatDirection="Horizontal" RepeatColumns="2"
                                            Height="30px">
                                        </asp:DropDownList>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style1">督導事項</td>
                                    <td class="style2">
                                        <asp:TextBox ID="txt_Item_name" runat="server" MaxLength="50" Width="300px"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style1"></td>
                                    <td class="style2">
                                        <asp:Button runat="server" Text="存檔" ID="save" OnClick="save_Click" />
                                        <asp:Button runat="server" ID="exit" Text="離開" OnClick="exit_Click" />
                                    </td>
                                </tr>
                            </table>
                        </asp:Panel>
                        <asp:Panel ID="Panel2" runat="server" HorizontalAlign="Left" Visible="False">
                            <table width="100%" border="1">
                                <tr>
                                    <td class="fail" colspan="2" style="text-align: center"><font color="blue">新增項目：</font></td>
                                </tr>                            
                                <tr>
                                    <td class="style1">項目內容</td>
                                    <td class="style2">
                                        <asp:TextBox ID="TextGroup" runat="server" MaxLength="50" Width="300px"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="style1"></td>
                                    <td class="style2">
                                        <asp:Button runat="server" Text="存檔" ID="SaveGroup" OnClick="SaveGroup_Click"/>
                                        <asp:Button runat="server" ID="eixt_group" Text="離開" OnClick="eixt_group_Click" />
                                    </td>
                                </tr>
                            </table>
                        </asp:Panel>
                    </div>
                </td>
            </tr>

            <tr>
                <td>
                    <div align="center" style="width: 990px; border: 1px #000000 solid; margin: 0px auto;">
                        <h3>自主輔導檢查申報自訂編輯</h3>
                        <table width="100%" border="1">
                            <tr>
                                <%-- <td width="30%" align="center">
                                    <label id="lblSites" runat="server" class="lblFieldDesc">請選擇單位:</label>
                                    <asp:DropDownList ID="drpSites" runat="server" AutoPostBack="True"
                                        OnSelectedIndexChanged="drpSites_SelectedIndexChanged">
                                        <asp:ListItem>一般科</asp:ListItem>
                                        <asp:ListItem>勞條科</asp:ListItem>
                                    </asp:DropDownList>
                                </td>--%>
                                <td style="text-align: center">
                    
                                    <asp:ImageButton ID="AddGroup" Width="40px" runat="server"   OnClick="AddGroup_Click"
                                        ImageUrl="~/images/icon/Modify.png" ToolTip="新增項目" />
                                    <asp:ImageButton ID="Button4" runat="server" CommandArgument="ADD" OnClick="Button4_Click"
                                        ImageUrl="~/images/icon/add.png" ToolTip="新增督導事項" />
                                    

                                </td>
                                <td style="width: 28%; text-align: right;">
                                    <asp:Label ID="Label2" runat="server" Font-Bold="False" Style="text-align: center" Font-Size="Medium" />
                                </td>
                            </tr>
                            <tr class="style8">
                                <td colspan="3">
                                    <asp:ImageButton ID="IB_GV1_First1" runat="server" ImageUrl="~/images/First.gif" PostBackUrl="Supload.aspx" ToolTip="第一頁" />
                                    <asp:ImageButton ID="IB_GV1_Previous1" runat="server" ImageUrl="~/images/Previous.gif" PostBackUrl="Supload.aspx" ToolTip="上一頁" />
                                    <asp:ImageButton ID="IB_GV1_Next1" runat="server" ImageUrl="~/images/Next.gif" PostBackUrl="Supload.aspx" ToolTip="下一頁" />
                                    <asp:ImageButton ID="IB_GV1_Last1" runat="server" ImageUrl="~/images/Last.gif" PostBackUrl="Supload.aspx" ToolTip="最後頁" />

                                    <asp:GridView ID="GridView1" runat="server" AllowPaging="True"
                                        AutoGenerateColumns="False" OnRowDataBound="GridView1_DataBound"
                                        DataSourceID="SqlDS1" BackColor="White" BorderColor="#CCCCCC" OnRowUpdating="GridView1_RowUpdating"
                                        BorderStyle="None" BorderWidth="1px" CellPadding="3">
                                        <RowStyle CssClass="grid4" ForeColor="#000066" />
                                        <Columns>
                                            <asp:TemplateField ItemStyle-Width="5%">
                                                <ItemTemplate>
                                                    <asp:ImageButton ID="ImageButton3" runat="server" CausesValidation="false"
                                                        CommandName="edit" ImageUrl="~/images/icon/edit.gif" CommandArgument="<%# Container.DataItemIndex %>"
                                                        Text="按鈕" ToolTip="修改" />
                                                </ItemTemplate>
                                                <EditItemTemplate>
                                                    <asp:Button ID="Button2" runat="server"
                                                        CommandName="Update" Text="更新" CommandArgument='<%# Container.DataItemIndex%>' />

                                                    <asp:Button ID="lbCancelUpdate" runat="server"
                                                        CommandName="Cancel" Text="取消" CommandArgument='<%# Container.DataItemIndex%>' />
                                                </EditItemTemplate>
                                                <HeaderStyle CssClass="grid1" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="項目" ItemStyle-Width="20%">
                                                <ItemTemplate>
                                                    <asp:Label runat="server" ID="group_name" Text='<%# Bind("group_name")%>'></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="督導事項">
                                                <ItemTemplate>
                                                    <asp:Label runat="server" ID="content" Text='<%# Bind("content")%>'></asp:Label>
                                                </ItemTemplate>
                                                <EditItemTemplate>
                                                    <asp:Textbox runat="server" Width="100%" Height="100px" ID="content" Text='<%# Bind("content")%>' ></asp:Textbox>
                                                </EditItemTemplate>
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="是否使用" ItemStyle-Width="10%">
                                                <ItemTemplate>
                                                    <asp:Label runat="server" ID="is_enabled" Text='<%# Bind("item_enabled")%>'></asp:Label>
                                                </ItemTemplate>
                                                <EditItemTemplate>
                                                    <asp:DropDownList runat="server" ID="is_enabled">
                                                        <asp:ListItem Text="是" Value="True"></asp:ListItem>
                                                        <asp:ListItem Text="否" Value="False"></asp:ListItem>
                                                    </asp:DropDownList>
                                                    <%--  <asp:Textbox runat="server" ID="is_enabled" Text='<%# Bind("is_enabled") %>'></asp:Textbox>--%>
                                                </EditItemTemplate>
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
                                    <asp:ImageButton ID="IB_GV1_First2" runat="server" ImageUrl="~/images/First.gif" PostBackUrl="Supload.aspx" ToolTip="第一頁" />
                                    <asp:ImageButton ID="IB_GV1_Previous2" runat="server" ImageUrl="~/images/Previous.gif" PostBackUrl="Supload.aspx" ToolTip="上一頁" />
                                    <asp:ImageButton ID="IB_GV1_Next2" runat="server" ImageUrl="~/images/Next.gif" PostBackUrl="Supload.aspx" ToolTip="下一頁" />
                                    <asp:ImageButton ID="IB_GV1_Last2" runat="server" ImageUrl="~/images/Last.gif" PostBackUrl="Supload.aspx" ToolTip="最後頁" />
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
                    </div>
                </td>
            </tr>

        </table>
    </form>
</body>
</html>
