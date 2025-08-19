<%@ Control Language="VB" AutoEventWireup="false" CodeFile="detail_edit.ascx.vb" Inherits="basic_LSI008A_detail_edit" %>
<table border ="0">
    <tr>
        <td class="fail2" style="text-align: center">
            <asp:Label ID="Label1" runat="server" Font-Bold="True" Style="text-align: center" />
            <asp:Label ID="l_mode" runat="server" />
        </td>
    </tr>
    <tr>
        <td>
<asp:Panel ID="Panel1" runat="server" Visible="False">
<table width="100%" border="1" id="customers" cellspacing='0.1'>
    <tr>
        <td class="style1-1">
            資料編號<asp:Image ID="Image1" runat="server" ImageUrl="~/images/tick.ico" /></td>
        <td class="style2-1">
            <asp:TextBox ID="txt_serial_no" runat="server" Width="100px" Enabled="False" BackColor="#999999"></asp:TextBox>
            <font color="red">系統自動取號</font></td>
    </tr>
    <tr>
        <td class="style1-1">
            預約編號<asp:Image ID="Image2" runat="server" ImageUrl="~/images/tick.ico" /></td>
        <td class="style2-1">
            <asp:TextBox ID="txt_mrd_code" runat="server" Width="120px" Enabled="False" BackColor="#999999"></asp:TextBox>
            <font color="red">系統自動取值</font></td>
    </tr>
    <tr>
        <td class="style1-1">
            使用者編號</td>
        <td class="style2-1">
            <asp:TextBox ID="txt_usr_code" runat="server" Width="200px" BackColor="#CCCCCC"></asp:TextBox>
            <asp:ImageButton ID="IB_usr_code" runat="server" CausesValidation="False" ImageUrl="~/images/bs_search.gif" />            
        </td>
    </tr>
    <tr>
        <td class="style1-1">
            使用者名稱<asp:Image ID="Image3" runat="server" ImageUrl="~/images/tick.ico" /></td>
        <td class="style2-1">
            <asp:TextBox ID="txt_usr_name" runat="server" Width="200px" MaxLength="20"></asp:TextBox></td>
    </tr>
    <tr>
        <td class="style1-1">
            科室名稱</td>
        <td class="style2-1">
            <asp:TextBox ID="txt_usr_dep" runat="server" Width="200px" MaxLength="20"></asp:TextBox></td>
    </tr>
    <tr>
        <td class="style1-1">
            連絡mail</td>
        <td class="style2-1">
            <asp:TextBox ID="txt_usr_mail" runat="server" Width="500px" MaxLength="100"></asp:TextBox></td>
    </tr>      
    </table>
        <asp:Button ID="bt_Save" runat="server" Text="明細存檔" /> 
        <asp:Button ID="bt_Cancel" runat="server" Text="隱藏" /> 
        <br />
        <asp:Label ID="l_ErrMsg" runat="server" Font-Bold="True" ForeColor="Red" />
</asp:Panel>
</td>
</tr>
<tr>
<td class="style1-1">
<table width="100%" border ="1">
            <tr class="style8">
                <td style="width: 70%; text-align:center">
                        <asp:Button ID="Button4" runat="server" CommandArgument="ADD"
                        Text="新增人員" />
                        <asp:Button ID="bt_usr_list" runat="server" Text="大量新增" />
                        <asp:Button ID="bt_InFile" runat="server" Text="匯入檔案" />
                        <asp:Button ID="IB_export" runat="server" Text="匯出檔案" />
                  </td>
                <td style="width: 30%; text-align: right;">
                    <asp:Label ID="Label2" runat="server" Font-Bold="False" Style="text-align: center" Font-Size="Medium"></asp:Label>
                </td>
            </tr>
            <tr class="style8">
                <td colspan="2">
                    <asp:ImageButton ID="IB_GV1_First1" runat="server" ImageUrl="~/images/First.gif" ToolTip="第一頁" />
                    <asp:ImageButton ID="IB_GV1_Previous1" runat="server" ImageUrl="~/images/Previous.gif" ToolTip="上一頁" />
                    <asp:ImageButton ID="IB_GV1_Next1" runat="server" ImageUrl="~/images/Next.gif" ToolTip="下一頁" />
                    <asp:ImageButton ID="IB_GV1_Last1" runat="server" ImageUrl="~/images/Last.gif" ToolTip="最後頁" />
                    <asp:GridView ID="GridView1" runat="server" AllowPaging="True" 
                        AutoGenerateColumns="False"
                        DataSourceID="SqlDS1" EmptyDataText="尚未輸入資料">
                        <RowStyle CssClass="grid4" />
                        <EmptyDataRowStyle HorizontalAlign="Center" />
                        <Columns>
                            <asp:TemplateField ShowHeader="False">
                                <ItemTemplate>
                                    <asp:ImageButton ID="ImageButton2" runat="server" CausesValidation="false" CommandName="delete"
                                        ImageUrl="~/images/icon/delete.gif" OnClientClick="return confirm('確定要刪除資料?');"
                                        Text="按鈕" ToolTip="刪除" />
                                </ItemTemplate>
                                <HeaderStyle CssClass="grid1" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:ImageButton ID="ImageButton3" runat="server" CausesValidation="false" CommandArgument="MDY"
                                        CommandName="select" ImageUrl="~/images/icon/edit.gif" onclick="Edit_mode"
                                        Text="按鈕" ToolTip="修改" />&nbsp;
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
                            </asp:TemplateField>
                        </Columns>
                        <FooterStyle CssClass="grid1" Font-Bold="True" ForeColor="White" />
                        <SelectedRowStyle CssClass="grid8"  />
                        <HeaderStyle CssClass="grid3" Font-Bold="True" />
                        <PagerStyle CssClass="grid5" HorizontalAlign="Center" />
                        <EditRowStyle CssClass="grid6" />
                        <AlternatingRowStyle CssClass="grid7" />
                    </asp:GridView>
                    <asp:ImageButton ID="IB_GV1_First2" runat="server" ImageUrl="~/images/First.gif" ToolTip="第一頁" />
                    <asp:ImageButton ID="IB_GV1_Previous2" runat="server" ImageUrl="~/images/Previous.gif" ToolTip="上一頁" />
                    <asp:ImageButton ID="IB_GV1_Next2" runat="server" ImageUrl="~/images/Next.gif" ToolTip="下一頁" />
                    <asp:ImageButton ID="IB_GV1_Last2" runat="server" ImageUrl="~/images/Last.gif" ToolTip="最後頁" />
                    <asp:Label ID="l_PageMsg" runat="server" />
                </td>
            </tr>
         </table>
</td>
</tr>
</table>
<asp:TextBox ID="txtPrimary" runat="server" Visible="False" />
<asp:TextBox ID="txtPgmType" runat="server" Visible="False" />
<asp:TextBox ID="txtKeyName" runat="server" Visible="False" />
<asp:SqlDataSource ID="SqlDS1" runat="server" />
<asp:SqlDataSource ID="SqlDS2" runat="server" />