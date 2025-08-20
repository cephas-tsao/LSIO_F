<%@ Control Language="VB" AutoEventWireup="false" CodeFile="WebSubMenu.ascx.vb" Inherits="usercontrol_WebSubMenu" %>
<%  Dim p1 As New PageBase%>
<link rel='stylesheet' href ="/Setting/style<%=P1.Get_SysPara("css_sys", "01")%>.css" type = 'text/css'/>
<asp:Panel ID="Panel1" runat="server" Visible="False" HorizontalAlign="Center">      
        <table width="100%" border ="1" style= "border-collapse:collapse ">
            <tr>
                <td colspan="3" class="style9" >
                    <asp:Label ID="l_Order" runat="server" Font-Bold="True" Text="進階排序"></asp:Label>
                </td>
                <td colspan="4" class="style10">
                    <asp:Label ID="l_Search" runat="server" Font-Bold="True" Text="進階搜尋"></asp:Label>
                </td>
            </tr>
            <tr>
                <td class="style9" >
                    順位</td>
                <td class="style9" >
                    欄位名稱</td>
                <td class="style9" >
                    排序</td>
                <td class="style10">
                    欄位名稱</td>
                <td class="style10">
                    篩選</td>
                <td class="style10">
                    篩選條件</td>
                <td class="style10">
                    <asp:DropDownList ID="ddl_bool1" runat="server" Visible="False">
                    </asp:DropDownList>
                    &nbsp;</td>
            </tr>
            <tr>
                <td class="style9">
                    <asp:Label ID="l_order1" runat="server" Text="第一順位"></asp:Label>
                </td>
                <td class="style9">
                    <asp:DropDownList ID="ddl_field1" runat="server">
                    </asp:DropDownList>
                </td>
                <td class="style9">
                    <asp:DropDownList ID="ddl_order1" runat="server">
                    </asp:DropDownList>
                </td>
                <td class="style10">
                    <asp:DropDownList ID="ddl_search1" runat="server" DataTextField="field_name" 
                        DataValueField="field_code">
                    </asp:DropDownList>
                </td>
                <td class="style10">
                    <asp:DropDownList ID="ddl_Math_s1" runat="server">
                    </asp:DropDownList>
                </td>
                <td class="style10">
                    <asp:TextBox ID="txt_Search1" runat="server"></asp:TextBox>
                </td>
                <td class="style10">
                    <asp:DropDownList ID="ddl_bool2" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="style9">
                    <asp:Label ID="l_order2" runat="server" Text="第二順位"></asp:Label>
                </td>
                <td class="style9">
                    <asp:DropDownList ID="ddl_field2" runat="server">
                    </asp:DropDownList>
                </td>
                <td class="style9">
                    <asp:DropDownList ID="ddl_order2" runat="server">
                    </asp:DropDownList>
                </td>
                <td class="style10">
                    <asp:DropDownList ID="ddl_search2" runat="server" DataTextField="field_name"
                        DataValueField="field_code">
                    </asp:DropDownList>
                </td>
                <td class="style10">
                    <asp:DropDownList ID="ddl_Math_s2" runat="server">
                    </asp:DropDownList>
                </td>
                <td class="style10">
                    <asp:TextBox ID="txt_Search2" runat="server"></asp:TextBox>
                </td>
                <td class="style10">
                    <asp:DropDownList ID="ddl_bool3" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="style9">
                    <asp:Label ID="l_order3" runat="server" Text="第三順位"></asp:Label>
                </td>
                <td class="style9">
                    <asp:DropDownList ID="ddl_field3" runat="server">
                    </asp:DropDownList>
                </td>
                <td class="style9">
                    <asp:DropDownList ID="ddl_order3" runat="server">
                    </asp:DropDownList>
                </td>
                <td class="style10">
                    <asp:DropDownList ID="ddl_search3" runat="server" DataTextField="field_name"
                        DataValueField="field_code">
                    </asp:DropDownList>
                </td>
                <td class="style10">
                    <asp:DropDownList ID="ddl_Math_s3" runat="server">
                    </asp:DropDownList>
                </td>
                <td class="style10">
                    <asp:TextBox ID="txt_Search3" runat="server"></asp:TextBox>
                </td>
                <td class="style10">
                    <asp:DropDownList ID="ddl_bool4" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="style9">
                    <asp:Label ID="l_order4" runat="server" Text="第四順位"></asp:Label>
                </td>
                <td class="style9">
                    <asp:DropDownList ID="ddl_field4" runat="server">
                    </asp:DropDownList>
                </td>
                <td class="style9">
                    <asp:DropDownList ID="ddl_order4" runat="server">
                    </asp:DropDownList>
                </td>
                <td class="style10">
                    <asp:DropDownList ID="ddl_search4" runat="server" DataTextField="field_name"
                        DataValueField="field_code">
                    </asp:DropDownList>
                </td>
                <td class="style10">
                    <asp:DropDownList ID="ddl_Math_s4" runat="server">
                    </asp:DropDownList>
                </td>
                <td class="style10">
                    <asp:TextBox ID="txt_Search4" runat="server"></asp:TextBox>
                </td>
                <td class="style10">
                    <asp:DropDownList ID="ddl_bool5" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="style9">
                    <asp:Label ID="l_order5" runat="server" Text="第五順位"></asp:Label>
                </td>
                <td class="style9">
                    <asp:DropDownList ID="ddl_field5" runat="server">
                    </asp:DropDownList>
                </td>
                <td class="style9">
                    <asp:DropDownList ID="ddl_order5" runat="server">
                    </asp:DropDownList>
                </td>
                <td class="style10">
                    <asp:DropDownList ID="ddl_search5" runat="server" DataTextField="field_name"
                        DataValueField="field_code">
                    </asp:DropDownList>
                </td>
                <td class="style10">
                    <asp:DropDownList ID="ddl_Math_s5" runat="server">
                    </asp:DropDownList>
                </td>
                <td class="style10">
                    <asp:TextBox ID="txt_Search5" runat="server"></asp:TextBox>
                </td>
                <td class="style10">
                    <asp:DropDownList ID="ddl_bool6" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="style9">
                    <asp:Label ID="l_order6" runat="server" Text="第六順位"></asp:Label>
                </td>
                <td class="style9">
                    <asp:DropDownList ID="ddl_field6" runat="server">
                    </asp:DropDownList>
                </td>
                <td class="style9">
                    <asp:DropDownList ID="ddl_order6" runat="server">
                    </asp:DropDownList>
                </td>
                <td class="style10">
                    <asp:DropDownList ID="ddl_search6" runat="server" DataTextField="field_name"
                        DataValueField="field_code">
                    </asp:DropDownList>
                </td>
                <td class="style10">
                    <asp:DropDownList ID="ddl_Math_s6" runat="server">
                    </asp:DropDownList>
                </td>
                <td class="style10">
                    <asp:TextBox ID="txt_Search6" runat="server"></asp:TextBox>
                </td>
                <td class="style10">
                    &nbsp;</td>
            </tr>
            <%--<tr class="style2">
                <td colspan="7" style= "border-right:none ">
                    
                </td>
            </tr>--%>
        </table>
        <asp:Button ID="Bt_ok" runat="server" Text="確定" />
        <asp:Button ID="Bt_no" runat="server" Text="隱藏" />
        <asp:TextBox ID="txtGridViewName" runat="server" Visible="False"></asp:TextBox>
        <asp:TextBox ID="txtOrderTBName" runat="server" Visible="False"></asp:TextBox>
        <asp:TextBox ID="txtSearchTBName" runat="server" Visible="False"></asp:TextBox>
        <asp:TextBox ID="txtSqlDSName" runat="server" Visible="False"></asp:TextBox>
        </asp:Panel>
