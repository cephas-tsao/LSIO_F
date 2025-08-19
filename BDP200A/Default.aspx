<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="management_BDP200A_Default" MaintainScrollPositionOnPostback="true" %>
<%@ Register src="../../usercontrol/WebUserMenu.ascx" tagname="WebUserMenu" tagprefix="uc1" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title><%=ConfigurationManager.AppSettings("Prj_name")%></title>
    <link rel="stylesheet" href="https://pro.fontawesome.com/releases/v5.10.0/css/all.css" integrity="sha384-AYmEC3Yw5cVb3ZcuHtOA93w35dYTsvhLPVnYs9eStHfGJvOvKxVfELGroGkvsg+p" crossorigin="anonymous"/><link rel='stylesheet' href ="/../Setting/style<%=Get_SysPara("css_sys", "01")%>.css" type = 'text/css'/>
<script language="javascript">
<!--
    function RecordDetail(id) {
        var detail = window.open('../../public/RecordDetail.aspx?id=' + id, 'RecordDetail', 'height=600,width=800,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no');
        detail.focus();
    }
//-->
</SCRIPT>
</head>
<body>
    <form id="form2" runat="server">
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

    <div style="text-align: center"><asp:Label ID="Label1" runat="server" style="text-align: center" Font-Bold="True"></asp:Label></div>
    <div>
        <table width="100%" border ="1">
            <tr class="style8">
                <td style="width: 70%">
                    <asp:DropDownList
                        ID="ddlField" runat="server" DataSourceID="SqlDS2" DataTextField="field_name"
                        DataValueField="field_code">
                    </asp:DropDownList>
                    <asp:TextBox ID="TxtFind" runat="server" CssClass="findtb"></asp:TextBox>
                    <asp:Button ID="Button1" runat="server" ToolTip="查詢" Text="查詢" />
                </td>
                <td style="width: 30%; text-align: right;">
                    <asp:Label ID="Label2" runat="server" Font-Bold="False" Style="text-align: center" Font-Size="Medium"></asp:Label>
                </td>
            </tr>
            <tr class="style8">
                <td colspan="2">
                    <asp:ImageButton ID="IB_GV1_First1" runat="server" ImageUrl="~/images/First.gif" PostBackUrl="Default.aspx" ToolTip="第一頁" />
                    <asp:ImageButton ID="IB_GV1_Previous1" runat="server" ImageUrl="~/images/Previous.gif" PostBackUrl="Default.aspx" ToolTip="上一頁" />
                    <asp:ImageButton ID="IB_GV1_Next1" runat="server" ImageUrl="~/images/Next.gif" PostBackUrl="Default.aspx" ToolTip="下一頁" />
                    <asp:ImageButton ID="IB_GV1_Last1" runat="server" ImageUrl="~/images/Last.gif" PostBackUrl="Default.aspx" ToolTip="最後頁" />
                    <asp:GridView ID="GridView1" runat="server" AllowPaging="True" 
                        AutoGenerateColumns="False" CssClass="grid9"
                        DataSourceID="SqlDS1">
                        <RowStyle cssclass="grid4" />
                        <Columns>
                            <asp:TemplateField ShowHeader="False">
                                <ItemTemplate>
                                    <asp:ImageButton ID="ImageButton2" runat="server" CausesValidation="false" CommandName="delete"
                                        ImageUrl="~/images/icon/delete.gif" OnClientClick="return confirm('確定要刪除資料?');" Text="按鈕"
                                        ToolTip="刪除" />
                                </ItemTemplate>
                                <HeaderStyle CssClass="grid1"  />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:ImageButton ID="ImageButton3" runat="server" CausesValidation="false" CommandArgument="MDY"
                                        ImageUrl="~/images/check.ico" Text="按鈕" ToolTip="詳細資料" />
                                </ItemTemplate>
                                <HeaderStyle cssclass="grid1" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>
                            <asp:TemplateField>
                                <ItemTemplate>
                                    <asp:ImageButton ID="ImageButton4" runat="server" CausesValidation="false" CommandArgument="COPY"
                                        CommandName="select" ImageUrl="~/images/icon/copy.png" PostBackUrl="edit.aspx" Text="按鈕"
                                        ToolTip="複製" Visible="False" />
                                </ItemTemplate>
                                 <HeaderStyle cssclass="grid1" Font-Bold="True" />
                                 
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>
                        </Columns>
                        <FooterStyle cssclass="grid1" />
                        <SelectedRowStyle cssclass="grid8" />
                        <HeaderStyle cssclass="grid3" />
                        <PagerStyle cssclass="grid5" />
                        <EditRowStyle cssclass="grid6" />
                        <AlternatingRowStyle cssclass="grid7" />
                    </asp:GridView>
                    <asp:ImageButton ID="IB_GV1_First2" runat="server" ImageUrl="~/images/First.gif" PostBackUrl="Default.aspx" ToolTip="第一頁" />
                    <asp:ImageButton ID="IB_GV1_Previous2" runat="server" ImageUrl="~/images/Previous.gif" PostBackUrl="Default.aspx" ToolTip="上一頁" />
                    <asp:ImageButton ID="IB_GV1_Next2" runat="server" ImageUrl="~/images/Next.gif" PostBackUrl="Default.aspx" ToolTip="下一頁" />
                    <asp:ImageButton ID="IB_GV1_Last2" runat="server" ImageUrl="~/images/Last.gif" PostBackUrl="Default.aspx" ToolTip="最後頁" />
                    <asp:Label ID="l_PageMsg" runat="server" />
                </td>
            </tr>
        </table>
        <asp:TextBox ID="txtSubWhere" runat="server" Visible="False" />
        <asp:TextBox ID="txtPrimary" runat="server" Visible="False" />
        <asp:TextBox ID="txtPgmType" runat="server" Visible="False" />
        <asp:SqlDataSource ID="SqlDS1" runat="server"></asp:SqlDataSource><asp:SqlDataSource ID="SqlDS2" runat="server" />
    </div>
     </td>
    </tr>
    </table>
    </form>
</body>
</html>
