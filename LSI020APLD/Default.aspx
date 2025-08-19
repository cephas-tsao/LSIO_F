<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="basic_LSI020A_Default" MaintainScrollPositionOnPostback="true" EnableEventValidation = "false" %>
<%@ Register src="../../usercontrol/WebUserMenu.ascx" tagname="WebUserMenu" tagprefix="uc1" %>
<%@ Register src="../../usercontrol/WebSubMenu.ascx" tagname="WebSubMenu" tagprefix="uc2" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
<title><%=ConfigurationManager.AppSettings("Prj_name")%></title>
<link rel="stylesheet" href="https://pro.fontawesome.com/releases/v5.10.0/css/all.css" integrity="sha384-AYmEC3Yw5cVb3ZcuHtOA93w35dYTsvhLPVnYs9eStHfGJvOvKxVfELGroGkvsg+p" crossorigin="anonymous"/><link rel='stylesheet' href ="/../Setting/style<%=Get_SysPara("css_sys", "01")%>.css" type = 'text/css'/>
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
<script language="javascript">
function Chk_All_1() {
    for (var i = 0; i < document.form1.length; i++) {
        if (document.form1.elements[i].type == 'checkbox') {
            if (document.form1.elements[i].name.charAt(0) == 'C') {
                if (document.forms[0].elements["CBL_Field1_0"].checked == true) {
                    document.form1.elements[i].checked = true;
                }
                else {
                    document.form1.elements[i].checked = false;
                }
            }
        }
    }
}

function Chk_All_2() {
    for (var i = 0; i < document.form1.length; i++) {
        if (document.form1.elements[i].type == 'checkbox') {
            if (document.form1.elements[i].name.charAt(10) == 'e') {
                if (document.forms[0].elements["Area_Field_0"].checked == true) {
                    document.form1.elements[i].checked = true;
                }
                else {
                    document.form1.elements[i].checked = false;
                }
            }
        }
    }
}
</SCRIPT>
</head>
<body >
    <form id="form1" runat="server">
     <table id="customers" style="position:relative; margin-left:auto; margin-right:auto; width:990px;"  border ="0">
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
            <asp:Label ID="Label1" runat="server" style="text-align: center" Font-Bold="True" />
        </div>
        <uc2:WebSubMenu ID="WebSubMenu" runat="server" />
        <div>
        <asp:Panel ID="Panel_ExSearch" runat="server" HorizontalAlign="Left" Visible="False">
        <table width="100%" border ="1">
        <tr><td class="fail" colspan="2" style="text-align: center"><font color="blue">請輸入查詢條件：</font></td></tr>
              <tr>
                <td class="style1">
                    施工地點</td>
                <td class="style2">
                    <input id="Area_Field_0" type="checkbox" name="Area_Field_0" onclick="Chk_All_2();" /><label for="Area_Field_0">全選</label><br>
                    <asp:CheckBoxList ID="ckl_att_area" runat="server" RepeatColumns="4" RepeatDirection="Horizontal"
                        Height="16px">
                    </asp:CheckBoxList>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    學校</td>
                <td class="style2">
                    區域:<asp:DropDownList ID="ddl_att_area" runat="server" AutoPostBack="True">
                    </asp:DropDownList>
                    校名:<asp:DropDownList ID="ddl_com_name" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
           
            <tr>
                <td class="style1">
                    開工日期-起迄</td>
                <td class="style2">
                    <asp:TextBox ID="txt_work_date_s_s" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                    <asp:ImageButton ID="img_date1" runat="server" 
                        ImageUrl="~/images/calendr.gif" />
                     <font color="red">~<asp:TextBox ID="txt_work_date_s_e" runat="server" 
                        MaxLength="10" Width="100px"></asp:TextBox><asp:ImageButton ID="img_date2" runat="server" 
                        ImageUrl="~/images/calendr.gif" />
                     Ex:2011/01/01</font></td>
            </tr>
            <tr>
                <td class="style1">
                    完工日期-起迄</td>
                <td class="style2">
                    <asp:TextBox ID="txt_work_date_e_s" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                    <asp:ImageButton ID="img_date3" runat="server" 
                        ImageUrl="~/images/calendr.gif" />
                     <font color="red">~<asp:TextBox ID="txt_work_date_e_e" runat="server" 
                        MaxLength="10" Width="100px"></asp:TextBox><asp:ImageButton ID="img_date4" runat="server" 
                        ImageUrl="~/images/calendr.gif" />
                     Ex:2011/01/01</font></td>
            </tr>
             <tr>
                <td class="style1">
                    通報日期-起迄</td>
                <td class="style2">
                    <asp:TextBox ID="txt_ins_date_s" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                    <asp:ImageButton ID="img_date5" runat="server" 
                        ImageUrl="~/images/calendr.gif" />
                     <font color="red">~<asp:TextBox ID="txt_ins_date_e" runat="server" 
                        MaxLength="10" Width="100px"></asp:TextBox><asp:ImageButton ID="img_date6" runat="server" 
                        ImageUrl="~/images/calendr.gif" />
                     Ex:2011/01/01</font></td>
            </tr>
             <tr>
                <td class="style1">
                   承攬金額超過</td>
                <td class="style2">
                    <asp:TextBox ID="txt_work_money" runat="server" MaxLength="20" Width="100px"></asp:TextBox> 元
                   </td>
            </tr>
            <tr>
                <td class="style1">
                    作業樓層</td>
                <td class="style2">
                    <asp:CheckBoxList ID="ckl_work_floor_type" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"
                        Height="16px">
                    </asp:CheckBoxList></td>
            </tr>
            <tr>
                <td class="style1">
                    作業種類</td>
                <td class="style2">
                    <asp:CheckBoxList ID="ckl_fix_type" runat="server" RepeatColumns="4" RepeatDirection="Horizontal"
                        Height="16px">
                    </asp:CheckBoxList></td>
            </tr></table>
        <div style="text-align:center">
        <asp:Button ID="Bt_ExSearch_ok" runat="server" Text="確定" />
        <asp:Button ID="Bt_ExSearch_no" runat="server" Text="隱藏" />
        </div>
        </asp:Panel> 
        <asp:Panel ID="Panel1" runat="server" HorizontalAlign="Left" Visible="False">
            <table width="100%" border ="1"><tr><td class="export01" style="width: 14%;text-align:right"><font color="blue">請選擇匯出欄位：</font></td>
            <td class="export02"><asp:CheckBoxList ID="CBL_Field1" runat="server" 
                    RepeatDirection="Horizontal" RepeatColumns="6">
            </asp:CheckBoxList>
            &nbsp;<asp:CheckBox ID="ckb_all" runat="server" Text="全部欄位(包括未列出欄位)" />
                </td>
            <td class="export02" style="width: 5%;"><asp:Button ID="bt_excel" runat="server" Text="匯出Excel" PostBackUrl="Default.aspx" />
            </td></tr></table>
        </asp:Panel>       
        <table width="100%" border ="1">
            <tr class="style8">
                <td style="width: 42%">
                    <asp:DropDownList
                        ID="ddlField" runat="server" DataSourceID="SqlDS2" DataTextField="field_name"
                        DataValueField="field_code">
                    </asp:DropDownList>
                    <asp:TextBox ID="TxtFind" runat="server" CssClass="findtb"></asp:TextBox>
                    <asp:Button ID="Button1" runat="server" ToolTip="查詢" Text="查詢" PostBackUrl="Default.aspx" />
                </td>
                <td style="text-align:center">                        
                    　<asp:ImageButton ID="Button4" runat="server" CommandArgument="ADD"
                        ImageUrl="~/images/icon/add.png" ToolTip="新增一筆校園通報資料" />
                    　<asp:ImageButton ID="Bt_subMenu" runat="server" CommandArgument="subMenu" 
                        ImageUrl="~/images/icon/search_ad.png" ToolTip="進階查詢" PostBackUrl="Default.aspx" />
                    　<asp:ImageButton ID="Bt_ExSearch" runat="server" CommandArgument="subMenu" 
                        ImageUrl="~/images/icon/search_ad5.png" ToolTip="進階查詢2" PostBackUrl="Default.aspx" />
                    　<asp:ImageButton ID="IB_export" runat="server" 
                        ImageUrl="~/images/icon/export.png" ToolTip="匯出Excel" PostBackUrl="Default.aspx" />                    　
                </td>
                <td style="width: 28%; text-align: right;">
                    <asp:Label ID="Label2" runat="server" Font-Bold="False" Style="text-align: center" Font-Size="Medium" />
                </td>
            </tr>
            <tr class="style8">
                <td colspan="3">
                    <asp:ImageButton ID="IB_GV1_First1" runat="server" ImageUrl="~/images/First.gif" PostBackUrl="Default.aspx" ToolTip="第一頁" />
                    <asp:ImageButton ID="IB_GV1_Previous1" runat="server" ImageUrl="~/images/Previous.gif" PostBackUrl="Default.aspx" ToolTip="上一頁" />
                    <asp:ImageButton ID="IB_GV1_Next1" runat="server" ImageUrl="~/images/Next.gif" PostBackUrl="Default.aspx" ToolTip="下一頁" />
                    <asp:ImageButton ID="IB_GV1_Last1" runat="server" ImageUrl="~/images/Last.gif" PostBackUrl="Default.aspx" ToolTip="最後頁" />
                    <asp:GridView ID="GridView1" runat="server" AllowPaging="True" 
                        AutoGenerateColumns="False"
                        DataSourceID="SqlDS1" BackColor="White" BorderColor="#CCCCCC" 
                        BorderStyle="None" BorderWidth="1px" CellPadding="3">
                        <RowStyle CssClass="grid4" ForeColor="#000066" />
                        <Columns>
                            <asp:TemplateField ShowHeader="False">
                                <ItemTemplate>
                                    <asp:ImageButton ID="ImageButton2" runat="server" CausesValidation="false" CommandName="delete"
                                        ImageUrl="~/images/icon/delete.gif" OnClientClick="return confirm('確定要刪除資料?');"
                                        Text="按鈕" ToolTip="刪除" PostBackUrl="Default.aspx" />
                                </ItemTemplate>
                                <HeaderStyle CssClass="grid1" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>
                            <asp:TemplateField>
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
                            </asp:TemplateField>
                        </Columns>
                        <FooterStyle CssClass="grid1" ForeColor="#000066" BackColor="White" />
                        <SelectedRowStyle CssClass="grid8" BackColor="#669999" Font-Bold="True" 
                            ForeColor="White"  />
                        <HeaderStyle CssClass="grid3" Font-Bold="True" BackColor="#006699" 
                            ForeColor="White" />
                        <PagerStyle CssClass="grid5" HorizontalAlign="Left" BackColor="White" 
                            ForeColor="#000066" />
                        <EditRowStyle CssClass="grid6" />
                        <AlternatingRowStyle CssClass="grid7" />
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
