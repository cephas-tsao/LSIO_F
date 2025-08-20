<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="basic_LSI026_Default" EnableViewStateMac="true" ViewStateEncryptionMode="Always" MaintainScrollPositionOnPostback="true" EnableEventValidation = "false" %>
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
</SCRIPT>
</head>
<body >
    <form id="form1" runat="server">
			 <div id="anticfrs" runat="server"></div>
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
                    名稱</td>
                <td class="style2">
                    <asp:TextBox ID="txt_eng_name" runat="server" MaxLength="50" Width="300px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    地址-區</td>
                <td class="style2">
                    <asp:DropDownList ID="ddl_eng_area" runat="server" AutoPostBack="True">
                    </asp:DropDownList></td>
            </tr>
            <tr>
                <td class="style1">
                    地址-路段</td>
                <td class="style2">
                    <asp:DropDownList ID="ddl_eng_road" runat="server">
                    </asp:DropDownList></td>
            </tr>

            <tr>
                <td class="style1">
                    通報單位1</td>
                <td class="style2">
                    <asp:TextBox ID="txt_com_name" runat="server" MaxLength="50" Width="300px"></asp:TextBox>
                                </td>
            </tr>

             <tr>
                <td class="style1">
                    通報日期</td>
                <td class="style2">
                    <asp:TextBox ID="txt_spc_date" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                    <asp:ImageButton ID="img_date1" runat="server" 
                        ImageUrl="~/images/calendr.gif" />
                     <font color="red">~<asp:TextBox ID="txt_epc_date" runat="server" 
                        MaxLength="10" Width="100px"></asp:TextBox><asp:ImageButton ID="img_date2" runat="server" 
                        ImageUrl="~/images/calendr.gif" />
                     Ex:yyyy/MM/dd</font></td>
            </tr>
            <tr style="display:none;">
                <td class="style1">
                    預計開工日</td>
                <td class="style2">
                    <asp:TextBox ID="txt_spc_date_s" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                    <asp:ImageButton ID="img_date3" runat="server" 
                        ImageUrl="~/images/calendr.gif" />
                     <font color="red">~<asp:TextBox ID="txt_epc_date_s" runat="server" 
                        MaxLength="10" Width="100px"></asp:TextBox><asp:ImageButton ID="img_date4" runat="server" 
                        ImageUrl="~/images/calendr.gif" />
                     Ex:yyyy/MM/dd</font></td>
            </tr>
         <tr>
              
             <td class="style1">預計完工日</td>
             <td class="style2">
                 <asp:TextBox ID="txt_spc_date_e" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                 <asp:ImageButton ID="img_date5" runat="server" ImageUrl="~/images/calendr.gif" />
                 <font color="red">~<asp:TextBox ID="txt_epc_date_e" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                 <asp:ImageButton ID="img_date6" runat="server" ImageUrl="~/images/calendr.gif" />
                 Ex:yyyy/MM/dd</font></td>
            <tr>
                <td class="style1">丁類危險性工作場所主要危害作業</td>
                <td class="style2">
                    <asp:CheckBoxList ID="ckl_pc_kind1" runat="server" Height="16px" RepeatColumns="2" RepeatDirection="Horizontal">
                    </asp:CheckBoxList>
                </td>
            </tr>
            <tr>
                <td class="style1">機械作業</td>
                <td class="style2">
                    <asp:CheckBoxList ID="ckl_pc_kind2" runat="server" Height="16px" RepeatColumns="2" RepeatDirection="Horizontal">
                    </asp:CheckBoxList>
                </td>
            </tr>

       

          
        </table>
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
            &nbsp;<asp:CheckBox ID="ckb_all" runat="server" Text="全部欄位" />
                </td>
            <td class="export02" style="width: 5%;"><asp:Button ID="bt_excel" runat="server" Text="匯出Excel" PostBackUrl="Default.aspx" />
            </td></tr></table>
        </asp:Panel>       
        <table width="100%" border ="1">
            <tr class="style8">
                <td style="width: 30%">
                    <asp:DropDownList
                        ID="ddlField" runat="server" DataSourceID="SqlDS2" DataTextField="field_name"
                        DataValueField="field_code">
                    </asp:DropDownList>
                    <asp:TextBox ID="TxtFind" runat="server" CssClass="findtb"></asp:TextBox>
                    <asp:Button ID="Button1" runat="server" ToolTip="查詢" Text="查詢" PostBackUrl="Default.aspx" />
                    <!--<asp:Button ID="Bt_3Days" runat="server" ToolTip="三日內" Text="三日內" PostBackUrl="Default.aspx" />-->
                </td>
                <td style="text-align:center">                        
                    <!-- <asp:ImageButton ID="ImageButton1" runat="server" CommandArgument="ADD" PostBackUrl="guidance.aspx" 
                        ImageUrl="~/images/icon/post.png" ToolTip="宣導事項" /> -->
                    　<asp:ImageButton ID="Button4" runat="server" CommandArgument="ADD"
                        ImageUrl="~/images/icon/add.png" ToolTip="新增" />
                    　<asp:ImageButton ID="Bt_subMenu" runat="server" CommandArgument="subMenu" 
                        ImageUrl="~/images/icon/search_ad.png" ToolTip="進階查詢" PostBackUrl="Default.aspx" />
                      <asp:ImageButton ID="IB_export" runat="server" 
                        ImageUrl="~/images/icon/export.png" ToolTip="匯出Excel" PostBackUrl="Default.aspx" />
                    　<asp:ImageButton ID="Bt_ExSearch" runat="server" CommandArgument="subMenu" 
                        ImageUrl="~/images/icon/search_ad5.png" ToolTip="進階查詢2" PostBackUrl="Default.aspx" />                    　
                    　<asp:ImageButton ID="IB_import" runat="server" 
                        ImageUrl="~/images/icon/insert.png" ToolTip="匯入Excel" />
                     <asp:ImageButton ID="IB_company" Width="40px" runat="server" PostBackUrl="Company.aspx" 
                        ImageUrl="~/images/icon/company.png" ToolTip="匯入公司資料" />
                    <asp:ImageButton ID="supload" Width="40px" runat="server" PostBackUrl="Supload.aspx" 
                        ImageUrl="~/images/icon/Modify.png" ToolTip="編輯自主檢查表" />
                    
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
					為確保查詢效能，預設只顯示當年度資料。
                    <asp:GridView ID="GridView1" runat="server" AllowPaging="True" 
                        AutoGenerateColumns="False"
                        DataSourceID="SqlDS1" BackColor="White" BorderColor="#CCCCCC" 
                        BorderStyle="None" BorderWidth="1px" CellPadding="3">
                        <RowStyle CssClass="grid4" ForeColor="#000066" />
                        <Columns>
                            <asp:TemplateField ShowHeader="False" HeaderStyle-Width="10%">
                                <ItemTemplate>
                                    <asp:ImageButton ID="ImageButton2" runat="server" CausesValidation="false" CommandName="delete"
                                        ImageUrl="~/images/icon/delete.gif" OnClientClick="return confirm('確定要刪除資料?');"
                                        Text="按鈕" ToolTip="刪除" PostBackUrl="Default.aspx" />
                                </ItemTemplate>
                                <HeaderStyle CssClass="grid1" />
                                <ItemStyle HorizontalAlign="Center" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderStyle-Width="10%">
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
