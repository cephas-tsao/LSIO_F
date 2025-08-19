<%@ Page Language="VB" AutoEventWireup="false" CodeFile="edit.aspx.vb" Inherits="basic_OTB020A_edit" MaintainScrollPositionOnPostback="true" validateRequest="False" %>
<%@ Register src="../../usercontrol/WebUserMenu.ascx" tagname="WebUserMenu" tagprefix="uc1" %>
<%@ Register src="detail_edit.ascx" tagname="detail_edit" tagprefix="uc2" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title><%=ConfigurationManager.AppSettings("Prj_name")%></title>
    <script language="javascript" id="js1" src="../../public/checkjs.js"></script>    
    <link rel="stylesheet" href="https://pro.fontawesome.com/releases/v5.10.0/css/all.css" integrity="sha384-AYmEC3Yw5cVb3ZcuHtOA93w35dYTsvhLPVnYs9eStHfGJvOvKxVfELGroGkvsg+p" crossorigin="anonymous"/><link rel='stylesheet' href ="/../Setting/style<%=Get_SysPara("css_sys", "01")%>.css" type = 'text/css'/>
    </head>
<body>
    <form id="form1" runat="server">
    <table style="position:relative; margin-left:auto; margin-right:auto; width:990px; "  border ="0">
        <tr>
            <td width="100%">
             <div>
                <uc1:WebUserMenu ID="WebUserMenu" runat="server" />
             </div>
           </td>
        </tr>
        <tr>
        
        <td>
    <div>
        <table width="100%" border="1" id="customers" cellspacing='0.1'>
            <tr>
                <td  class="fail" colspan="2" style="text-align: center">
                    <asp:Label ID="Label1" runat="server" Font-Bold="True" Style="text-align: center"></asp:Label></td>
            </tr>
            <tr>
                <td class="style1">
                    留言板編號</td>
                <td class="style2" style="width: 100%">
                    <asp:TextBox ID="txt_bull_code" runat="server" Width="120px" CssClass="tb2" 
                        Enabled="False"></asp:TextBox>
                    <font color="red"><asp:Label ID="Label2" runat="server" 
                        CssClass="lab_mark" Text="系統自動取號"></asp:Label></font></td>
            </tr>
            <tr>
                <td class="style1">
                                        發表人</td>
                <td class="style2" style="width: 35%">
                    <asp:TextBox ID="txt_usr_name" runat="server" MaxLength="20" Width="200px" 
                        ></asp:TextBox>IP:<asp:TextBox ID="txt_usr_ip" runat="server" 
                        MaxLength="15" Width="150px"></asp:TextBox></td>
            </tr>
                                 
            <tr>
                <td class="style1">
                                        電子郵件</td>
                <td class="style2" style="width: 35%">
                    <asp:TextBox ID="txt_usr_mail" runat="server" MaxLength="100" Width="500px" 
                       ></asp:TextBox></td>
            </tr>
                                 
            <tr>
                <td class="style1">
                                        留言主題</td>
                <td class="style2" style="width: 35%">
                    <asp:TextBox ID="txt_theme" runat="server" MaxLength="100" Width="500px"></asp:TextBox>
                            </td>
            </tr>
            <tr>
                <td class="style1">
                                        發表日期</td>
                <td class="style2" style="width: 35%">
                    <asp:TextBox ID="txt_ins_date" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                    發表時間<asp:TextBox ID="txt_ins_time" runat="server" MaxLength="8" Width="80px"></asp:TextBox></td>
            </tr>                               
            <tr>
                <td class="style1">
                                        <span id="l_bull_con">留言內容</span></td>
                <td class="style2" style="width: 35%">
                    <asp:TextBox ID="txt_bull_con" runat="server" Width="500px" Height="100px" 
                        TextMode="MultiLine"></asp:TextBox></td>
            </tr> 
                                 
            <tr>
                <td class="style1">
                                        管理顯示</td>
                <td class="style2" style="width: 35%">
                    <asp:DropDownList ID="ddl_is_hid" runat="server" AutoPostBack="True">
                    </asp:DropDownList>
                </td>
            </tr>
        </table>
    </div>
            <asp:Button ID="Button1" runat="server" Text="主檔存檔" />
            <asp:Button ID="Bt_BackUp" runat="server" Text="離開" />
            <asp:TextBox ID="txtPgmType" runat="server" Visible="False" />
            <asp:TextBox ID="txtPrimary" runat="server" Visible="False" />
            <asp:TextBox ID="txtSubWhere" runat="server" Visible="False" />
            <asp:TextBox ID="txtOrder" runat="server" Visible="False" />
            <br />
            <asp:Label ID="l_ErrMsg" runat="server" Font-Bold="True" ForeColor="Red" />
            <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="True"
             ShowSummary="False" />
            </td>
        </tr>
        <tr>        
          <td>
          <br />
          <asp:Panel ID="Panel1" runat="server" Visible="False">
              <uc2:detail_edit ID="detail_edit" runat="server" />          
         </asp:Panel>
          </td>
        </tr>
      </table>    
    </form>
</body>
</html>
