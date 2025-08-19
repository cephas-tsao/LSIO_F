<%@ Page Language="VB" AutoEventWireup="false" CodeFile="edit.aspx.vb" Inherits="basic_LSI008A_edit" aspcompat="true" MaintainScrollPositionOnPostback="true" EnableEventValidation = "false" %>
<%@ Register src="../../usercontrol/WebUserMenu.ascx" tagname="WebUserMenu" tagprefix="uc1" %>
<%@ Register src="detail_edit.ascx" tagname="detail_edit" tagprefix="uc2" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title><%=ConfigurationManager.AppSettings("Prj_name")%></title>
    <script language="javascript" id="js1" src="../../public/checkjs.js"></script> 
    <link rel="stylesheet" href="https://pro.fontawesome.com/releases/v5.10.0/css/all.css" integrity="sha384-AYmEC3Yw5cVb3ZcuHtOA93w35dYTsvhLPVnYs9eStHfGJvOvKxVfELGroGkvsg+p" crossorigin="anonymous"/><link rel='stylesheet' href ="/../Setting/style<%=Get_SysPara("css_sys", "01")%>.css" type = 'text/css'/>
    <style type="text/css">
        .style1
        {
            width: 13%;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <table style="position:relative; margin-left:auto; margin-right:auto; width:990px; "  border ="0">
        <tr>        
           <td>
    <div>
        <table width="100%" border="1" id="customers" cellspacing='0.1'>
            <tr>
                <td class="fail" colspan="2" style="text-align: center">
                    <asp:Label ID="Label1" runat="server" Font-Bold="True" Style="text-align: center"></asp:Label></td>
            </tr>
            <tr>
                <td class="style1">
                    資料編號<asp:Image ID="Image2" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_mrd_code" runat="server" MaxLength="12" Width="120px" BackColor="#999999"></asp:TextBox>
                    <font color="red">系統自動取號</font></td>
            </tr>
            <tr>
                <td class="style1">
                    租用會議室編號<asp:Image ID="Image3" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:DropDownList ID="ddl_room_code" runat="server">
                    </asp:DropDownList>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    租用日期</td>
                <td class="style2">
                    <asp:TextBox ID="txt_mrd_date" runat="server" MaxLength="10" Width="110px" BackColor="#CCCCCC"></asp:TextBox>                    
                    <asp:ImageButton ID="img_date1" runat="server" 
                        ImageUrl="~/images/calendr.gif" Width="16px" Visible="False" /></td>
            </tr>
            <tr>
                <td class="style1">
                    租用時間</td>
                <td class="style2">
                    自 <asp:DropDownList ID="ddl_mrd_time_s" runat="server" /> 起
                    至 <asp:DropDownList ID="ddl_mrd_time_e" runat="server"/> 止</td>
            </tr>
            <tr>
                <td class="style1">
                    租用者名稱</td>
                <td class="style2">
                    <asp:TextBox ID="txt_usr_code" runat="server" MaxLength="20" Width="120px"></asp:TextBox>
                    <asp:TextBox ID="txt_usr_name" runat="server" MaxLength="20" Width="120px" BackColor="#CCCCCC"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    會議名稱</td>
                <td class="style2">
                    <asp:TextBox ID="txt_mrd_name" runat="server" MaxLength="100" Width="500px"></asp:TextBox>
                    <font color="red">如果會議名稱有危評2個字，則在前台會查詢的到</font>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    主持人名字</td>
                <td class="style2">
                    <asp:TextBox ID="txt_mrd_host" runat="server" MaxLength="20" Width="200px"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="style1">
                    通知發送時間</td>
                <td class="style2">
                    <asp:CheckBoxList ID="ckl_send_list" runat="server" RepeatColumns="2" Width="400px">
                    </asp:CheckBoxList></td>
            </tr>
            <tr>
                <td class="style1">
                    是否為危評會議</td>
                <td class="style2">
                    <asp:DropDownList ID="ddl_is_show" runat="server">
                    </asp:DropDownList>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    關鍵字</td>
                <td class="style2">
                    <asp:TextBox ID="txt_mrd_memo" runat="server" Width="624px" Height="150px" 
                        TextMode="MultiLine"></asp:TextBox></td>
            </tr>              
        </table>
    </div>
            <asp:Button ID="Button1" runat="server" Text="存檔" />
            <asp:Button ID="Bt_BackUp" runat="server" Text="關閉視窗" />
            <asp:TextBox ID="txtPgmType" runat="server" Visible="False" />
            <asp:TextBox ID="txtSaved" runat="server" Visible="False" />
            <br />
            <asp:Label ID="l_ErrMsg" runat="server" Font-Bold="True" ForeColor="Red" />
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
