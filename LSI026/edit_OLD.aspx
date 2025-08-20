<%@ Page Language="VB" AutoEventWireup="false" CodeFile="edit.aspx.vb" Inherits="basic_LSI026_edit" aspcompat="true" MaintainScrollPositionOnPostback="true" %>
<%@ Register src="../../usercontrol/WebUserMenu.ascx" tagname="WebUserMenu" tagprefix="uc1" %>
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
    <table style="position:relative; margin-left:auto; margin-right:auto; width:990px; top: 0px; left: 0px;"  
        border ="0">
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
        <table cellpadding=5 cellspacing=0 width="100%" border="1" id="customers">
            <tr>
                <td class="fail" colspan="20" style="text-align: center">
                    <asp:Label ID="Label1" runat="server" Font-Bold="True" Style="text-align: center"></asp:Label>
                    <asp:Label ID="l_PgmType" runat="server" Font-Bold="True"></asp:Label>
                </td>
            </tr>
            <tr>
                <td colspan="4" class="style1">
                    資料編號<asp:Image ID="Image2" runat="server" /></td>
                <td colspan="16" class="style2"><asp:TextBox ID="txt_pc_code" runat="server" MaxLength="12" Width="120px" Enabled="False"></asp:TextBox>
                    <font color="red">系統自動取號</font></td>
            </tr>
            <tr align="center"><td colspan="20" class="style1">自主管理輔導稽查表回傳</td></tr>
              <tr><td colspan="20" class="style1" align="center">事業單位資料</td></tr>
              <tr><td colspan="4" align="left" class="style1"><font color=red>*</font>通報單位:</td>
                    <td colspan="16" class="style2"><asp:TextBox ID="txt_com_name" runat="server" maxlength="50" Width="300px" TabIndex="1" /></td>
                </tr>
				<tr>
                <td colspan="4" class="style1"><asp:Label ID="l_att_name1" runat="server" Text="<font color=red>*</font>聯絡人：" /></td>
                <td colspan="6" class="style2"><asp:TextBox ID="txt_att_name1" runat="server" maxlength="50" Width="180px" TabIndex="2" /></td>
                <td colspan="4" class="style1"><asp:Label ID="l_att_mail" runat="server" Text="<font color=red>*</font>電子信箱：" /></td>
                <td colspan="6" class="style2"><asp:TextBox ID="txt_att_mail" runat="server" maxlength="50" Width="180px" TabIndex="3" /></td>
              </tr>
			  <tr>
                <td colspan="4" class="style1"><asp:Label ID="l_att_tel1" runat="server" Text="<font color=red>*</font>聯絡電話：" /></td>
                <td colspan="6" class="style2"><asp:TextBox ID="txt_att_tel1" runat="server" maxlength="16" Width="150px" TabIndex="4" />
                    <font color="red">Ex:0212345678</font></td>
                <td colspan="4" class="style1"><asp:Label ID="l_att_cell_phone" runat="server" Text="<font color=red>*</font>手機：" /></td>
                <td colspan="6" class="style2"><asp:TextBox ID="txt_att_cell_phone" runat="server" maxlength="16" Width="150px" TabIndex="5" />
                    <font color="red">Ex:09XX-XXXXXX</font></td>
              </tr>
			  
			  <tr>
              <td colspan="4" class="style1"><asp:Label ID="l_pc_date" runat="server" Text="<font color=red>*</font>通報日期：" /></td>
              <td colspan="16" class="style2"><asp:Label ID="txt_pc_date" runat="server" maxlength="50" Width="120px"/></td>
              </tr>
			  <tr><td colspan="20" class="style1" align="center">承攬資料</td></tr>
			  <tr>
                <td colspan="4" class="style1"><asp:Label ID="l_eng_name" runat="server" Text="<font color=red>*</font>工程名稱：" /></td>
                <td colspan="6" class="style2"><asp:TextBox ID="txt_eng_name" runat="server" maxlength="50" Width="180px" TabIndex="9" /></td>
                <td colspan="4" class="style1"><asp:Label ID="l_eng_code" runat="server" Text="建號：" /></td>
                <td colspan="6" class="style2"><asp:TextBox ID="txt_eng_code" runat="server" maxlength="20" Width="180px" TabIndex="10" /></td>
              </tr>
              <tr>
                <td colspan="4" class="style1"><asp:Label ID="l_eng_add" runat="server" Text="<font color=red>*</font>工程地址：" /></td>
					<td colspan="16" class="style2">臺北市
                    <asp:DropDownList ID="ddl_eng_area" runat="server" TabIndex="11" AutoPostBack="True" />
                    <asp:DropDownList ID="ddl_eng_road" runat="server" TabIndex="12" /><br />
                    <asp:TextBox ID="txt_eng_add" runat="server" maxlength="100" Width="400px" TabIndex="13" />

                    </td>
                </tr>
				<tr>
                <td colspan="4" class="style1"><asp:Label ID="l_eng_floor" runat="server" Text="作業區域及樓層：" /></td>
					<td colspan="16" class="style2"><asp:TextBox ID="txt_eng_floor" runat="server" maxlength="50" Width="200px" TabIndex="14" />
                    </td>
                </tr>
                <tr>
                <td colspan="4" class="style1"><asp:Label ID="l_eng_date" runat="server" Text="<font color=red>*</font>預計施工期間：" /></td>
                <td colspan="16" class="style2">起:西元 <asp:DropDownList ID="ddl_year_s" runat="server" TabIndex="15" AutoPostBack="True"/> 年 
                                <asp:DropDownList ID="ddl_month_s" runat="server" TabIndex="16" AutoPostBack="True"/> 月 
                                <asp:DropDownList ID="ddl_date_s" runat="server" TabIndex="17" /> 日 
                                <br />迄:西元 <asp:DropDownList ID="ddl_year_e" runat="server" TabIndex="18" AutoPostBack="True"/> 年 
                                <asp:DropDownList ID="ddl_month_e" runat="server" TabIndex="19" AutoPostBack="True"/> 月 
                                <asp:DropDownList ID="ddl_date_e" runat="server" TabIndex="20" /> 日 </td>
              </tr>
			  
			  <tr>
                <td colspan="4" class="style1"><asp:Label ID="l_pc_kind" runat="server" Text="<font color=red>*</font>種類：" /></td>
					<td colspan="8" class="style2" VALIGN="top">
					<asp:Label ID="l_pc_kind1" runat="server" Text="丁類危險性工作場所主要危害作業：" /><br>
					<asp:CheckBoxList ID="chk_pc_kind1" runat="server" TabIndex="21" 
                        RepeatColumns="1" RepeatDirection="Horizontal" />
                    </td>
					<td colspan="8" class="style2" VALIGN="top">
					<asp:Label ID="l_pc_kind2" runat="server" Text="機械作業：" /><br>
					<asp:CheckBoxList ID="chk_pc_kind2" runat="server" TabIndex="21" 
                        RepeatColumns="1" RepeatDirection="Horizontal" />
                    </td>
                </tr>
            </tr>
            </table>
    </div>
            <asp:Button ID="Button1" runat="server" Text="存檔" />
            <asp:Button ID="Bt_BackUp" runat="server" Text="離開" />
            <asp:Button ID="Bt_send" runat="server" Text="補發" Visible="False" />
            <asp:TextBox ID="txtPgmType" runat="server" Visible="False" />
            <asp:TextBox ID="txtPrimary" runat="server" Visible="False" />
            <asp:TextBox ID="txtSubWhere" runat="server" Visible="False" />
            <asp:TextBox ID="txtOrder" runat="server" Visible="False" />
            <br />
            <asp:Label ID="l_ErrMsg" runat="server" Font-Bold="True" ForeColor="Red" />
            </td>
        </tr>
      </table>    
    </form>
</body>
</html>
