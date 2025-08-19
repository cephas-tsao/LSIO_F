<%@ Page Language="VB" AutoEventWireup="false" CodeFile="edit.aspx.vb" Inherits="basic_LSI015A_edit" aspcompat="true" MaintainScrollPositionOnPostback="true" %>
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
                <td class="fail" colspan="2" style="text-align: center">
                    <asp:Label ID="Label1" runat="server" Font-Bold="True" Style="text-align: center"></asp:Label></td>
            </tr>
            <tr>
                <td class="style1">
                    資料編號<asp:Image ID="Image2" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_danc_code" runat="server" MaxLength="12" Width="120px" 
                        Enabled="False"></asp:TextBox>
                    <font color="red">系統自動取號</font></td>
            </tr>
            <tr>
                <td class="style1">
                    工程名稱<asp:Image ID="Image1" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_eng_name" runat="server" MaxLength="70" Width="500px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    通報日期<asp:Image ID="Image3" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_pet_date" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                    <asp:ImageButton ID="img_date1" runat="server" 
                        ImageUrl="~/images/calendr.gif" />
                     <font color="red">EX:2011/01/01<asp:TextBox ID="txt_pet_time" runat="server" 
                        MaxLength="5" Width="50px"></asp:TextBox></font></td>
            </tr>
            <tr>
                <td class="style1">
                    施工地址<asp:Image ID="Image4" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_eng_add" runat="server" MaxLength="100" Width="500px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    丁類危險性工作項目<asp:Image ID="Image5" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:CheckBoxList ID="ckl_danc_kind" runat="server" Height="16px">
                    </asp:CheckBoxList>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    送審方式<asp:Image ID="Image6" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:RadioButtonList ID="rbl_danc_step" runat="server">
                    </asp:RadioButtonList>
                    <asp:TextBox ID="txt_danc_div" runat="server" MaxLength="5" Width="120px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    送審類別<asp:Image ID="Image7" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:RadioButtonList ID="rbl_danc_type" runat="server">
                    </asp:RadioButtonList>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    事業單位名稱<asp:Image ID="Image8" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_com_name" runat="server" MaxLength="100" Width="500px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    事業單位統編<asp:Image ID="Image9" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_com_idno" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    公司地址<asp:Image ID="Image10" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:DropDownList ID="ddl_att_area" runat="server" AutoPostBack="True">
                    </asp:DropDownList>
                    <asp:DropDownList ID="ddl_att_road" runat="server">
                    </asp:DropDownList>
                    <asp:TextBox ID="txt_com_add" runat="server" MaxLength="100" Width="500px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    公司電話<asp:Image ID="Image11" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_com_tel1" runat="server" MaxLength="16" Width="160px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    負責人-事業經營<asp:Image ID="Image12" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_att_name1" runat="server" MaxLength="20" Width="200px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    負責人-工作場所<asp:Image ID="Image13" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_att_name2" runat="server" MaxLength="20" Width="200px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    聯絡人姓名<asp:Image ID="Image14" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_con_name" runat="server" MaxLength="20" Width="200px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    聯絡人電話一<asp:Image ID="Image15" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_con_tel1" runat="server" MaxLength="16" Width="160px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    聯絡人電話二<asp:Image ID="Image16" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_con_tel2" runat="server" MaxLength="16" Width="160px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    聯絡人mail<asp:Image ID="Image17" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_con_email" runat="server" MaxLength="60" Width="500px"></asp:TextBox>
                    <asp:Button ID="Bt_send" runat="server" Text="補發" Visible="False" /></td>
            </tr>
            <tr>
                <td class="style1">
                    查詢進度密碼</td>
                <td class="style2">
                    <asp:TextBox ID="txt_con_pw" runat="server" MaxLength="12" Width="120px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    案件狀態</td>
                <td class="style2">
                    <asp:RadioButtonList ID="rbl_danc_state" runat="server">
                    </asp:RadioButtonList>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    審查結果</td>
                <td class="style2">
                    <asp:RadioButtonList ID="rbl_result" runat="server">
                    </asp:RadioButtonList>
                                </td>
            </tr>
            </table>
    </div>
            <asp:Button ID="Button1" runat="server" Text="存檔" />
            <asp:Button ID="Bt_BackUp" runat="server" Text="離開" />
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
