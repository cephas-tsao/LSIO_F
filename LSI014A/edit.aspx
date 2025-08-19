<%@ Page Language="VB" AutoEventWireup="false" CodeFile="edit.aspx.vb" Inherits="basic_LSI014A_edit" aspcompat="true" MaintainScrollPositionOnPostback="true" %>
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
        <table width="100%" border="1" id="customers" cellspacing='0.1'>
            <tr>
                <td class="fail" colspan="2" style="text-align: center">
                    <asp:Label ID="Label1" runat="server" Font-Bold="True" Style="text-align: center"></asp:Label>
                    <asp:Label ID="l_PgmType" runat="server" Font-Bold="True"></asp:Label>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    資料編號<asp:Image ID="Image2" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_dan_code" runat="server" MaxLength="12" Width="120px" 
                        Enabled="False"></asp:TextBox>
                    <font color="red">系統自動取號</font></td>
            </tr>
            <tr>
                <td class="style1">
                    事業單位名稱<asp:Image ID="Image1" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_com_name" runat="server" MaxLength="50" Width="500px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    通報人<asp:Image ID="Image3" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_pet_name" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    通報日期</td>
                <td class="style2">
                    <asp:TextBox ID="txt_pet_date" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                    <asp:ImageButton ID="img_date1" runat="server" 
                        ImageUrl="~/images/calendr.gif" />
                     <font color="red">Ex:2011/01/01<asp:TextBox ID="txt_pet_time" runat="server" 
                        MaxLength="5" Width="50px"></asp:TextBox></font></td>
            </tr>
            <tr>
                <td class="style1">
                    通報人聯絡電話<asp:Image ID="Image4" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_pet_tel1" runat="server" MaxLength="16" Width="160px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    廠商名稱<asp:Image ID="Image5" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_att_name1" runat="server" MaxLength="50" Width="300px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    廠商電話<asp:Image ID="Image6" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_att_tel1" runat="server" MaxLength="16" Width="160px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    施工人員姓名<asp:Image ID="Image7" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_att_name2" runat="server" MaxLength="50" Width="300px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    施工人員電話<asp:Image ID="Image8" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_att_tel2" runat="server" MaxLength="16" Width="160px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    施工地址<asp:Image ID="Image9" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:DropDownList ID="ddl_att_area" runat="server" AutoPostBack="True">
                    </asp:DropDownList>
                    <asp:DropDownList ID="ddl_att_road" runat="server">
                    </asp:DropDownList>
                    <asp:TextBox ID="txt_att_add1" runat="server" MaxLength="100" Width="500px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    大廈名稱</td>
                <td class="style2">
                    <asp:TextBox ID="txt_bulid_name" runat="server" MaxLength="20" Width="200px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    建號名稱</td>
                <td class="style2">
                    <asp:TextBox ID="txt_bulid_code" runat="server" MaxLength="20" Width="200px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    施工期間<asp:Image ID="Image10" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_work_date_s" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                    <asp:ImageButton ID="img_date2" runat="server" 
                        ImageUrl="~/images/calendr.gif" />
                     <font color="red">~<asp:TextBox ID="txt_work_date_e" runat="server" 
                        MaxLength="10" Width="100px"></asp:TextBox><asp:ImageButton ID="img_date3" runat="server" 
                        ImageUrl="~/images/calendr.gif" />
                     Ex:2011/01/01</font></td>
            </tr>
            <tr>
                <td class="style1">
                    是否夜間施工</td>
                <td class="style2">
                     <font color="red">
                    <asp:DropDownList ID="ddl_is_night" runat="server">
                    </asp:DropDownList>
                    </font>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    週六、週日是否施工</td>
                <td class="style2">
                     <font color="red">
                    <asp:DropDownList ID="ddl_is_holiday" runat="server">
                    </asp:DropDownList>
                    </font>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    作業樓層<asp:Image ID="Image11" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_work_floor_s" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                                ~<asp:TextBox ID="txt_work_floor_e" runat="server" MaxLength="10" 
                        Width="100px"></asp:TextBox><font color=red>EX:B6~40F</font></td>
            </tr>
            <tr>
                <td class="style1">
                    作業類別<asp:Image ID="Image12" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                                <asp:DropDownList ID="ddl_work_type1" runat="server">
                    </asp:DropDownList>
                    其他：<asp:TextBox ID="txt_work_type_memo" runat="server" MaxLength="30" 
                        Width="300px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    使用機具</td>
                <td class="style2">
                    <asp:CheckBoxList ID="ckl_use_tool1" runat="server" RepeatColumns="2" RepeatDirection="Horizontal"
                        Height="16px">
                    </asp:CheckBoxList>
                    其他：<asp:TextBox ID="txt_use_tool5_memo" runat="server" MaxLength="30" 
                        Width="300px"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="style1">
                    安全裝置</td>
                <td class="style2">
                    <asp:CheckBoxList ID="ckl_save_tool" runat="server" RepeatColumns="2" RepeatDirection="Horizontal"
                        Height="16px">
                    </asp:CheckBoxList>
                    其他：<asp:TextBox ID="txt_save_tool_memo" runat="server" MaxLength="30" 
                        Width="300px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    備註</td>
                <td class="style2">
                    <asp:TextBox ID="txt_dan_memo_memo" runat="server" MaxLength="30" Width="300px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    通報回覆1<asp:Image ID="Image13" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_dan_mail" runat="server" MaxLength="30" Width="300px"></asp:TextBox>
                                <font color="red">Ex:AA@yahoo.com.tw</font>
                                <asp:Button ID="Bt_send" runat="server" Text="補發" Visible="False" /></td>
            </tr>
            <tr>
                <td class="style1">
                    通報回覆2</td>
                <td class="style2">
                    <asp:TextBox ID="txt_dan_mail2" runat="server" MaxLength="30" Width="300px"></asp:TextBox>
                                <font color="red">Ex:AA@yahoo.com.tw</font></td>
            </tr>
            <tr>
                <td class="style1">
                    裝修通報單號</td>
                <td class="style2">
                    <asp:TextBox ID="txt_fix_code" runat="server" MaxLength="12" Width="120px"></asp:TextBox></td>
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
