<%@ Page Language="VB" AutoEventWireup="false" CodeFile="edit.aspx.vb" Inherits="basic_LSI013A_edit" aspcompat="true" MaintainScrollPositionOnPostback="true" %>
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
            width: 18%;
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
                    <asp:TextBox ID="txt_fix_code" runat="server" MaxLength="12" Width="120px" 
                        Enabled="False"></asp:TextBox>
                    <font color="red">系統自動取號</font></td>
            </tr>
            <tr>
                <td class="style1">
                    通報單位<asp:Image ID="Image1" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_com_name" runat="server" MaxLength="50" Width="500px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    聯絡電話<asp:Image ID="Image3" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_pet_tel1" runat="server" MaxLength="16" Width="160px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    通報對象<asp:Image ID="Image10" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:RadioButtonList ID="l_pet_man" runat="server" 
                      RepeatDirection="Horizontal">
                      <asp:ListItem Value="0" TabIndex="3">業主</asp:ListItem>
                      <asp:ListItem Value="1" TabIndex="4">主辦機關</asp:ListItem>
                      <asp:ListItem Value="2" TabIndex="5">施工廠商</asp:ListItem>
                  </asp:RadioButtonList>
                  </td>
                                
            </tr>
            <tr>
                <td class="style1">
                    通報人員<asp:Image ID="Image11" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_pet_name" runat="server" MaxLength="50" Width="250px"></asp:TextBox>
                                </td>
            </tr>
                <tr>
                <td class="style1">
                    通報回覆<asp:Image ID="Image13" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_pet_email" runat="server" MaxLength="50" Width="250px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    行動電話<asp:Image ID="Image12" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_pet_tel2" runat="server" MaxLength="16" Width="160px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    施工廠商(大包)<asp:Image ID="Image14" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_pet_name1" runat="server" MaxLength="20" Width="200px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    現場負責人(大包)<asp:Image ID="Image15" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_pet_name2" runat="server" MaxLength="100" Width="500px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    聯絡電話(大包)<asp:Image ID="Image16" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_pet_tel3" runat="server" MaxLength="16" Width="160px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    行動電話(大包)<asp:Image ID="Image17" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_pet_tel4" runat="server" MaxLength="16" Width="160px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    施工廠商(小包)<asp:Image ID="Image18" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_pet_name3" runat="server" MaxLength="20" Width="200px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    現場負責人(小包)<asp:Image ID="Image19" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_pet_name4" runat="server" MaxLength="100" Width="500px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    聯絡電話(小包)<asp:Image ID="Image20" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_pet_tel5" runat="server" MaxLength="16" Width="160px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    行動電話(小包)<asp:Image ID="Image21" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_pet_tel6" runat="server" MaxLength="16" Width="160px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    工程地址<asp:Image ID="Image5" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:DropDownList ID="ddl_att_area" runat="server"  
                        AutoPostBack="True">
                    </asp:DropDownList>
                    <asp:DropDownList ID="ddl_att_road" runat="server" >
                    </asp:DropDownList>
                    <asp:TextBox ID="txt_att_add1" runat="server" MaxLength="100" Width="500px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    工程名稱<asp:Image ID="Image9" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_safe_item" runat="server" MaxLength="30" Width="200px"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="style1">
                    工程種類<asp:Image ID="Image22" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:RadioButtonList ID="radchoose" runat="server" RepeatColumns="3" 
                        RepeatDirection="Horizontal">
                        <asp:ListItem Value="1" TabIndex="22">建築新工程</asp:ListItem>
                        <asp:ListItem Value="2" TabIndex="23">建築修繕工程</asp:ListItem>
                        <asp:ListItem Value="3" TabIndex="24">其他建築工程</asp:ListItem>
                        <asp:ListItem Value="4" TabIndex="25">土木新建工程</asp:ListItem>
                        <asp:ListItem Value="5" TabIndex="26">土木修繕工程</asp:ListItem>
                        <asp:ListItem Value="6" TabIndex="27">其他土木工程</asp:ListItem>
                    </asp:RadioButtonList>
                    </td>
            </tr>
             <tr>
                <td class="style1">
                    是否為丁類危險場所<asp:Image ID="Image23" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:RadioButtonList ID="rad_att_area" runat="server" 
                      RepeatDirection="Horizontal" TabIndex="28">
                  </asp:RadioButtonList></td>
            </tr>
             <tr>
                <td class="style1">
                    通報種類<asp:Image ID="Image8" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:CheckBoxList ID="ckl_fix_type" runat="server" RepeatColumns="2" 
                        Height="16px">
                    </asp:CheckBoxList>
                    <asp:CheckBox ID="chkEng" runat="server" TabIndex="30"/>
                    建築工程之外牆施工架拆組(總高度：<asp:TextBox ID="totheigh" runat="server" maxlength="30" Width="100px" TabIndex="31" />公尺)<br>
                    <asp:CheckBox ID="chktoher" runat="server" TabIndex="32"/>
                    其他：請加註<asp:TextBox ID="txt_work_type_memo" runat="server" maxlength="30" Width="350px" TabIndex="33" /></td>
                                </td>
            </tr>
            
            <tr>
                <td class="style1">
                    施工期間<asp:Image ID="Image7" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_work_date_s" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                    <asp:ImageButton ID="img_date2" runat="server" 
                        ImageUrl="~/images/calendr.gif" />
                     <font color="red">~<asp:TextBox ID="txt_work_date_e" runat="server" 
                        MaxLength="10" Width="100px"></asp:TextBox><asp:ImageButton ID="img_date3" runat="server" 
                        ImageUrl="~/images/calendr.gif" />
                     EX:2011/01/01</font></td>
            </tr>
            <tr>
                <td class="style1">
                    是否夜間施工<asp:Image ID="Image24" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:RadioButtonList ID="rad_att_night" runat="server" 
                      RepeatDirection="Horizontal" TabIndex="28">
                  </asp:RadioButtonList></td>
            </tr>
            <tr>
                <td class="style1">
                    週六日是否施工<asp:Image ID="Image25" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:RadioButtonList ID="rad_att_holiday" runat="server" 
                      RepeatDirection="Horizontal" TabIndex="28">
                  </asp:RadioButtonList></td>
            </tr>
             <tr>
                <td class="style1">
                    作業區域<asp:Image ID="Image6" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_work_floor" runat="server" MaxLength="20" Width="200px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    業主名稱</td>
                <td class="style2">
                    <asp:TextBox ID="txt_att_cname1" runat="server" MaxLength="50" Width="300px"></asp:TextBox>
                    </td>
            </tr>
            <tr>
                <td class="style1">
                    聯絡電話</td>
                <td class="style2">
                    <asp:TextBox ID="txt_att_tel1" runat="server" MaxLength="16" Width="160px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    聯絡人</td>
                <td class="style2">
                    <asp:TextBox ID="txt_att_name1" runat="server" MaxLength="50" Width="300px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    行動電話</td>
                <td class="style2">
                    <asp:TextBox ID="txt_att_tel2" runat="server" MaxLength="50" Width="300px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    是否為公共工程<asp:Image ID="Image26" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:RadioButtonList ID="rad_att_public" runat="server" 
                      RepeatDirection="Horizontal" TabIndex="28">
                  </asp:RadioButtonList></td>
            </tr>
            <tr>
                <td class="style1">
                    通報日期<asp:Image ID="Image4" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_pet_date" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                    <asp:ImageButton ID="img_date1" runat="server" 
                        ImageUrl="~/images/calendr.gif" />
                     <font color="red">EX:2011/01/01　時間：<asp:TextBox ID="txt_pet_time" runat="server" 
                        MaxLength="5" Width="50px"></asp:TextBox></font></td>
            </tr>
            <tr>
                <td class="style1">
                    備註</td>
                <td class="style2">
                    <asp:TextBox ID="txt_att_comment" runat="server" MaxLength="50" Width="300px"></asp:TextBox>
                </td>
            </tr>
            
           
            <!--tr>
                <td class="style1">
                    作業處所</td>
                <td class="style2">
                    <asp:CheckBoxList ID="ckl_work_floor_type" runat="server" Height="16px" 
                        RepeatDirection="Horizontal">
                    </asp:CheckBoxList>
                                </td>
            </tr-->
            
            
            
            
           
            
            <tr>
                <td class="style1">
                    危險性作業單號</td>
                <td class="style2">
                    <asp:TextBox ID="txt_dan_code" runat="server" MaxLength="12" Width="120px"></asp:TextBox></td>
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
