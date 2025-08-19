<%@ Page Language="VB" AutoEventWireup="false" CodeFile="edit.aspx.vb" Inherits="basic_LSI003A_edit" aspcompat="true" MaintainScrollPositionOnPostback="true" validateRequest="False"%>
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
                    <asp:TextBox ID="txt_Down_code" runat="server" MaxLength="5" Width="300px" 
                        Enabled="False"></asp:TextBox>
                    <font color="red">系統自動取號</font>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    事業單位名稱<asp:Image ID="Image1" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_Down_institutions" runat="server" MaxLength="100" Width="300px"></asp:TextBox>
                    <font color="red">100字元以內</font>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    工程名稱<asp:Image ID="Image8" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_Down_title" runat="server" MaxLength="100" Width="300px"></asp:TextBox>
                    <font color="red">100字元以內</font>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    停工地點<asp:Image ID="Image3" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_Down_address" runat="server" MaxLength="100" Width="300px"></asp:TextBox>
                    <font color="red">100字元以內</font>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    緯度<asp:Image ID="Image4" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_Down_X" runat="server" MaxLength="100" Width="300px"></asp:TextBox>
                    <font color="red">100字元以內</font>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    經度<asp:Image ID="Image5" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_Down_Y" runat="server" MaxLength="100" Width="300px"></asp:TextBox>
                    <font color="red">100字元以內</font>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    停工日期<asp:Image ID="Image7" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_Down_stopDate" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                    <asp:ImageButton ID="img_date2" runat="server" ImageUrl="~/images/calendr.gif" />
                    <font color="red">格式為yyyy/MM/dd</font>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    復工日期<asp:Image ID="Image9" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_Down_returnDate" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                    <asp:ImageButton ID="img_date3" runat="server" ImageUrl="~/images/calendr.gif" />
                    <font color="red">格式為yyyy/MM/dd</font>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    停工範圍<asp:Image ID="Image6" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_Down_range" runat="server" MaxLength="300" Width="300px" Height="100" TextMode="MultiLine"></asp:TextBox><font color="red">300字元以內</font>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    停工原因<asp:Image ID="Image10" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_Down_reason" runat="server" MaxLength="300" Width="300px" Height="100" TextMode="MultiLine"></asp:TextBox><font color="red">300字元以內</font>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    備註</td>
                <td class="style2">
                    <asp:TextBox ID="txt_Down_remarks" runat="server" MaxLength="300" Width="300px" Height="100" TextMode="MultiLine"></asp:TextBox><font color="red">300字元以內</font>
                </td>
            </tr>
<%--            <tr>
                <td class="style1">
                    置頂開始日期<asp:Image ID="Image12" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_Down_sDate" runat="server" MaxLength="10" Width="300px"></asp:TextBox>
                    <asp:ImageButton ID="ImageButton5" runat="server" ImageUrl="~/images/calendr.gif" />
                </td>
            </tr>
            <tr>
                <td class="style1">
                    置頂結束日期<asp:Image ID="Image13" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_Down_eDate" runat="server" MaxLength="10" Width="300px"></asp:TextBox>
                    <asp:ImageButton ID="ImageButton4" runat="server" ImageUrl="~/images/calendr.gif" />
                </td>
            </tr>
            <tr>
                <td class="style1">
                    發布日期<asp:Image ID="Image14" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_Down_pDate" runat="server" MaxLength="10" Width="300px"></asp:TextBox>
                    <asp:ImageButton ID="ImageButton3" runat="server" ImageUrl="~/images/calendr.gif" />
                </td>
            </tr>
--%>            <tr>
                <td class="style1">
                    建立日期</td>
                <td class="style2">
                    <asp:TextBox ID="txt_Down_cDate" runat="server" MaxLength="10" Width="300px" ReadOnly="true"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    修改日期</td>
                <td class="style2">
                    <asp:TextBox ID="txt_Down_uDate" runat="server" MaxLength="10" Width="300px" ReadOnly="true"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    建立者編號</td>
                <td class="style2">
                    <asp:TextBox ID="txt_Down_creatorID" runat="server" MaxLength="5" Width="300px" ReadOnly="true"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    修改者編號</td>
                <td class="style2">
                    <asp:TextBox ID="txt_Down_modifyID" runat="server" MaxLength="5" Width="300px" ReadOnly="true"></asp:TextBox>
                </td>
            </tr>
<%--            <tr>
                <td class="style1">
                    建立者IP<asp:Image ID="Image19" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_Down_creatorIP" runat="server" MaxLength="50" Width="300px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    修改者IP<asp:Image ID="Image20" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_Down_modifyIP" runat="server" MaxLength="50" Width="300px"></asp:TextBox>
                </td>
            </tr>--%>
            </table>
    </div>
            <asp:Button ID="Button1" runat="server" Text="存檔" />
            <asp:Button ID="Bt_BackUp" runat="server" Text="離開" />
            <asp:TextBox ID="txtPgmType" runat="server" Visible="False" />
            <asp:TextBox ID="txtPrimary" runat="server" Visible="False" />
            <asp:TextBox ID="txtSubWhere" runat="server" Visible="False" />
            <asp:TextBox ID="txtOrder" runat="server" Visible="False" />
            <asp:TextBox ID="txt_book_premiums" runat="server" Visible="False" /> <!--RadioButton贈品-->
            <br />
            <asp:Label ID="l_ErrMsg" runat="server" Font-Bold="True" ForeColor="Red" />
            </td>
        </tr>
      </table>    
    </form>
</body>
</html>
