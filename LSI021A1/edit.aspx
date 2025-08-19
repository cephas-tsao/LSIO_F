<%@ Page Language="VB" AutoEventWireup="false" CodeFile="edit.aspx.vb" Inherits="basic_LSI021A_edit" AspCompat="true" MaintainScrollPositionOnPostback="true" %>

<%@ Register Src="../../usercontrol/WebUserMenu.ascx" TagName="WebUserMenu" TagPrefix="uc1" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title><%=ConfigurationManager.AppSettings("Prj_name")%></title>
    <script language="javascript" id="js1" src="../../public/checkjs.js"></script>
    <link rel="stylesheet" href="https://pro.fontawesome.com/releases/v5.10.0/css/all.css" integrity="sha384-AYmEC3Yw5cVb3ZcuHtOA93w35dYTsvhLPVnYs9eStHfGJvOvKxVfELGroGkvsg+p" crossorigin="anonymous" />
    <link rel='stylesheet' href="/../Setting/style<%=Get_SysPara("css_sys", "01")%>.css" type='text/css' />
    <style type="text/css">
        .style1 {
            width: 13%;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <table style="position: relative; margin-left: auto; margin-right: auto; width: 990px; top: 0px; left: 0px;"
            border="0">
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
                                <td class="style1">資料來源</td>
                                <td class="style2">
                                    <asp:RadioButtonList ID="selwork" runat="server">
                                        <asp:ListItem Value="01">網路通報</asp:ListItem>
                                        <asp:ListItem Value="02">電子郵件</asp:ListItem>
                                        <asp:ListItem Value="03">傳真通報</asp:ListItem>
                                        <asp:ListItem Value="04">親自通報</asp:ListItem>
                                    </asp:RadioButtonList>
                            </tr>
                            <tr>
                                <td class="style1">資料編號<asp:Image ID="Image2" runat="server" ImageUrl="~/images/tick.png" /></td>
                                <td class="style2">
                                    <asp:TextBox ID="txt_roof_cod" runat="server" MaxLength="12" Width="120px"
                                        Enabled="False"></asp:TextBox>
                                    <font color="red">系統自動取號</font></td>
                            </tr>
                            <tr>
                                <td class="style1">工程名稱<asp:Image ID="Image1" runat="server" ImageUrl="~/images/tick.png" /></td>
                                <td class="style2">
                                    <asp:TextBox ID="txt_com_name" runat="server" MaxLength="50" Width="500px"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="style1">工程地址<asp:Image ID="Image9" runat="server" ImageUrl="~/images/tick.png" /></td>
                                <td class="style2">
                                    <asp:DropDownList ID="ddl_att_area" runat="server" AutoPostBack="True">
                                    </asp:DropDownList>
                                    <asp:DropDownList ID="ddl_att_road" runat="server">
                                    </asp:DropDownList>
                                    <asp:TextBox ID="txt_att_add1" runat="server" MaxLength="100" Width="500px"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="style1">廠商名稱<asp:Image ID="Image5" runat="server" ImageUrl="~/images/tick.png" /></td>
                                <td class="style2">
                                    <asp:TextBox ID="txt_att_name1" runat="server" MaxLength="50" Width="300px"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="style1">廠商聯絡電話<asp:Image ID="Image6" runat="server" ImageUrl="~/images/tick.png" /></td>
                                <td class="style2">
                                    <asp:TextBox ID="txt_att_tel1" runat="server" MaxLength="16" Width="160px"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="style1">現場負責人<asp:Image ID="Image7" runat="server" ImageUrl="~/images/tick.png" /></td>
                                <td class="style2">
                                    <asp:TextBox ID="txt_att_name2" runat="server" MaxLength="50" Width="300px"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="style1">行動電話<asp:Image ID="Image8" runat="server" ImageUrl="~/images/tick.png" /></td>
                                <td class="style2">
                                    <asp:TextBox ID="txt_att_tel2" runat="server" MaxLength="16" Width="160px"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="style1">聯絡人Email<asp:Image ID="Image3" runat="server" ImageUrl="~/images/tick.png" /></td>
                                <td class="style2">
                                    <asp:TextBox ID="txt_roof_mail" runat="server" MaxLength="50" Width="300px"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="style1">通報日期<asp:Image ID="Image4" runat="server" ImageUrl="~/images/tick.png" /></td>
                                <td class="style2">
                                    <asp:TextBox ID="txt_pet_date" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                                    <asp:ImageButton ID="img_date1" runat="server"
                                        ImageUrl="~/images/calendr.gif" />
                                    <font color="red">Ex:2011/01/01</font></td>
                            </tr>
                            <tr>
                                <td class="style1" rowspan="4">輕質屋頂作業</td>
                            </tr>
                            <tr>
                                <td class="style2">預定施工期間<asp:TextBox ID="txt_roof_sdate" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                                    <asp:ImageButton ID="roof_sdate1" runat="server"
                                        ImageUrl="~/images/calendr.gif" />至
                    <asp:TextBox ID="txt_roof_edate" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                                    <asp:ImageButton ID="roof_edate1" runat="server"
                                        ImageUrl="~/images/calendr.gif" />
                                </td>
                            </tr>
                            <tr>
                                <td class="style2">屋頂高度：<asp:TextBox ID="txt_roof_height" runat="server" MaxLength="3"
                                    Width="30px"></asp:TextBox>
                                    層樓或樓高：<asp:TextBox ID="txt_floor_height" runat="server" MaxLength="3"
                                        Width="30px"></asp:TextBox>公尺
                                </td>
                            </tr>
                            <tr>
                                <td class="style2">
                                    <asp:CheckBoxList ID="chk_roof_tool1" runat="server" RepeatColumns="2" RepeatDirection="Horizontal"
                                        Height="16px">
                                    </asp:CheckBoxList>
                                    其他：<asp:TextBox ID="txt_roof_memo" runat="server" MaxLength="30"
                                        Width="300px"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="style1" rowspan="5">施工架組配與拆除</td>
                            </tr>
                            <tr>
                                <td class="style2">預定組配時間：<asp:TextBox ID="txt_tearon_sdate" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                                    <asp:ImageButton ID="tearon_sdate1" runat="server"
                                        ImageUrl="~/images/calendr.gif" />至
                    <asp:TextBox ID="txt_tearon_edate" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                                    <asp:ImageButton ID="tearon_edate1" runat="server"
                                        ImageUrl="~/images/calendr.gif" />
                                </td>
                            </tr>
                            <tr>
                                <td class="style2">預定拆除時間：<asp:TextBox ID="txt_tearoff_sdate" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                                    <asp:ImageButton ID="tearoff_sdate1" runat="server"
                                        ImageUrl="~/images/calendr.gif" />至
                    <asp:TextBox ID="txt_tearoff_edate" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                                    <asp:ImageButton ID="tearoff_edate1" runat="server"
                                        ImageUrl="~/images/calendr.gif" />
                                </td>
                            </tr>
                            <tr>
                                <td class="style2">施工架組配或拆除高度：<asp:TextBox ID="txt_tear_height" runat="server" MaxLength="3"
                                    Width="30px"></asp:TextBox>
                                    層樓或高度：<asp:TextBox ID="txt_tearfloor_height" runat="server" MaxLength="3"
                                        Width="30px"></asp:TextBox>層樓
                                </td>
                            </tr>
                            <tr>
                                <td class="style2">
                                    <asp:CheckBoxList ID="chk_tear_tool1" runat="server" RepeatColumns="2" RepeatDirection="Horizontal"
                                        Height="16px">
                                    </asp:CheckBoxList>
                                    其他：<asp:TextBox ID="txt_tear_memo" runat="server" MaxLength="30"
                                        Width="300px"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="style1" rowspan="4">使用吊籠作業</td>
                            </tr>
                            <tr>
                                <td class="style2">預定施工期間<asp:TextBox ID="txt_cage_sdate" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                                    <asp:ImageButton ID="cage_sdate1" runat="server"
                                        ImageUrl="~/images/calendr.gif" />至
                    <asp:TextBox ID="txt_cage_edate" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                                    <asp:ImageButton ID="cage_edate1" runat="server"
                                        ImageUrl="~/images/calendr.gif" />
                                </td>
                            </tr>
                            <tr>
                                <td class="style2">建築物高度：<asp:TextBox ID="txt_cageroof_height" runat="server" MaxLength="3"
                                    Width="30px"></asp:TextBox>公尺
                                    層樓：<asp:TextBox ID="txt_cagefloor_height" runat="server" MaxLength="3"
                                        Width="30px"></asp:TextBox>層樓
                                </td>
                            </tr>
                            <tr>
                                <td class="style2">作業類別:
 <asp:CheckBoxList ID="chk_rtype2" runat="server" RepeatColumns="2" RepeatDirection="Horizontal"
                                        Height="16px">
                                    </asp:CheckBoxList>作業項目:
                                    <asp:CheckBoxList ID="chk_cage_tool1" runat="server" RepeatColumns="2" RepeatDirection="Horizontal"
                                        Height="16px">
                                    </asp:CheckBoxList>
                                    其他：<asp:TextBox ID="txt_cage_memo" runat="server" MaxLength="30"
                                        Width="300px"></asp:TextBox>
                                </td>

                                <tr>
                                    <td class="style1">回覆日期<asp:Image ID="Image10" runat="server" ImageUrl="~/images/tick.png" /></td>
                                    <td class="style2">
                                        <asp:TextBox ID="redate" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                                        <asp:ImageButton ID="redate1" runat="server"
                                            ImageUrl="~/images/calendr.gif" />
                                        <font color="red">Ex:2011/01/01</font></td>
                                </tr>
                            </tr>
                            <tr>
                                <td class="style1">檢核表</td>
                                <td class="style2">
                                    <asp:HyperLink ID="HyperLink1" runat="server">HyperLink</asp:HyperLink>

                            </tr>
                            <tr>
                                <td class="style1">圖片檢視請直接點選(開新視窗)</td>
                                <td class="style2">
                                   <%--<asp:HyperLink ID="HyperLink0" runat="server" target="_blank"> <asp:Image ID="IMG1" runat="server" Width="100px" Height="100px" /></asp:HyperLink>--%>
                                    <asp:HyperLink ID="HyperLink2" runat="server" target="_blank"> <asp:Image ID="IMG2" runat="server" Width="100px" Height="100px" /></asp:HyperLink>
                                    <asp:HyperLink ID="HyperLink3" runat="server" target="_blank"> <asp:Image ID="IMG3" runat="server" Width="100px" Height="100px" /></asp:HyperLink>
                                    <asp:HyperLink ID="HyperLink4" runat="server" target="_blank"> <asp:Image ID="IMG4" runat="server" Width="100px" Height="100px" /></asp:HyperLink>
                                   <asp:HyperLink ID="HyperLink5" runat="server" target="_blank">  <asp:Image ID="IMG5" runat="server" Width="100px" Height="100px" /></asp:HyperLink>
                                   <asp:HyperLink ID="HyperLink6" runat="server" target="_blank">  <asp:Image ID="IMG6" runat="server" Width="100px" Height="100px" /></asp:HyperLink>
                                  <asp:HyperLink ID="HyperLink7" runat="server" target="_blank">   <asp:Image ID="IMG7" runat="server" Width="100px" Height="100px" /></asp:HyperLink>
 <asp:HyperLink ID="HyperLink8" runat="server" target="_blank">   <asp:Image ID="IMG8" runat="server" Width="100px" Height="100px" /></asp:HyperLink>
 <asp:HyperLink ID="HyperLink9" runat="server" target="_blank">   <asp:Image ID="IMG9" runat="server" Width="100px" Height="100px" /></asp:HyperLink>
    <asp:HyperLink ID="HyperLink10" runat="server" target="_blank">   <asp:Image ID="IMG10" runat="server" Width="100px" Height="100px" /></asp:HyperLink>
   <asp:HyperLink ID="HyperLink11" runat="server" target="_blank">   <asp:Image ID="IMG11" runat="server" Width="100px" Height="100px" /></asp:HyperLink>
      <asp:HyperLink ID="HyperLink12" runat="server" target="_blank">   <asp:Image ID="IMG12" runat="server" Width="100px" Height="100px" /></asp:HyperLink>
   <asp:HyperLink ID="HyperLink13" runat="server" target="_blank">   <asp:Image ID="IMG13" runat="server" Width="100px" Height="100px" /></asp:HyperLink>

 <asp:HyperLink ID="HyperLink14" runat="server" target="_blank">   <asp:Image ID="IMG14" runat="server" Width="100px" Height="100px" /></asp:HyperLink>
  <asp:HyperLink ID="HyperLink15" runat="server" target="_blank">   <asp:Image ID="IMG15" runat="server" Width="100px" Height="100px" /></asp:HyperLink>
 <asp:HyperLink ID="HyperLink16" runat="server" target="_blank">   <asp:Image ID="IMG16" runat="server" Width="100px" Height="100px" /></asp:HyperLink>
 <asp:HyperLink ID="HyperLink17" runat="server" target="_blank">   <asp:Image ID="IMG17" runat="server" Width="100px" Height="100px" /></asp:HyperLink>
                

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
