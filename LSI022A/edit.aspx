<%@ Page Language="VB" AutoEventWireup="false" CodeFile="edit.aspx.vb" Inherits="basic_LSI002A_edit" aspcompat="true" MaintainScrollPositionOnPostback="true" %>
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
                <td class="fail" colspan="6" style="text-align: center">
                    <asp:Label ID="Label1" runat="server" Font-Bold="True" Style="text-align: center"></asp:Label></td>
            </tr>
            
            <tr>
                <td class="style1">
                    受理日期</td>
                <td class="style2" colspan="5">
                    <asp:TextBox ID="ins_date" runat="server" MaxLength="12" Width="120px" 
                        Enabled="False"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="style1">
                    受理時間</td>
                <td class="style2" colspan="5">
                    <asp:TextBox ID="cmp_date" runat="server" MaxLength="12" Width="120px" 
                        Enabled="False"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="style1">
                    案件編號<asp:Image ID="Image2" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2" colspan="5">
                    <asp:TextBox ID="txt_cmp_code" runat="server" MaxLength="12" Width="120px" 
                        Enabled="False"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="style1">
                    信件內容<asp:Image ID="Image1" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2" colspan="5">
                    <asp:TextBox ID="txt_cmp_memo" runat="server" Width="624px" Height="150px" 
                        TextMode="MultiLine"></asp:TextBox></td>
            </tr>    
            <tr>
                <td class="style1">
                    姓名</td>
                <td class="style2" colspan="5">
                    <asp:TextBox ID="txt_cmp_name" runat="server" MaxLength="16" Width="200px"></asp:TextBox>
                    <font color=red></font></td>
            </tr>
           
            <tr>
                <td class="style1">
                    電子信箱E-Mail</td>
                <td class="style2" colspan="5">
                    <asp:TextBox ID="txt_cmp_email" runat="server" MaxLength="100" Width="500px"></asp:TextBox>
                    <asp:Button ID="Bt_send" runat="server" Text="補發" Visible="true" />
                    <font color=red></font></td>
            </tr>
            <tr>
                <td class="style1">
                    聯絡電話</td>
                <td class="style2" colspan="5">
                    <asp:TextBox ID="txt_cmp_tel1" runat="server" MaxLength="16" Width="160px"></asp:TextBox>
                    <font color=red></font></td>
            </tr>
           
            <!--<tr>
                <td class="style1">
                    身份證字號</td>
                <td class="style2" colspan="5">
                                <asp:TextBox ID="txt_cmp_idno" runat="server" MaxLength="10" 
                        Width="100px"></asp:TextBox>
                    <font color=red>(申訴類別為具名時必填)</font></td>
            </tr>-->
            
            
            
           
            <!--<tr>
                <td class="style1">
                    個資限制</td>
                <td class="style2" colspan="5">
                    <asp:DropDownList ID="ddl_is_agree" runat="server">
                    </asp:DropDownList>是否同意本處繼續使用資料
                    <font color=red>(申訴類別為具名時必填)</font></td>
            </tr>-->
              <tr >
                <td class="style1">
                    附檔</td>
                <td class="style2">
                    <asp:Label ID="lab_cmp_file1" runat="server" Text="附檔1" ></asp:Label>
                                </td>
                <td class="style2">
                    <asp:Label ID="lab_cmp_file2" runat="server" Text="附檔2"></asp:Label>
                                </td>
                <td class="style2">
                    <asp:Label ID="lab_cmp_file3" runat="server" Text="附檔3"></asp:Label>
                                </td>
                <td class="style2">
                    <asp:Label ID="lab_cmp_file4" runat="server" Text="附檔4"></asp:Label>
                                </td>
                <td class="style2">
                    <asp:Label ID="lab_cmp_file5" runat="server" Text="附檔5"></asp:Label>
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
            <asp:TextBox ID="txtFileName1" runat="server" Visible="False" />
            <asp:TextBox ID="txtFileName2" runat="server" Visible="False" />
            <asp:TextBox ID="txtFileName3" runat="server" Visible="False" />
            <asp:TextBox ID="txtFileName4" runat="server" Visible="False" />
            <asp:TextBox ID="txtFileName5" runat="server" Visible="False" />
            <asp:TextBox ID="cmp_idclass" runat="server" Visible="False" />
            <asp:TextBox ID="com_disputetype" runat="server" Visible="False" />
            <asp:TextBox ID="com_Mediationtype" runat="server" Visible="False" />
            <asp:TextBox ID="txtVerCode" runat="server" Visible="False" />
            <asp:TextBox ID="com_sectype" runat="server" Visible="False" />
            <asp:TextBox ID="com_sex" runat="server" Visible="False" />
            <br />
            <asp:Label ID="l_ErrMsg" runat="server" Font-Bold="True" ForeColor="Red" />
            </td>
        </tr>
      </table>    
    </form>
</body>
</html>
