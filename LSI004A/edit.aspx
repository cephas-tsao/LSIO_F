<%@ Page Language="VB" AutoEventWireup="false" CodeFile="edit.aspx.vb" Inherits="basic_LSI004A_edit" MaintainScrollPositionOnPostback="true" validateRequest="False" %>
<%@ Register src="../../usercontrol/WebUserMenu.ascx" tagname="WebUserMenu" tagprefix="uc1" %>
<%@ Register src="detail_edit.ascx" tagname="detail_edit" tagprefix="uc2" %>
<%@ Register assembly="FreeTextBox" namespace="FreeTextBoxControls" tagprefix="FTB" %>
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
                <td class="style1">資料編號</td>
                <td class="style2" style="width: 100%">
                    <asp:TextBox ID="txt_paper_code" runat="server" Width="120px" CssClass="tb2"></asp:TextBox>
                    <font color="red">系統自動取號</font></td>
            </tr>
            <tr>
                <td class="style1">電子報名稱<asp:Image ID="Image3" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2" style="width: 35%">
                    <asp:TextBox ID="txt_paper_name" runat="server" MaxLength="100" Width="400px"></asp:TextBox></td>
            </tr>
                                 
            <tr>
                <td class="style1">電子報期數<asp:Image ID="Image4" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2" style="width: 35%">
                    <asp:TextBox ID="txt_paper_num" runat="server" MaxLength="4" Width="80px"></asp:TextBox></td>
            </tr>                            
                                 
            <tr>
                <td class="style1">電子報類型</td>
                <td class="style2" style="width: 35%">
                    <asp:DropDownList ID="ddl_paper_type" runat="server">
                    </asp:DropDownList>
                                </td>
            </tr>                            
            <tr>
                <td class="style1">電子報抬頭圖片<asp:Image ID="Image1" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2" style="width: 35%">
                    <asp:FileUpload ID="fu_paper_logo" runat="server" />
                    <asp:Label ID="l_paper_logo" runat="server"></asp:Label>
                                </td>
            </tr>         
            <tr>
                <td class="style1">樣式區<asp:Image ID="Image2" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2" style="width: 35%">
                                                    <asp:DropDownList ID="ddl_paper_css" runat="server">
                                                    </asp:DropDownList>
                                                    </td>
            </tr>                      
            <tr>
                <td class="style1">電子報表尾1</td>
                <td class="style2" style="width: 35%">
                    <%--<asp:TextBox ID="txt_paper_bottom1" runat="server" Width="800px" Height="100px" 
                        TextMode="MultiLine"></asp:TextBox>--%>
                    <FTB:FreeTextBox ID="txt_paper_bottom1" runat="server" 
                        Height="100px" Width="100%" Language="zh-TW" BreakMode="LineBreak">
                    </FTB:FreeTextBox></td>
            </tr> 
            <tr>
                <td class="style1">電子報表尾2</td>
                <td class="style2" style="width: 35%">
                    <%--<asp:TextBox ID="txt_paper_bottom2" runat="server" Width="800px" Height="100px" 
                        TextMode="MultiLine"></asp:TextBox>--%>
                    <FTB:FreeTextBox ID="txt_paper_bottom2" runat="server" 
                        Height="100px" Width="100%" Language="zh-TW" BreakMode="LineBreak">
                    </FTB:FreeTextBox></td>
            </tr> 
            <tr>
                <td class="style1">使用者姓名</td>
                <td class="style2">
                    <asp:TextBox ID="txt_usr_code" runat="server" MaxLength="20" 
                        AutoPostBack="True"></asp:TextBox>&nbsp;<asp:ImageButton
                        ID="ib_usr_code" runat="server" CausesValidation="False" 
                        ImageUrl="~/images/bs_search.gif" />
                    <asp:TextBox ID="txt_usr_name" runat="server" Enabled="False"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">建立日期</td>
                <td class="style2">
                    <asp:TextBox ID="txt_ins_date" runat="server" MaxLength="10" Width="100px" 
                        BackColor="#999999"></asp:TextBox></td>
            </tr>
        </table>
    </div>
            <asp:Button ID="Button1" runat="server" Text="主檔存檔" />
            <asp:Button ID="Bt_BackUp" runat="server" Text="離開" />
            <asp:Button ID="Bt_preview" runat="server" Text="視覺化模式預覽" Visible="False" />
            <asp:TextBox ID="txtPgmType" runat="server" Visible="False" />
            <asp:TextBox ID="txtPrimary" runat="server" Visible="False" />
            <asp:TextBox ID="txtSubWhere" runat="server" Visible="False" />
            <asp:TextBox ID="txtOrder" runat="server" Visible="False" />
            <asp:TextBox ID="txt_paper_logo" runat="server" Visible="False" />
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
