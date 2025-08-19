<%@ Page Language="VB" AutoEventWireup="false" CodeFile="edit.aspx.vb" Inherits="basic_LSI016A_edit" aspcompat="true" MaintainScrollPositionOnPostback="true" ValidateRequest="false"  %>
<%@ Register src="../../usercontrol/WebUserMenu.ascx" tagname="WebUserMenu" tagprefix="uc1" %>
<%@ Register assembly="FreeTextBox" namespace="FreeTextBoxControls" tagprefix="FTB" %>
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
    
        <SCRIPT type="text/javascript">
            function storeCaret(textEl) {
                if (textEl.createTextRange)
                    textEl.caretPos = document.selection.createRange().duplicate();
            }
            function insertAtCaret(textEl, text) {
                if (textEl.createTextRange && textEl.caretPos) {
                    var caretPos = textEl.caretPos;
                    caretPos.text =
           caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
             text + ' ' : text;
                }
                else
                    textEl.value = text;
            }
     </SCRIPT>        
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
                    <asp:TextBox ID="txt_news_code" runat="server" MaxLength="12" Width="120px" 
                        Enabled="False"></asp:TextBox>
                    <font color="red">系統自動取號</font></td>
            </tr>
            <tr>
                <td class="style1">
                    新聞抬頭<asp:Image ID="Image3" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_news_title" runat="server" MaxLength="100" Width="500px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    新聞類型<asp:Image ID="Image1" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:RadioButtonList ID="rbl_news_type" runat="server" 
                        RepeatDirection="Horizontal">
                    </asp:RadioButtonList>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    圖片</td>
                <td class="style2">
                    <asp:FileUpload ID="fup_news_pic1" runat="server" />
                    <asp:Label ID="lab_news_pic1" runat="server"></asp:Label>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    本文區<asp:Image ID="Image4" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <%--<asp:TextBox ID="txt_news_contect" runat="server" Width="500px" Height="300px" 
                        TextMode="MultiLine"></asp:TextBox>--%>
                    <FTB:FreeTextBox ID="txt_news_contect" runat="server" 
                        Height="400px" Width="100%" Language="zh-TW" BreakMode="LineBreak">
                    </FTB:FreeTextBox></td>
            </tr>
            <tr>
                <td class="style1">
                    建立者</td>
                <td class="style2">
                    <asp:TextBox ID="txt_usr_code" runat="server" MaxLength="20" Width="200px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    建立日期</td>
                <td class="style2">
                    <asp:TextBox ID="txt_ins_date" runat="server" MaxLength="10" Width="100px" 
                        BackColor="#999999"></asp:TextBox>
                                </td>
            </tr>
            </table>
    </div>
            <asp:Button ID="Button1" runat="server" Text="存檔" />
            <asp:Button ID="Bt_preview" runat="server" Text="預覽" Visible="False" />
            <asp:Button ID="Bt_BackUp" runat="server" Text="離開" />
            <asp:TextBox ID="txtPgmType" runat="server" Visible="False" />
            <asp:TextBox ID="txtPrimary" runat="server" Visible="False" />
            <asp:TextBox ID="txtSubWhere" runat="server" Visible="False" />
            <asp:TextBox ID="txtOrder" runat="server" Visible="False" />
            <asp:TextBox ID="txt_news_pic1" runat="server" Visible="False" />
            <br />
            <asp:Label ID="l_ErrMsg" runat="server" Font-Bold="True" ForeColor="Red" />
            </td>
        </tr>
      </table>    
    </form>
</body>
</html>
