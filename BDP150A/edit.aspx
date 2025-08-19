<%@ Page Language="VB" AutoEventWireup="false" CodeFile="edit.aspx.vb" Inherits="management_BDP150A_edit" aspcompat="true" MaintainScrollPositionOnPostback="true" %>
<%@ Register src="../../usercontrol/WebUserMenu.ascx" tagname="WebUserMenu" tagprefix="uc1" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title><%=ConfigurationManager.AppSettings("Prj_name")%></title>
    <script language="javascript" id="js1" src="../../public/checkjs.js"></script> 
    <link rel="stylesheet" href="https://pro.fontawesome.com/releases/v5.10.0/css/all.css" integrity="sha384-AYmEC3Yw5cVb3ZcuHtOA93w35dYTsvhLPVnYs9eStHfGJvOvKxVfELGroGkvsg+p" crossorigin="anonymous"/><link rel='stylesheet' href ="/../Setting/style<%=Get_SysPara("css_sys", "01")%>.css" type = 'text/css'/>
    <script>
        function next(obj, next) {
            if (obj.value.length == obj.maxLength)  //注意此處maxLength的大小寫  
                obj.form.elements[next].focus();
        }  
	</script> 
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
                <td class="style1" style="width: 13%">
                    開始IP<asp:Image ID="Image2" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_ip_s1" runat="server" MaxLength="3" Width="40px" onKeyUp="next(this,'txt_ip_s2')" style="ime-mode: disabled;" onFocus="this.select()" 
                    OnKeyPress="if(((event.keyCode&gt;=48)&amp;&amp;(event.keyCode &lt;=57))||(event.keyCode==46)) {event.returnValue=true;} else{event.returnValue=false;}"></asp:TextBox>
                    .
                    <asp:TextBox ID="txt_ip_s2" runat="server" MaxLength="3" Width="40px" onKeyUp="next(this,'txt_ip_s3')" style="ime-mode: disabled;" onFocus="this.select()" 
                    OnKeyPress="if(((event.keyCode&gt;=48)&amp;&amp;(event.keyCode &lt;=57))||(event.keyCode==46)) {event.returnValue=true;} else{event.returnValue=false;}"
                        Height="19px"></asp:TextBox>
                        .
                    <asp:TextBox ID="txt_ip_s3" runat="server" onKeyUp="next(this,'txt_ip_s4')" style="ime-mode: disabled;" onFocus="this.select()" 
                    OnKeyPress="if(((event.keyCode&gt;=48)&amp;&amp;(event.keyCode &lt;=57))||(event.keyCode==46)) {event.returnValue=true;} else{event.returnValue=false;}"
                        MaxLength="3" Width="40px"></asp:TextBox>
                        .
                        <asp:TextBox ID="txt_ip_s4" style="ime-mode: disabled;" onFocus="this.select()" 
                    OnKeyPress="if(((event.keyCode&gt;=48)&amp;&amp;(event.keyCode &lt;=57))||(event.keyCode==46)) {event.returnValue=true;} else{event.returnValue=false;}"
                        runat="server" MaxLength="3" Width="40px"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="style1" style="width: 13%">
                    結束IP<asp:Image ID="Image3" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:TextBox ID="txt_ip_e1" runat="server" MaxLength="3" Width="40px" style="ime-mode: disabled;" onFocus="this.select()" 
                    OnKeyPress="if(((event.keyCode&gt;=48)&amp;&amp;(event.keyCode &lt;=57))||(event.keyCode==46)) {event.returnValue=true;} else{event.returnValue=false;}"
                        onKeyUp="next(this,'txt_ip_e2')"></asp:TextBox>
                    .
                    <asp:TextBox ID="txt_ip_e2" runat="server" MaxLength="3" Width="40px" style="ime-mode: disabled;" onFocus="this.select()" 
                    OnKeyPress="if(((event.keyCode&gt;=48)&amp;&amp;(event.keyCode &lt;=57))||(event.keyCode==46)) {event.returnValue=true;} else{event.returnValue=false;}"
                        onKeyUp="next(this,'txt_ip_e3')"></asp:TextBox>
                        .
                        <asp:TextBox ID="txt_ip_e3" runat="server" MaxLength="3" Width="40px" onKeyUp="next(this,'txt_ip_e4')" style="ime-mode: disabled;" onFocus="this.select()" 
                    OnKeyPress="if(((event.keyCode&gt;=48)&amp;&amp;(event.keyCode &lt;=57))||(event.keyCode==46)) {event.returnValue=true;} else{event.returnValue=false;}"
                    ></asp:TextBox>
                    .
                    <asp:TextBox ID="txt_ip_e4" runat="server" MaxLength="3" Width="40px" style="ime-mode: disabled;" onFocus="this.select()" 
                    OnKeyPress="if(((event.keyCode&gt;=48)&amp;&amp;(event.keyCode &lt;=57))||(event.keyCode==46)) {event.returnValue=true;} else{event.returnValue=false;}"></asp:TextBox>
                    </td>
            </tr>
            <tr>
                <td class="style1" style="width: 13%">
                    啟用代碼<asp:Image ID="Image4" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2">
                    <asp:DropDownList ID="ddl_ip_type" runat="server">
                    </asp:DropDownList>
                                    </td>
            </tr>
            <tr>
                <td class="style1" style="width: 13%">
                    備註</td>
                <td class="style2">
                    <asp:TextBox ID="txt_ip_name" runat="server" MaxLength="20" Width="200px"></asp:TextBox></td>
                <!--<td width="10%">
                </td>
                <td width="23%">
                </td>-->
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
