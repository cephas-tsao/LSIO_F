<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="management_BDP002A_Default" MaintainScrollPositionOnPostback="true" %>
<%@ Register src="../../usercontrol/WebUserMenu.ascx" tagname="WebUserMenu" tagprefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title><%=ConfigurationManager.AppSettings("Prj_name")%></title>
    <link rel="stylesheet" href="https://pro.fontawesome.com/releases/v5.10.0/css/all.css" integrity="sha384-AYmEC3Yw5cVb3ZcuHtOA93w35dYTsvhLPVnYs9eStHfGJvOvKxVfELGroGkvsg+p" crossorigin="anonymous"/><link rel='stylesheet' href ="/../Setting/style<%=Get_SysPara("css_sys", "01")%>.css" type = 'text/css'/>
    <script language="javascript">
      function Chk_All_E(){
        //var tmp_flag = confirm("確定(全選),取消(全不選)");
        for (var i = 0; i < document.form1.length; i++) {
          if (document.form1.elements[i].type == 'checkbox'){
            if (document.form1.elements[i].name.charAt(0) == 'E') {
              //if (tmp_flag==true){
              if (document.forms[0].elements["check_box_E"].checked == true ){
//              if (document.forms[0].elements["check_button_E"].value == '欄全選') {
                document.form1.elements[i].checked = true;
              }
              else{
                document.form1.elements[i].checked = false;
              }
            }
          }
        }
//        if (document.forms[0].elements["check_button_E"].value == '欄全選') {
//            document.forms[0].elements["check_button_E"].value = '欄取消';
//        }
//        else {
//            document.forms[0].elements["check_button_E"].value = '欄全選';
//        }
      }     
      
      function Chk_All_A(){
        //var tmp_flag = confirm("確定(全選),取消(全不選)");
        for (var i = 0; i < document.form1.length; i++) {
          if (document.form1.elements[i].type == 'checkbox'){
            if (document.form1.elements[i].name.charAt(0) == 'A') {
              //if (tmp_flag==true){
              if (document.forms[0].elements["check_box_A"].checked == true ){
//                if (document.forms[0].elements["check_button_A"].value == '欄全選') {
                document.form1.elements[i].checked = true;
              }
              else{
                document.form1.elements[i].checked = false;
              }
            }
          }
        }
//        if (document.forms[0].elements["check_button_A"].value == '欄全選') {
//            document.forms[0].elements["check_button_A"].value = '欄取消';
//        }
//        else {
//            document.forms[0].elements["check_button_A"].value = '欄全選';
//        }
      }  
      
      function Chk_All_M(){
        //var tmp_flag = confirm("確定(全選),取消(全不選)");
        for (var i = 0; i < document.form1.length; i++) {
          if (document.form1.elements[i].type == 'checkbox'){
            if (document.form1.elements[i].name.charAt(0) == 'M') {
              //if (tmp_flag==true){
              if (document.forms[0].elements["check_box_M"].checked == true ){
//                if (document.forms[0].elements["check_button_M"].value == '欄全選') {
                document.form1.elements[i].checked = true;
              }
              else{
                document.form1.elements[i].checked = false;
              }
            }
          }
        }
//        if (document.forms[0].elements["check_button_M"].value == '欄全選') {
//            document.forms[0].elements["check_button_M"].value = '欄取消';
//        }
//        else {
//            document.forms[0].elements["check_button_M"].value = '欄全選';
//        }
      }  
      
      function Chk_All_D(){
        //var tmp_flag = confirm("確定(全選),取消(全不選)");
        for (var i = 0; i < document.form1.length; i++) {
          if (document.form1.elements[i].type == 'checkbox'){
            if (document.form1.elements[i].name.charAt(0) == 'D') {
              //if (tmp_flag==true){
              if (document.forms[0].elements["check_box_D"].checked == true ){
//                if (document.forms[0].elements["check_button_D"].value == '欄全選') {
                document.form1.elements[i].checked = true;
              }
              else{
                document.form1.elements[i].checked = false;
              }
            }
          }
        }
//        if (document.forms[0].elements["check_button_D"].value == '欄全選') {
//            document.forms[0].elements["check_button_D"].value = '欄取消';
//        }
//        else {
//            document.forms[0].elements["check_button_D"].value = '欄全選';
//        }
      }  
      
      function Chk_All_S(){
        //var tmp_flag = confirm("確定(全選),取消(全不選)");
        for (var i = 0; i < document.form1.length; i++) {
          if (document.form1.elements[i].type == 'checkbox'){
            if (document.form1.elements[i].name.charAt(0) == 'S') {
              //if (tmp_flag==true){
              if (document.forms[0].elements["check_box_S"].checked == true ){
//                if (document.forms[0].elements["check_button_S"].value == '欄全選') {
                document.form1.elements[i].checked = true;
              }
              else{
                document.form1.elements[i].checked = false;
              }
            }
          }
        }
//        if (document.forms[0].elements["check_button_S"].value == '欄全選') {
//            document.forms[0].elements["check_button_S"].value = '欄取消';
//        }
//        else {
//            document.forms[0].elements["check_button_S"].value = '欄全選';
//        }
    }
    
    function Chk_Row(id) {
        //var tmp_flag = confirm("確定(全選),取消(全不選)");
        for (var i = 0; i < document.form1.length; i++) {
            if (document.form1.elements[i].type == 'checkbox') {
                if (document.form1.elements[i].name.substr(2) == id) {
                    //if (tmp_flag==true){
                    if (document.forms[0].elements["C|" + id].checked == true) {
//                    if (document.forms[0].elements["BT|" + id].value == '行全選') {
                        document.form1.elements[i].checked = true;
                    }
                    else {
                        document.form1.elements[i].checked = false;
                    }
                }
            }
        }
//        if (document.forms[0].elements["BT|" + id].value == '行全選') {
//            document.forms[0].elements["BT|" + id].value = '行取消';
//        }
//        else {
//            document.forms[0].elements["BT|" + id].value = '行全選';
//        }
    } 
            
     </script>
</head>
<body>
    <form id="form1" runat="server">

     <table style="position:relative; margin-left:auto; margin-right:auto; width:990px; "  border ="0">
     <tr>
         <td>
           <div>
                <uc1:WebUserMenu ID="WebUserMenu" runat="server" />
            </div>
         </td>
     </tr>
     <tr>
        <td>
        
    <div>
        選擇角色<asp:DropDownList ID="DDL1" runat="server">
        </asp:DropDownList>
        <asp:Button ID="Button2" runat="server" Text="讀取設定" />
        <asp:Button ID="Button3" runat="server" Text="複製權限" Visible="False" />
        <asp:Button ID="Bt_save" runat="server" Text="存檔" OnClick="Save_User_Data"
            Visible="False" />
        <asp:Button ID="Bt_cancel" runat="server" Text="取消" OnClick="Button_Cancel"
            Visible="False" />
        <asp:Panel ID="Panel1" runat="server" Height="50px" Width="100%">
            <asp:Table ID="Table1" runat="server" Width="100%" Visible="False" />
        <asp:Button ID="Button1" runat="server" Text="存檔" OnClick="Save_User_Data"
             Visible="False" />
        <asp:Button ID="Button4" runat="server" Text="取消" OnClick="Button_Cancel"
            Visible="False" />
         </asp:Panel>
    </div>
        </td>
    </tr>
    </table>
    <asp:SqlDataSource ID="SqlDS1" runat="server" />
    </form>
</body>
</html>
