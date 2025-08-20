<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default3.aspx.vb" Inherits="basic_LSI028_Default3" %>

<!DOCTYPE html>
<%@ Register src="../../usercontrol/WebUserMenu.ascx" tagname="WebUserMenu" tagprefix="uc1" %>
<%@ Register src="../../usercontrol/WebSubMenu.ascx" tagname="WebSubMenu" tagprefix="uc2" %>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
   <title><%=ConfigurationManager.AppSettings("Prj_name")%></title>


<script src="js/jquery.min.js"></script>
<script type="text/javascript" defer src="js/alphafilter.js"></script>

<link rel="stylesheet" type="text/css" media="all" href="js/calendar/calendar-blue.css" title="win2k-cold-1" />
<script type="text/javascript" src="js/calendar/calendar.js"></script>
<script type="text/javascript" src="js/calendar/lang/calendar-en.js"></script>
<script type="text/javascript" src="js/calendar/calendar-setup.js"></script>
    
 <script src="js/jquery.table2excel.js"></script>
 
	
	<script language="javascript">

	    function xls() {
	        $(".table2excel").table2excel({
	            exclude: ".noExl",
	            name: "Excel Document Name",
	            filename: "" + new Date().toISOString().replace(/[\-\:\.]/g, ""),
	            fileext: ".xls",
	            exclude_img: true,
	            exclude_links: true,
	            exclude_inputs: true
	        });
	    }


</script>
    </head>
<body>
    <form id="form1" runat="server">
           <div id="anticfrs" runat="server"></div>
          <table id="customers" style="position:relative; margin-left:auto; margin-right:auto; width:990px;"  border ="0">
       <tr>
         <td width="100%">
           <div>
                <uc1:WebUserMenu ID="WebUserMenu" runat="server" />
           </div>
         </td>
     </tr>
    </table>
   <div align="center" style="width: 990px;border:1px #000000 solid;margin:0px auto;">
	  <h3>自主輔導檢查申報-各季統計</h3>																				
      <table width="100%" border="0" id="tbData"  data-tableName="Test Table 1">
        <tbody>
          <tr>
            <td width="30%" align="center"><input type="button" style="width: 100px; font-size: 14px; background-color:darkgreen; color: #fff;"  onclick="window.location.href = 'Default.aspx'" value="回首頁" />
&nbsp;&nbsp;
<%--<input type="button" style="width: 100px; font-size: 14px; background-color:olivedrab; color: #fff;"  onclick="window.location.href = 'Default-1.aspx'" value="勞條科" />--%>

            </td>
            <td width="30%" align="center">
                <label id="lblSites" runat="server" class="lblFieldDesc" >請選擇年度單位:</label>
                <asp:DropDownList
                        ID="DDLYear" runat="server"
                        DataValueField="field_year">
                        <asp:ListItem Text="110" Value="110"></asp:ListItem>
                        <asp:ListItem Text="111" Value="111" Selected="True" ></asp:ListItem>
                        <asp:ListItem Text="112" Value="112"></asp:ListItem>
                    </asp:DropDownList>
                 <asp:DropDownList ID="drpSites" runat="server"    
            OnSelectedIndexChanged="drpSites_SelectedIndexChanged" >
                      <asp:ListItem Selected>-請選擇-</asp:ListItem>
            <asp:ListItem>一般科</asp:ListItem>
            <asp:ListItem>勞條科</asp:ListItem>          
        </asp:DropDownList></td>
                <td width="20%" align="center">
                       <asp:Button ID="search" runat="server" Text="計算" OnClick="search_Click" />
                    </td>
              <td width="20%" align="center">
               
&nbsp;&nbsp;
<input type="button"  onclick="xls()" value="報表匯出" /></td>
          </tr>
        </tbody>
      </table>
      <br />
      <table width="95%" class="table2excel" border="0" align="center" cellpadding="1" cellspacing="1" >
        <tbody>
    <tr style="color: #fff;">
      <td height="30" align="center" bgcolor="#39A96B">&nbsp;</td>
      <td align="center" bgcolor="#39A96B"><div id="lbyear" runat="server"/>第1季</td>
      <td align="center" bgcolor="#39A96B"><div id="lbyear1" runat="server"/>第2季</td>
	  <td align="center" bgcolor="#39A96B"><div id="lbyear2" runat="server"/>第3季</td>
	  <td align="center" bgcolor="#39A96B"><div id="lbyear3" runat="server"/>第4季</td>
      </tr>
			
 <tr>
      <td height="44" align="center" bgcolor="#9AEB65" >自主管理家數</td>
      <td align="center" bgcolor="#9AEB65" ><input type="text" id="total1" name="total1" runat="server" /></td>
      <td align="center" bgcolor="#9AEB65" ><input type="text" id="total2" name="total2" runat="server" /></td>
      <td align="center" bgcolor="#9AEB65" ><input type="text" id="total3" name="total3" runat="server" /></td>
      <td align="center" bgcolor="#9AEB65" ><input type="text" id="total4" name="total4" runat="server" /></td>
    </tr>
    <tr>
      <td height="44" align="center" bgcolor="#9AEB65" >繳回表單之自主管理家數</td>
      <td align="center" bgcolor="#9AEB65" ><input type="text" id="num_1" name="num1" runat="server" /></td>
      <td align="center" bgcolor="#9AEB65" ><input type="text" id="num_2" name="num2" runat="server" /></td>
      <td align="center" bgcolor="#9AEB65" ><input type="text" id="num_3" name="num3" runat="server" /></td>
      <td align="center" bgcolor="#9AEB65" ><input type="text" id="num_4" name="num4" runat="server" /></td>
      </tr>
   
    <tr>
      <td height="44" align="center" bgcolor="#9AEB65" >繳交比率</td>
      <td align="center" bgcolor="#9AEB65" ><input type="text" id="rate1" name="rate1" runat="server" /></td>
      <td align="center" bgcolor="#9AEB65" ><input type="text" id="rate2" name="rate2" runat="server" /></td>
      <td align="center" bgcolor="#9AEB65" ><input type="text" id="rate3" name="rate3" runat="server" /></td>
      <td align="center" bgcolor="#9AEB65" ><input type="text" id="rate4" name="rate4" runat="server" /></td>
    </tr>

					
   
  </tbody>
</table>
<p><br />
</p>
	
																						
</div>
    </form>
</body>
</html>
