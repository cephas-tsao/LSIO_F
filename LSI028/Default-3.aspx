<%@ Page Language="VB" AutoEventWireup="false"  validateRequest="False" ContentType="text/html" ResponseEncoding="utf-8"  Debug="true"  %>

<%
response.Buffer = true
session.Codepage =65001
response.Charset = "utf-8"
%>
<%@ Import Namespace = "System.Data" %>
<%@ Import Namespace = "System.Data.SqlClient" %>
<%@ Import NameSpace = "System.IO"%>
<%@ Import NameSpace = "System.Web.Configuration" %>
<% 
Dim Conn As SqlConnection = New SqlConnection
      Conn.ConnectionString = WebConfigurationManager.ConnectionStrings("AppSysConnectionString").ConnectionString
%>
	
<%@ Register src="../../usercontrol/WebUserMenu.ascx" tagname="WebUserMenu" tagprefix="uc1" %>
<%@ Register src="../../usercontrol/WebSubMenu.ascx" tagname="WebSubMenu" tagprefix="uc2" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">



<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
<title><%=ConfigurationManager.AppSettings("Prj_name")%></title>


<script src="js/jquery.min.js"></script>
<script type="text/javascript" defer src="js/alphafilter.js"></script>

<link rel="stylesheet" type="text/css" media="all" href="js/calendar/calendar-blue.css" title="win2k-cold-1" />
<script type="text/javascript" src="js/calendar/calendar.js"></script>
<script type="text/javascript" src="js/calendar/lang/calendar-en.js"></script>
<script type="text/javascript" src="js/calendar/calendar-setup.js"></script>
    
 <script src="js/jquery.table2excel.js"></script>
 
	
	<script language="javascript">

function xls(){
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
<script language="javascript">
function Chk_All_1() {
    for (var i = 0; i < document.form1.length; i++) {
        if (document.form1.elements[i].type == 'checkbox') {
            if (document.form1.elements[i].name.charAt(0) == 'C') {
                if (document.forms[0].elements["CBL_Field1_0"].checked == true) {
                    document.form1.elements[i].checked = true;
                }
                else {
                    document.form1.elements[i].checked = false;
                }
            }
        }
    }
}
</SCRIPT>
</head>
	
	 <%

 Dim pic2,seq1,seq2,seq3,seq4,seq5,seq6,seq7,seq8,seqtt
  Dim p As String = Request("p")
  Dim snum As String = Request("snum")
  Dim c As String = Request("c") 
  Dim IDA As String = Request("ID") 
  
	Dim haveRec = False

        '-- p 就是「目前在第幾頁?」    
   Conn.open()
	     Dim SqlStra1 As String
	     Dim SqlStra2 As String
	     Dim SqlStra3 As String
	     Dim SqlStra4 As String
	     Dim SqlStra5, SqlStra6, SqlStra7, SqlStra8, SqlStra9, SqlStra10 As String
  
	     SqlStra1 = "SELECT  *  FROM LSI026 where pc_type = N'一般科' and pc_season=1"
	     SqlStra2 = "SELECT  *  FROM LSI026 where pc_type = N'一般科' and pc_season=2"
	     SqlStra3 = "SELECT  *  FROM LSI026 where pc_type = N'一般科' and pc_season=3"
	     SqlStra4 = "SELECT  *  FROM LSI026 where pc_type = N'一般科' and pc_season=4"
	     
	     SqlStra5 = "SELECT  *  FROM Self_com where unit = N'一般科' "
	     SqlStra6 = "SELECT  *  FROM Self_com where unit = N'勞條科' "
	

	     Dim da1 As SqlDataAdapter = New SqlDataAdapter(SqlStra1, Conn)
	     Dim da2 As SqlDataAdapter = New SqlDataAdapter(SqlStra2, Conn)
	     Dim da3 As SqlDataAdapter = New SqlDataAdapter(SqlStra3, Conn)
	     Dim da4 As SqlDataAdapter = New SqlDataAdapter(SqlStra4, Conn)
	     Dim da5 As SqlDataAdapter = New SqlDataAdapter(SqlStra5, Conn)
	     Dim da6 As SqlDataAdapter = New SqlDataAdapter(SqlStra6, Conn)
	     
	     Dim DS1 As DataSet = New DataSet
	     Dim DS2 As DataSet = New DataSet
	     Dim DS3 As DataSet = New DataSet
	     Dim DS4 As DataSet = New DataSet
	     Dim DS5 As DataSet = New DataSet
	     Dim DS6 As DataSet = New DataSet
	     da1.Fill(DS1)    '-- 把SQL指令執行完成的結果，填入DataSet裡面。
	     da2.Fill(DS2)    '-- 把SQL指令執行完成的結果，填入DataSet裡面。
	     da3.Fill(DS3)    '-- 把SQL指令執行完成的結果，填入DataSet裡面。
	     da4.Fill(DS4)    '-- 把SQL指令執行完成的結果，填入DataSet裡面。
	     da5.Fill(DS5)    '-- 把SQL指令執行完成的結果，填入DataSet裡面。
	     da6.Fill(DS6)    '-- 把SQL指令執行完成的結果，填入DataSet裡面。

	     Dim DT1 As DataTable = DS1.Tables(0)
	     Dim DT2 As DataTable = DS2.Tables(0)
	     Dim DT3 As DataTable = DS3.Tables(0)
	     Dim DT4 As DataTable = DS4.Tables(0)
	     Dim DT5 As DataTable = DS5.Tables(0)
	     Dim DT6 As DataTable = DS6.Tables(0)
	     '=============  ADO.NET / DataSet ==(End)======
	     Dim num1 As Integer = 0
	     Dim num2 As Integer = 0
	     Dim num3 As Integer = 0
	     Dim num4 As Integer = 0
	     Dim num5 As Integer = 0
	     Dim num6 As Integer = 0
	     num1 = DT1.Rows.Count()
	     num2 = DT2.Rows.Count()
	     num3 = DT3.Rows.Count()
	     num4 = DT4.Rows.Count()
	     num5 = DT5.Rows.Count()
	     num6 = DT6.Rows.Count()
	     
	     num_1.Value = num1 & "%"
	     num_2.Value = num2 & "%"
	     num_3.Value = num3 & "%"
	     num_4.Value = num4 & "%"
	     
	     total1.Value = num5
	     total2.Value = num5
	     total3.Value = num5
	     total4.Value = num5
	   
	     rate1.Value = Int((num1 / num5) * 100) & "%"
	     rate2.Value = Int((num2 / num5) * 100) & "%"
	     rate3.Value = Int((num3 / num5) * 100) & "%"
	     rate4.Value = Int((num4 / num5) * 100) & "%"
	     '---每頁展示 5 筆資料
        Dim PageSize As Integer = 100

        '--SQL指令共撈到多少筆（列）資料。RecordCount資料總筆（列）數
	     'Dim RecordCount As Integer = 0
	     'RecordCount = DT.Rows.Count()

	     '--SQL指令共撈到多少筆（列）資料。RecordCount資料總筆（列）數
	     Dim RecordCount As Integer = 1
	     'RecordCount = DT1.Rows.Count()

	     '--如果撈不到資料，程式就結束。-- Start --------------
	     If RecordCount = 0 Then
	         'Response.Write("<h2>已送訂單是空的</h2>")
	         'Conn.Close()
	         Response.End()
	     End If    '--如果撈不到資料，程式就結束。-- End --


	     '--Pages 資料的總頁數。搜尋到的所有資料，共需「幾頁」才能全部呈現？
	     Dim Pages As Integer = 0
	     Pages = ((RecordCount + PageSize) - 1) \ PageSize

	     '--底下這一段IF判別式，是用來防呆，防止一些例外狀況。-- start --
	     '--有任何問題，就強制跳回第一頁（p=1）。
	     If IsNumeric(Request("p")) Then
	         '--頁數（p）務必是一個整數。而且需要大於零、比起「資料的總頁數」要少
	         If Request("p") <> "" And CInt(Request("p")) > 0 And CInt(Request("p")) <= Pages Then
	             p = CInt(Request("p"))
	         Else
	             p = 1
	         End If
	     Else
	         p = 1
	     End If  '--上面這一段IF判別式，是用來防呆，防止一些例外狀況。-- end --

	     '--NowPageCount，目前這頁的資料，要從 DataSet裡面的第幾筆（列）開始撈資料？？
	     Dim NowPageCount As Integer = 0
	     If (p > 0) Then
	         NowPageCount = (p - 1) * PageSize    '--PageSize，每頁展示5筆資料（上面設定過了）
	     End If

	     'Response.Write("<h3>已送訂單筆數共計" & RecordCount & "筆</h3>")
	  
	     '--rowNo，目前畫面出現的這一頁，要撈出幾筆（列）資料
	     Dim rowNo As Integer = 0
	     Dim html As String = ""
	     
		%>
   
		 
<body>
    <form id="form1" runat="server">
			
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
            <td width="30%" align="center"><input type="button" style="width: 100px; font-size: 14px; background-color:darkgreen; color: #fff;"  onclick="window.location.href='Default.aspx'" value="回首頁" />
&nbsp;&nbsp;
<%--<input type="button" style="width: 100px; font-size: 14px; background-color:olivedrab; color: #fff;"  onclick="window.location.href='Default-1.aspx'" value="勞條科" />--%></td>
            <td width="30%" align="center">
            <%--    <label id="lblSites" runat="server" class="lblFieldDesc" >請選擇單位:</label>
                 <asp:DropDownList ID="drpSites" runat="server" AutoPostBack="True" >
            <asp:ListItem>一般科</asp:ListItem>
            <asp:ListItem>勞條科</asp:ListItem>          
        </asp:DropDownList>--%>

            </td>
              <td width="40%" align="center">
                  	</form>
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
      <td align="center" bgcolor="#39A96B">110第1季</td>
      <td align="center" bgcolor="#39A96B">110第2季</td>
	  <td align="center" bgcolor="#39A96B">110第3季</td>
	  <td align="center" bgcolor="#39A96B">110第4季</td>
      </tr>
			
		  <%

	While ((rowNo < PageSize) And (NowPageCount < RecordCount))
            haveRec = True
	%> 
    <tr>
      <td height="44" align="center" <%if rowNo mod 2 = 0 then%>bgcolor="#9AEB65" <%else%> bgcolor="#D9F86B"<%end if%> >已通報單位</td>
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="#9AEB65" <%else%> bgcolor="#D9F86B"<%end if%>><input type="text" id="num_1" name="num1" runat="server" /></td>
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="#9AEB65" <%else%> bgcolor="#D9F86B"<%end if%>><input type="text" id="num_2" name="num2" runat="server" /></td>
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="#9AEB65" <%else%> bgcolor="#D9F86B"<%end if%>><input type="text" id="num_3" name="num3" runat="server" /></td>
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="#9AEB65" <%else%> bgcolor="#D9F86B"<%end if%>><input type="text" id="num_4" name="num4" runat="server" /></td>
      </tr>
    <tr>
      <td height="44" align="center" <%if rowNo mod 2 = 0 then%>bgcolor="#9AEB65" <%else%> bgcolor="#D9F86B"<%end if%> >應通報總數</td>
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="#9AEB65" <%else%> bgcolor="#D9F86B"<%end if%>><input type="text" id="total1" name="total1" runat="server" /></td>
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="#9AEB65" <%else%> bgcolor="#D9F86B"<%end if%>><input type="text" id="total2" name="total2" runat="server" /></td>
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="#9AEB65" <%else%> bgcolor="#D9F86B"<%end if%>><input type="text" id="total3" name="total3" runat="server" /></td>
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="#9AEB65" <%else%> bgcolor="#D9F86B"<%end if%>><input type="text" id="total4" name="total4" runat="server" /></td>
    </tr>
    <tr>
      <td height="44" align="center" <%if rowNo mod 2 = 0 then%>bgcolor="#9AEB65" <%else%> bgcolor="#D9F86B"<%end if%> >通報率</td>
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="#9AEB65" <%else%> bgcolor="#D9F86B"<%end if%>><input type="text" id="rate1" name="rate1" runat="server" /></td>
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="#9AEB65" <%else%> bgcolor="#D9F86B"<%end if%>><input type="text" id="rate2" name="rate2" runat="server" /></td>
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="#9AEB65" <%else%> bgcolor="#D9F86B"<%end if%>><input type="text" id="rate3" name="rate3" runat="server" /></td>
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="#9AEB65" <%else%> bgcolor="#D9F86B"<%end if%>><input type="text" id="rate4" name="rate4" runat="server" /></td>
    </tr>
			    <%

            NowPageCount = NowPageCount + 1
           rowNo = rowNo + 1
		   
    End While
	%> 
					
   
  </tbody>
</table>
<p><br />
</p>
	
																						
</div>
	
	<%

Conn.close() 
	   %>
</body>

</html>
		
		
		
