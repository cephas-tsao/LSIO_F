<%@ Page Language="VB" AutoEventWireup="false"  validateRequest="False" ContentType="text/html" ResponseEncoding="utf-8"  Debug="true" %>

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
<link rel="stylesheet" href="https://pro.fontawesome.com/releases/v5.10.0/css/all.css" integrity="sha384-AYmEC3Yw5cVb3ZcuHtOA93w35dYTsvhLPVnYs9eStHfGJvOvKxVfELGroGkvsg+p" crossorigin="anonymous"/>

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
  
       SqlStra1  = "SELECT * FROM CheckTable order by indx asc"
	

        Dim da As SqlDataAdapter = New SqlDataAdapter(SqlStra1, Conn)
        Dim DS As DataSet = New DataSet
        da.Fill(DS)    '-- 把SQL指令執行完成的結果，填入DataSet裡面。

        Dim DT As DataTable = DS.Tables(0)
        '=============  ADO.NET / DataSet ==(End)======

        '---每頁展示 5 筆資料
        Dim PageSize As Integer = 100

        '--SQL指令共撈到多少筆（列）資料。RecordCount資料總筆（列）數
        Dim RecordCount As Integer = 0
        RecordCount = DT.Rows.Count()

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
	</form>
   
	<div align="center" style="width: 990px;border:1px #000000 solid;margin:0px auto;">
	  <h3>表單驗證設定</h3>																				
      <table width="100%" border="0">
        <tbody>
          <tr>
            <td width="45%" align="center">&nbsp;</td>
            <td width="55%" align="center">&nbsp;</td>
          </tr>
        </tbody>
      </table>
      <br />
      <table width="95%" id="tbData" class="table2excel" border="0" align="center" cellpadding="1" cellspacing="1" >
        <tbody>
    <tr style="color: #fff;">
      <td height="30" align="center" bgcolor="#39A96B">操作</td>
      <td align="center" bgcolor="#39A96B">tableNU</td>
      <td align="center" bgcolor="#39A96B">tableName</td>
	  <td align="center" bgcolor="#39A96B">name</td>
	  <td align="center" bgcolor="#39A96B">value</td>
      <td align="center" bgcolor="#39A96B">chka</td>
      </tr>
			
		  <%

	While ((rowNo < PageSize) And (NowPageCount < RecordCount))
            haveRec = True
	%> 
			  <form name="form_<%=DT.Rows(NowPageCount).Item("indx")%>" id="form_<%=DT.Rows(NowPageCount).Item("indx")%>" method="post" action="test_uppost2.aspx"   onsubmit='return goch();'><input type="hidden" id="active" name="active" value="22">
				  <input type="hidden" id="indx" name="indx" value="<%=DT.Rows(NowPageCount).Item("indx")%>">
    <tr>
      <td height="44" align="center" <%if rowNo mod 2 = 0 then%>bgcolor="#9AEB65" <%else%> bgcolor="#D9F86B"<%end if%>>
		<button type="submit" class="btn btn-primary" style="font-size:10pt; padding:0px; width:60px"  >更新</button>
		</td>
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="#9AEB65" <%else%> bgcolor="#D9F86B"<%end if%>>
	  <%=DT.Rows(NowPageCount).Item("tableNU")%>
		</td>
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="#9AEB65" <%else%> bgcolor="#D9F86B"<%end if%>>
		<%=DT.Rows(NowPageCount).Item("tableName")%>
		</td>
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="#9AEB65" <%else%> bgcolor="#D9F86B"<%end if%>>
		<input type="text" style="width: 80%" id="name" name="name" value="<%=DT.Rows(NowPageCount).Item("namea")%>"/>
		</td>
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="#9AEB65" <%else%> bgcolor="#D9F86B"<%end if%>>
		<input type="text" style="width: 80%" id="value" name="value" value="<%=DT.Rows(NowPageCount).Item("valuea")%>"/>
		</td>
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="#9AEB65" <%else%> bgcolor="#D9F86B"<%end if%>>
		
		
		  <select name="chka" id="chka">
		    <option <%if DT.Rows(NowPageCount).Item("chka")="true"%>selected<%end if%> value="true">true</option>
		    <option <%if DT.Rows(NowPageCount).Item("chka")="false"%>selected<%end if%>  value="false">false</option>
		  </select>
      

		</td>
      </tr>
		</form>
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
