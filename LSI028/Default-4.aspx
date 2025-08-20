<%@ Page Language="VB" AutoEventWireup="false"  validateRequest="False" ContentType="text/html" ResponseEncoding="BIG5"  Debug="true" %>

<%
response.Buffer = true
session.Codepage =950
response.Charset = "BIG5"
%>
<%@ Import Namespace = "System.Data" %>
<%@ Import Namespace = "System.Data.SqlClient" %>
<%@ Import NameSpace = "System.IO"%>
<%@ Import NameSpace = "System.Web.Configuration" %>
<%@ Import NameSpace = "NPOI.HSSF.UserModel" %>
<%@ Import NameSpace = "NPOI.HPSF" %>
<%@ Import NameSpace = "NPOI.POIFS.FileSystem"  %>
<%
Dim Conn As SqlConnection = New SqlConnection
      Conn.ConnectionString = WebConfigurationManager.ConnectionStrings("AppSysConnectionString").ConnectionString
%>
	

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
<title><%=ConfigurationManager.AppSettings("Prj_name")%></title>

<script src="js/jquery.min.js"></script>
<script type="text/javascript" defer src="js/alphafilter.js"></script>


<script type="text/javascript" src="js/calendar/calendar.js"></script>
<script type="text/javascript" src="js/calendar/lang/calendar-en.js"></script>
<script type="text/javascript" src="js/calendar/calendar-setup.js"></script>
    
 <script src="js/jquery.table2excel.js"></script>
 
	
	<script language="javascript">
		
	//	$(document).ready(function() {
   //xls();
		//	history.back();
//});

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

         SqlStra1  = "SELECT * FROM self_chk order by date desc"


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
    
 <table width="95%" id="tbData" class="table2excel" border="1" align="center" cellpadding="1" cellspacing="1" >
        <tbody>
    <tr bgcolor="#ccc">
    <td align="center" style="background-color: #EEFAFF">申報日期</td>
	<td align="center" style="background-color: #EEFAFF">申報編號</td>	
	<td align="center" style="background-color: #EEFAFF">申報單位</td>
	<td align="center" style="background-color: #EEFAFF">季別</td>
    <td align="center" style="background-color: #EEFAFF">1</td>
<td align="center" style="background-color: #EEFAFF">2</td>
<td align="center" style="background-color: #EEFAFF">3</td>
<td align="center" style="background-color: #EEFAFF">4</td>
<td align="center" style="background-color: #EEFAFF">5</td>
<td align="center" style="background-color: #EEFAFF">6</td>
<td align="center" style="background-color: #EEFAFF">7</td>
<td align="center" style="background-color: #EEFAFF">8</td>
<td align="center" style="background-color: #EEFAFF">9</td>
<td align="center" style="background-color: #EEFAFF">10</td>
<td align="center" style="background-color: #EEFAFF">11</td>
<td align="center" style="background-color: #EEFAFF">12</td>
<td align="center" style="background-color: #EEFAFF">13</td>
<td align="center" style="background-color: #EEFAFF">14</td>
<td align="center" style="background-color: #EEFAFF">15</td>
<td align="center" style="background-color: #EEFAFF">16</td>
<td align="center" style="background-color: #EEFAFF">17</td>
<td align="center" style="background-color: #EEFAFF">18</td>
<td align="center" style="background-color: #EEFAFF">19</td>
<td align="center" style="background-color: #EEFAFF">20</td>
<td align="center" style="background-color: #EEFAFF">21</td>
<td align="center" style="background-color: #EEFAFF">22</td>
<td align="center" style="background-color: #EEFAFF">23</td>
<td align="center" style="background-color: #EEFAFF">24</td>
<td align="center" style="background-color: #EEFAFF">25</td>
<td align="center" style="background-color: #EEFAFF">26</td>
<td align="center" style="background-color: #EEFAFF">27</td>
<td align="center" style="background-color: #EEFAFF">28</td>
<td align="center" style="background-color: #EEFAFF">29</td>
<td align="center" style="background-color: #EEFAFF">30</td>
<td align="center" style="background-color: #EEFAFF">31</td>
<td align="center" style="background-color: #EEFAFF">32</td>
<td align="center" style="background-color: #EEFAFF">33</td>
<td align="center" style="background-color: #EEFAFF">34</td>
<td align="center" style="background-color: #EEFAFF">35</td>
<td align="center" style="background-color: #EEFAFF">36</td>
<td align="center" style="background-color: #EEFAFF">37</td>
<td align="center" style="background-color: #EEFAFF">38</td>
<td align="center" style="background-color: #EEFAFF">39</td>
<td align="center" style="background-color: #EEFAFF">40</td>
<td align="center" style="background-color: #EEFAFF">41</td>
<td align="center" style="background-color: #EEFAFF">42</td>
<td align="center" style="background-color: #EEFAFF">43</td>
<td align="center" style="background-color: #EEFAFF">44</td>
<td align="center" style="background-color: #EEFAFF">45</td>
<td align="center" style="background-color: #EEFAFF">46</td>
<td align="center" style="background-color: #EEFAFF">47</td>
<td align="center" style="background-color: #EEFAFF">48</td>
<td align="center" style="background-color: #EEFAFF">49</td>
<td align="center" style="background-color: #EEFAFF">50</td>
<td align="center" style="background-color: #EEFAFF">51</td>
<td align="center" style="background-color: #EEFAFF">52</td>
<td align="center" style="background-color: #EEFAFF">53</td>
<td align="center" style="background-color: #EEFAFF">54</td>
<td align="center" style="background-color: #EEFAFF">55</td>
<td align="center" style="background-color: #EEFAFF">56</td>
<td align="center" style="background-color: #EEFAFF">57</td>
<td align="center" style="background-color: #EEFAFF">58</td>
<td align="center" style="background-color: #EEFAFF">59</td>
<td align="center" style="background-color: #EEFAFF">60</td>
<td align="center" style="background-color: #EEFAFF">61</td>
<td align="center" style="background-color: #EEFAFF">62</td>
<td align="center" style="background-color: #EEFAFF">63</td>
<td align="center" style="background-color: #EEFAFF">64</td>
<td align="center" style="background-color: #EEFAFF">65</td>
<td align="center" style="background-color: #EEFAFF">66</td>
<td align="center" style="background-color: #EEFAFF">67</td>
<td align="center" style="background-color: #EEFAFF">68</td>
<td align="center" style="background-color: #EEFAFF">69</td>
<td align="center" style="background-color: #EEFAFF">70</td>
<td align="center" style="background-color: #EEFAFF">71</td>
<td align="center" style="background-color: #EEFAFF">72</td>
<td align="center" style="background-color: #EEFAFF">73</td>
<td align="center" style="background-color: #EEFAFF">74</td>
<td align="center" style="background-color: #EEFAFF">75</td>
<td align="center" style="background-color: #EEFAFF">76</td>
<td align="center" style="background-color: #EEFAFF">77</td>


	
    </tr>
			
		  <%

              While ((NowPageCount < RecordCount)) '(rowNo < PageSize) And
                  haveRec = True
	%> 
            <%
                Dim season As String = ""
                season = DT.Rows(NowPageCount).Item("pc_season")
	     
                If season = "1" Then season = "一" Else 
                If season = "2" Then season = "二" Else 
                If season = "3" Then season = "三" Else 
                If season = "4" Then season = "四"
                 %>
    <tr>
		
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=Format(DT.Rows(NowPageCount).Item("date"), "yyyyMMdd") %></td>
		 <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("pc_code")%></td>
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("com_name")%></td>
 <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=season%></td>
 
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("1")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("2")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("3")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("4")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("6")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("7")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("8")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("9")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("10")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("11")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("12")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("13")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("14")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("15")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("16")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("17")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("18")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("19")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("20")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("21")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("22")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("23")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("24")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("25")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("26")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("27")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("28")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("29")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("30")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("31")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("32")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("33")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("34")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("35")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("36")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("37")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("38")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("39")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("40")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("41")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("42")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("43")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("44")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("45")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("46")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("47")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("48")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("49")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("50")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("51")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("52")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("53")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("64")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("65")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("66")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("67")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("68")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("69")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("70")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("71")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("72")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("73")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("74")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("75")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("76")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("77")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("78")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("79")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("80")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("81")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("82")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("83")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("84")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("85")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("86")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("87")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("88")%></td>

    
    </tr>
			    <%

            NowPageCount = NowPageCount + 1
           rowNo = rowNo + 1
		   
    End While
	%> 
					
   
  </tbody>
</table>

	
																						

	
	<%

Conn.close() 
	   %>
</body>
</html>
		<%
		    'Response.Clear()
		    Response.HeaderEncoding = System.Text.Encoding.GetEncoding("utf-8")
		    'Response.Write("<meta http-equiv=Content-Type content=text/html;charset=BIG5>")
		    Response.AddHeader("Content-Disposition", "attachment;filename=Export.xls")
		    Response.ContentType = "application/ms-excel"
		   response.Write ("<script>history.back();</script>")
		    'history.back();		   
		    'Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
		    'Response.AddHeader("Content-Disposition", string.Format("attachment; filename={0}", saveAsFileName));
		    'workbook.Write(Response.OutputStream);
		    Response.Flush()
		    Response.Close()
		    
		    

		   %>
