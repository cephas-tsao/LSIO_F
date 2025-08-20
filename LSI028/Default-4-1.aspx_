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
    
 <table width="95%" id="tbData" class="table2excel" border="0" align="center" cellpadding="1" cellspacing="1" >
        <tbody>
    <tr>
    <td align="center" bgcolor="">申報日期</td>
	<td align="center" bgcolor="">申報編號</td>	
	<td align="center" bgcolor="">申報單位</td>
	
    <td align="center" bgcolor="">1</td>
<td align="center" bgcolor="">2</td>
<td align="center" bgcolor="">3</td>
<td align="center" bgcolor="">4</td>
<td align="center" bgcolor="">5</td>
<td align="center" bgcolor="">6</td>
<td align="center" bgcolor="">7</td>
<td align="center" bgcolor="">8</td>
<td align="center" bgcolor="">9</td>
<td align="center" bgcolor="">10</td>
<td align="center" bgcolor="">11</td>
<td align="center" bgcolor="">12</td>
<td align="center" bgcolor="">13</td>
<td align="center" bgcolor="">14</td>
<td align="center" bgcolor="">15</td>
<td align="center" bgcolor="">16</td>
<td align="center" bgcolor="">17</td>
<td align="center" bgcolor="">18</td>
<td align="center" bgcolor="">19</td>
<td align="center" bgcolor="">20</td>
<td align="center" bgcolor="">21</td>
<td align="center" bgcolor="">22</td>
<td align="center" bgcolor="">23</td>
<td align="center" bgcolor="">24</td>
<td align="center" bgcolor="">25</td>
<td align="center" bgcolor="">26</td>
<td align="center" bgcolor="">27</td>
<td align="center" bgcolor="">28</td>
<td align="center" bgcolor="">29</td>
<td align="center" bgcolor="">30</td>
<td align="center" bgcolor="">31</td>
<td align="center" bgcolor="">32</td>
<td align="center" bgcolor="">33</td>
<td align="center" bgcolor="">34</td>
<td align="center" bgcolor="">35</td>
<td align="center" bgcolor="">36</td>
<td align="center" bgcolor="">37</td>
<td align="center" bgcolor="">38</td>
<td align="center" bgcolor="">39</td>
<td align="center" bgcolor="">40</td>
<td align="center" bgcolor="">41</td>
<td align="center" bgcolor="">42</td>
<td align="center" bgcolor="">43</td>
<td align="center" bgcolor="">44</td>
<td align="center" bgcolor="">45</td>
<td align="center" bgcolor="">46</td>
<td align="center" bgcolor="">47</td>
<td align="center" bgcolor="">48</td>
<td align="center" bgcolor="">49</td>
<td align="center" bgcolor="">50</td>
<td align="center" bgcolor="">51</td>
<td align="center" bgcolor="">52</td>
<td align="center" bgcolor="">53</td>
<td align="center" bgcolor="">54</td>
<td align="center" bgcolor="">55</td>
<td align="center" bgcolor="">56</td>
<td align="center" bgcolor="">57</td>
<td align="center" bgcolor="">58</td>
<td align="center" bgcolor="">59</td>
<td align="center" bgcolor="">60</td>
<td align="center" bgcolor="">61</td>
<td align="center" bgcolor="">62</td>
<td align="center" bgcolor="">63</td>
<td align="center" bgcolor="">64</td>
<td align="center" bgcolor="">65</td>
<td align="center" bgcolor="">66</td>
<td align="center" bgcolor="">67</td>
<td align="center" bgcolor="">68</td>
<td align="center" bgcolor="">69</td>
<td align="center" bgcolor="">70</td>
<td align="center" bgcolor="">71</td>
<td align="center" bgcolor="">72</td>
<td align="center" bgcolor="">73</td>
<td align="center" bgcolor="">74</td>
<td align="center" bgcolor="">75</td>
<td align="center" bgcolor="">76</td>
<td align="center" bgcolor="">77</td>
<td align="center" bgcolor="">78</td>
<td align="center" bgcolor="">79</td>
<td align="center" bgcolor="">80</td>
<td align="center" bgcolor="">81</td>
<td align="center" bgcolor="">82</td>
<td align="center" bgcolor="">83</td>
<td align="center" bgcolor="">84</td>
<td align="center" bgcolor="">85</td>
<td align="center" bgcolor="">86</td>
<td align="center" bgcolor="">87</td>
<td align="center" bgcolor="">88</td>

	
    </tr>
			
		  <%

	While ((rowNo < PageSize) And (NowPageCount < RecordCount))
            haveRec = True
	%> 
    <tr>
		
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=Format(DT.Rows(NowPageCount).Item("date"), "yyyyMMdd") %></td>
		 <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("pc_code")%></td>
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("com_name")%></td>
     
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("1")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("2")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("3")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("4")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("5")%></td>
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
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("54")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("55")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("56")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("57")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("58")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("59")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("60")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("61")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("62")%></td>
<td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("63")%></td>
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
	Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")
        Response.AddHeader("content-disposition", "attachment;filename=Export.ods")
       Response.ContentType = "application/vnd.xls"	
		   response.Write ("<script>history.back();</script>")
		   'history.back();
		  
		   %>
