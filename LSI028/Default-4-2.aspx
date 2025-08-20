<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default-4-2.aspx.vb" Inherits="basic_LSI028_Default_4" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
<title><%=ConfigurationManager.AppSettings("Prj_name")%></title>

<script src="js/jquery.min.js"></script>
<script type="text/javascript" defer src="js/alphafilter.js"></script>


<script type="text/javascript" src="js/calendar/calendar.js"></script>
<script type="text/javascript" src="js/calendar/lang/calendar-en.js"></script>
<script type="text/javascript" src="js/calendar/calendar-setup.js"></script>
    
</head>

<body>
  <table width="95%" id="tbData" class="table2excel" border="1" align="center" cellpadding="1" cellspacing="1" >
        <tbody>
    <tr bgcolor="#ccc">
    <td align="center" bgcolor="">申報日期</td>
	<td align="center" bgcolor="">申報編號</td>	
	<td align="center" bgcolor="">申報單位</td>
	<td align="center" bgcolor="">季別</td>
        <asp:Literal ID="Literal1" runat="server"></asp:Literal>
   <%-- <td align="center" bgcolor="">1</td>
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
<td align="center" bgcolor="">77</td>--%>


	
    </tr>
			
	
    <tr>
		
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=Format(DT.Rows(NowPageCount).Item("date"), "yyyyMMdd") %></td>
		 <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("pc_code")%></td>
      <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("com_name")%></td>
 <td align="center" <%if rowNo mod 2 = 0 then%>bgcolor="" <%else%> bgcolor=""<%end if%>><%=DT.Rows(NowPageCount).Item("pc_season")%></td>
 
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
</body>
</html>
