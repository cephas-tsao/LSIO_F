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
Dim Conn1 As SqlConnection = New SqlConnection
      Conn1.ConnectionString = WebConfigurationManager.ConnectionStrings("sql").ConnectionString  
%>

<!DOCTYPE html>
<html lang="en">

<head>

    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="description" content="">
    <meta name="author" content="">

    <title>食農知味-臺灣農民直銷站資訊網</title>

    <!-- Bootstrap Core CSS -->
    <link href="../vendor/bootstrap/css/bootstrap.min.css" rel="stylesheet">

    <!-- MetisMenu CSS -->
    <link href="../vendor/metisMenu/metisMenu.min.css" rel="stylesheet">

    <!-- Custom CSS -->
    <link href="../dist/css/sb-admin-2.css" rel="stylesheet">

    <!-- Custom Fonts -->
    <link href="../vendor/font-awesome/css/font-awesome.min.css" rel="stylesheet" type="text/css">

    <!-- HTML5 Shim and Respond.js IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
        <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
        <script src="https://oss.maxcdn.com/libs/respond.js/1.4.2/respond.min.js"></script>
    <![endif]-->

</head>

<body style="font-family:'微軟正黑體'">

    <div id="wrapper">

     
<!--#include file="main.aspx" -->
        <!-- Page Content -->
       
       <div id="page-wrapper">
         <div class="container-fluid">
                <div class="row">
                    <div class="col-lg-12">
                        <h1 class="page-header">(直銷站)資料管理</h1>
                    </div>
                </div> 
                      
              <div  onClick="" class="row"></div><!--
                <button type="button" class="btn btn-primary" style="font-size:10pt;"  onClick="window.location.href='5-1-web-add.aspx'">新增()</button>
                -->
                
         <form name="form" method="post" action="test_uppost2.aspx"   onsubmit='return goch();'> 
         
           <table width="50%" border="0" align="left" style="width:50%; padding:0px;" class="table table-striped table-bordered table-hover">
                  <tr>
                    <td>地區</td>
                    <td>直銷站名稱</td>
                    <td>新增</td>
                  </tr>
                  <tr>
         <td ><input type="text" id="d2" name="d2" value="" placeholder="請填北區、中區、東區、南區其中一個" style="width:100%;"></td>
        <td ><input type="text" id="d3" name="d3" value=""></td>
        <td><button type="submit" class="btn btn-primary" style="font-size:10pt; padding:0px; width:60px"  >新增</button></td>
                  </tr>
           </table>
           <input name="active" type="hidden" id="active" value="11" size="2" > 
         </form>
           <br><br>
           <%

 Dim pic2,seq1,seq2,seq3,seq4,seq5,seq6,seq7,seq8,seqtt
  Dim p As String = Request("p")
  Dim snum As String = Request("snum")
  Dim c As String = Request("c") 
  
	 Dim haveRec = False

        '-- p 就是「目前在第幾頁?」    
   Conn1.open()
   Dim SqlStra1 As String
  
       SqlStra1  = "SELECT * FROM accounting where a4='直銷站' order by indx desc"
	   
        Dim da As SqlDataAdapter = New SqlDataAdapter(SqlStra1, Conn1)
        Dim DS As DataSet = New DataSet
        da.Fill(DS)    '-- 把SQL指令執行完成的結果，填入DataSet裡面。

        Dim DT As DataTable = DS.Tables(0)
        '=============  ADO.NET / DataSet ==(End)======

        '---每頁展示 5 筆資料
        Dim PageSize As Integer = 200

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
		'
		%>
        
     
           <table width="100%" style="width:100%; padding:0px;" class="table table-striped table-bordered table-hover">
                  <thead>
                    <tr>               
					<th width="8%">操作</th>
                    <th width="8%">環境照片</th>
					<th width="9%">地區</th>
                    <th width="17%">直銷站名稱</th>
                    <th width="11%">地址</th>
                    <th width="11%">連絡電話</th>
                    <th width="11%">營業星期</th>
                    <th width="11%">營業時間</th>
                    <th width="11%">官網</th>
                    <th width="11%">FB</th>
                    </tr>
                  </thead>
    <%
	While ((rowNo < PageSize) And (NowPageCount < RecordCount))
            haveRec = True
	%>
                   
   <form name="form" method="post" action="test_uppost2.aspx"   onsubmit='return goch();'>
                     <tr>
   <input name="active" type="hidden" id="active" value="9" size="1" >
   <input name="indx" type="hidden" id="indx" value="<%=DT.Rows(NowPageCount).Item("indx")%>" size="10" >
					   
					   <td align="center" valign="middle">
<button type="submit" class="btn btn-primary" style="font-size:10pt; padding:0px; width:60px"  >更新</button>
                       </td>
        <td ><button type="button" class="btn btn-info" style="font-size:10pt; padding:0px; width:60px" onClick="location.href='7-2-web.aspx?a1=<%=DT.Rows(NowPageCount).Item("a1")%>&a3=<%=DT.Rows(NowPageCount).Item("d3")%>'" >照片管理</button></td>           
        <td ><input type="text" id="d2" name="d2" value="<%=DT.Rows(NowPageCount).Item("d2")%>"></td>
       
        <td ><input type="text" id="d3" name="d3" value="<%=DT.Rows(NowPageCount).Item("d3")%>"></td>
        <td ><input type="text" id="d5" name="d5" value="<%=DT.Rows(NowPageCount).Item("d5")%>"></td>
        <td ><input type="text" id="d6" name="d6" value="<%=DT.Rows(NowPageCount).Item("d6")%>"></td>
        <td ><input type="text" id="d7" name="d7" value="<%=DT.Rows(NowPageCount).Item("d7")%>"></td>
        <td ><input type="text" id="d8" name="d8" value="<%=DT.Rows(NowPageCount).Item("d8")%>"></td>
        <td ><input type="text" id="d10" name="d10" value="<%=DT.Rows(NowPageCount).Item("d10")%>"></td>
        <td ><input type="text" id="d11" name="d11" value="<%=DT.Rows(NowPageCount).Item("d11")%>"></td>
		
                       
                     </tr>                  
                     </form>
               <%
   'html = "<li class='on'><img width='165' height='100' src='adimages/"& DT.Rows(NowPageCount).Item("pic2")&"' /></li>"
            NowPageCount = NowPageCount + 1
           rowNo = rowNo + 1
		'Response.Write(html)
		   
    End While
		

	%>  
         
                   <tbody>

                   </tbody>
                 </table>
         </div>
                <!-- /.row -->
            </div>
            <!-- /.container-fluid -->
        </div>
        <!-- /#page-wrapper -->

    </div>
    <!-- /#wrapper -->

    <!-- jQuery -->
    <script src="../vendor/jquery/jquery.min.js"></script>

    <!-- Bootstrap Core JavaScript -->
    <script src="../vendor/bootstrap/js/bootstrap.min.js"></script>

    <!-- Metis Menu Plugin JavaScript -->
    <script src="../vendor/metisMenu/metisMenu.min.js"></script>

    <!-- Custom Theme JavaScript -->
    <script src="../dist/js/sb-admin-2.js"></script>
    
    <script src="js/jquery.blockUI.js" type="text/javascript" ></script>
    
    <script>
  
  function msgre_post(x){

var int = x ;
var a1 = $("#a_"+ x +"_1").text();
var a2 = $("#a_"+ x +"_2").text();
var a3 = $("#a_"+ x +"_3").text();
var a4 = $("#a_"+ x +"_4").text();
var a5 = $("#a_"+ x +"_5").text();
var a6 = $("#a_"+ x +"_6").text();
var a7 = $("#a_"+ x +"_7").text();
var a8 = $("#a_"+ x +"_8").text();
var a9 = $("#a_"+ x +"_9").text();

//alert(""+ int +","+ a1 +","+ a2 +","+ a3 +","+ a4 +","+ a5 +","+ a6 +","+ a7 +"");
//return false;

var tidname = "<%=tidname%>";
var dateuser ="<%=Format(Now(), "yyyyMMdd")%>"; 
 
  


var IDuser ="<%=ID%>"; 
var tindx="<%=tindx%>";
var active="up";
  
  
//alert(""+ tidname +","+ dateuser +","+ IDuser +","+ tindx +","+ active +""); 
//return false;
 
 
  
  $.blockUI({
	message:'<img src=loading.gif class=loading>記錄中.....',
	css:{border:'3px solid white',background:'#CCCCCC',padding:'10px',color:'#000000' }
	
	});

	$.ajax({ 
		type: "post",  			
                url: 'msg_post5.aspx',
                cache: true,
				data: {
					int:int,
					//ta:typea,
					tidname:tidname,
					dateuser:dateuser,
					IDuser:IDuser,
					a1:a1,
					a2:a2,
					a3:a3,
					a4:a4,
					a5:a5,
					a6:a6,
					a7:a7,
					a8:a8,
					a9:a9
				},
				success: function(data) { 
					$.unblockUI();
					//$("#qqa").html(data);
					//timego();            
				} ,
				error: function(data){
					alert('網路錯誤暫時無法存檔');	
				}
            }); 

  }



   </script>

</body>

</html>
