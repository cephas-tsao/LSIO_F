﻿<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
</script>
<%
    Session("usr_code") = ""
    Session("usr_name") = ""
    'Session("usr_dep_code") = ""
    Session("Dpname") = ""
    Session("dpnos") = ""
    Session("Updpno") = ""
    Session.Abandon()
    Response.Redirect("default.aspx")
%>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
    </div>
    </form>
</body>
</html>
