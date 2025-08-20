<%@ Page Language="VB" AutoEventWireup="false" CodeFile="RecordDetail.aspx.vb" Inherits="public_RecordDetail" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>詳細資料</title>
    <link rel='stylesheet' href ="/../Setting/style<%=Get_SysPara("css_sys", "01")%>.css" type = 'text/css'/>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table width=100%  border="1" style="border-style: groove" cellspacing='0.1'>
            <tr align="center">
                <td style="width: 15%" class="fail">欄位名稱</td>
                <td style="width: 15%" class="fail">欄位代碼</td>
                <td style="width: 35%" class="fail">修改前資料</td>
                <td style="width: 35%" class="fail">修改後資料</td>
            </tr>
        <asp:TextBox ID="ID" runat="server" Visible="False"></asp:TextBox>
    </div>
    </form>
            <tr align="center">
                <td class="style1" colspan="4">
                    <asp:Table ID="Table1" runat="server" Width="100%" />
                </td>
            </tr>
        </body>
</html>
