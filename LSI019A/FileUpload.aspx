<%@ Page Language="VB" AutoEventWireup="false" CodeFile="FileUpload.aspx.vb" Inherits="basic_LSI019A_FileUpload" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>匯入Excel</title>
</head>
<body>
<form id="form1" runat="server">
<div align=center>
  <table style="background-color:#EDEDED" border="1" Width="580px">
    <tr style="text-align:center;background-color:#4F94CD">
      <td><b><font color="#FFFFFF" size="4">上傳Excel檔案</font></b></td>
    </tr>
    <tr style="text-align:left;">
      <td>請選擇欲上傳Excel檔：<br />
        <asp:FileUpload ID="FileUpload1" runat="server" Width="500px" />
        <asp:Button ID="Button1" runat="server" Text="確定" Width="60px" />
      </td>
    </tr>
    <tr style="text-align:center;background-color:#4F94CD">
      <td><b><font color="#FFFFFF" size="4">注意事項</font></b></td>
    </tr>
    <tr style="text-align:left;">
      <td>
        <p>1.匯入檔案格式<font color="red">僅限XLS</font><br />
        2.匯入檔案範本請"<a href="../../FileUpload/LSI019A_sample.xls">按此下載</a>"<br />
        3.匯入完成後請<font color="red">重新查詢資料</font><br />
        4.匯入資料會<font color="red">將原有資料清空</font>，請注意</p>   
      </td>
    </tr>
  </table>
</div>
</form>
</body>
</html>
