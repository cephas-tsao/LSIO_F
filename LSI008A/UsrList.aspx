<%@ Page Language="VB" AutoEventWireup="false" CodeFile="UsrList.aspx.vb" Inherits="basic_LSI008A_UsrList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>資料檢視頁面</title>
    <script src="../../public/jquery-1.3.2.js" type="text/javascript"></script>
    <script src="../../public/jquery.autocomplete.js" type="text/javascript"></script>
    <link href="../../public/jquery.autocomplete.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript">
            $(function() {
            //選項說明: http://docs.jquery.com/Plugins/Autocomplete/autocomplete#url_or_dataoptions
            $("#TxtFind").autocomplete("../../public/AutoComplete.aspx?c=UsrList",
            {
                delay: 10,
                width: 120,
                minChars: 1, //至少輸入幾個字元才開始給提示?
                matchSubset: false,
                matchContains: false,
                cacheLength: 1,
                noCache: true, //黑暗版自訂參數，每次都重新連後端查詢(適用總資料筆數很多時)
                onItemSelect: findValue,
                onFindValue: findValue,
                formatItem: function(row) {
                    return "<div style='height:12px'><div style='float:left'>" + row[0] +
                            "</div>";
                    /*                
                    return "<div style='height:12px'><div style='float:left'>" + row[0] +
                    "</div><div style='float:right;padding-right:5px;'>" +
                    row[1] + "/" + row[2] + "</div></div>";
                    */
                },
                autoFill: false,
                mustMatch: false //是否允許輸入提示清單上沒有的值?
            });
            function findValue(li) {
                if (li == null) return alert("No match!");
                $("#TxtFind").val(li.extra[0]);
                //                $("#txtCName").val(li.extra[1]);
            }
        });
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div style="text-align: center"><asp:Label ID="Label1" runat="server" style="text-align: center" Font-Bold="True"></asp:Label></div>
    <div>
        <table width="100%" border ="1">
            <tr style="background-color:#d3d3d3">
                <td style="width: 70%" colspan="2">
                    <asp:Image ID="Image1" runat="server" ImageUrl="~/images/bs_search.gif" />搜尋<asp:DropDownList
                        ID="ddlField" runat="server" DataSourceID="SqlDS2" DataTextField="field_name"
                        DataValueField="sel_field">
                    </asp:DropDownList>
                    <asp:TextBox ID="TxtFind" runat="server"></asp:TextBox>
                    <asp:Button ID="Button1" runat="server" Text="查詢" /><br/>
                    <asp:Button ID="Button2" runat="server" Text="確定" />
                </td>
            </tr>
            <tr>
                <td colspan="2" bgcolor="#99CCFF">
                    <asp:CheckBox ID="CheckBox1" runat="server" AutoPostBack="True" Text="全選" />
                    <asp:CheckBoxList ID="CheckBoxList1" runat="server">
                    </asp:CheckBoxList>
                </td>
            </tr>
            <tr>
                <td style="width: 70%">
                    <asp:TextBox ID="MrdCode" runat="server" Visible="False"></asp:TextBox>
                    <asp:TextBox ID="txtSubWhere" runat="server" Visible="False"></asp:TextBox></td>
                <td style="width: 30%">
                    </td>
            </tr>
        </table>
        <asp:SqlDataSource ID="SqlDS2" runat="server"></asp:SqlDataSource>
    </div>
    </form>
</body>
</html>
