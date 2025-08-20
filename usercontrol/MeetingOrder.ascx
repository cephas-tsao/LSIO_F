<%@ Control Language="VB" AutoEventWireup="false" CodeFile="MeetingOrder.ascx.vb" Inherits="usercontrol_MeetingOrder" %>
<asp:GridView ID="GridView1" runat="server" BackColor="White" Width="100%" DataSourceID="SqlDS1" 
    BorderColor="#336666" BorderStyle="Double" BorderWidth="3px" 
    CellPadding="4" AutoGenerateColumns="False" EmptyDataText="無資料" 
    Font-Size="Small">
    <Columns>
        <asp:BoundField DataField="mrd_code" ItemStyle-BackColor="#d6dff7" />
        <asp:BoundField DataField="mrd_time_s" ItemStyle-BackColor="#d6dff7" />
        <asp:BoundField DataField="mrd_time_e" ItemStyle-BackColor="#d6dff7" />
        <asp:BoundField DataField="mrd_date" ItemStyle-BackColor="#d6dff7" />
        <asp:BoundField DataField="usr_name" ItemStyle-BackColor="#d6dff7" />
        <asp:BoundField DataField="room_name" ItemStyle-BackColor="#d6dff7" />
        <asp:BoundField DataField="day" ItemStyle-BackColor="#d6dff7" />
        <asp:BoundField DataField="mrd_name" ItemStyle-BackColor="#d6dff7" />
        <asp:BoundField DataField="mrd_host" ItemStyle-BackColor="#d6dff7" />
        <asp:BoundField HeaderText="日期" ItemStyle-BackColor="#d6dff7" >
            <ItemStyle  Width="16%" />
        </asp:BoundField>
        <asp:BoundField HeaderText="租借情形" ItemStyle-BackColor="#d6dff7" />
        <asp:BoundField HeaderText="" ItemStyle-BackColor="#d6dff7">
            <ItemStyle  Width="5%" />
        </asp:BoundField>
       <asp:TemplateField HeaderText="" ItemStyle-BackColor="#d6dff7">
            <ItemTemplate>
                <asp:ImageButton ID="IB_DEL" runat="server" CausesValidation="false" CommandName="delete"
                    ImageUrl="~/images/icon/del.png" OnClientClick="return confirm('確定要註銷資料?');"
                    ToolTip="註銷" ItemStyle-BackColor="#d6dff7" />
            </ItemTemplate>
            <ItemStyle HorizontalAlign="Center" Width="5%" />
        </asp:TemplateField>
    </Columns>
    <RowStyle BackColor="White" ForeColor="#333333" />
    <FooterStyle BackColor="White" ForeColor="#333333" />
    <PagerStyle BackColor="#336666" ForeColor="White" HorizontalAlign="Center" />
    <SelectedRowStyle BackColor="#339966" Font-Bold="True" ForeColor="White" />
    <HeaderStyle BackColor="#336666" Font-Bold="True" ForeColor="White" />
</asp:GridView>


<asp:TextBox ID="txt_yyyymm" runat="server" Visible="False" />
<asp:TextBox ID="txt_room_code" runat="server" Visible="False" />
<asp:SqlDataSource ID="SqlDS1" runat="server" />
