<%@ Page Language="VB" AutoEventWireup="false" CodeFile="edit.aspx.vb" Inherits="basic_LSI002A_edit" aspcompat="true" MaintainScrollPositionOnPostback="true" %>
<%@ Register src="../../usercontrol/WebUserMenu.ascx" tagname="WebUserMenu" tagprefix="uc1" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title><%=ConfigurationManager.AppSettings("Prj_name")%></title>
    <script language="javascript" id="js1" src="../../public/checkjs.js"></script> 
    <link rel="stylesheet" href="https://pro.fontawesome.com/releases/v5.10.0/css/all.css" integrity="sha384-AYmEC3Yw5cVb3ZcuHtOA93w35dYTsvhLPVnYs9eStHfGJvOvKxVfELGroGkvsg+p" crossorigin="anonymous"/><link rel='stylesheet' href ="/../Setting/style<%=Get_SysPara("css_sys", "01")%>.css" type = 'text/css'/>
    <style type="text/css">
        .style1
        {
            width: 13%;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <table style="position:relative; margin-left:auto; margin-right:auto; width:990px; "  border ="0">
        <tr>
           <td width="100%">
             <div>
                <uc1:WebUserMenu ID="WebUserMenu" runat="server" />
             </div>
           </td>
        </tr>
        <tr>        
           <td>
    <div>
        <table width="100%" border="1" id="customers" cellspacing='0.1'>
            <tr>
                <td class="fail" colspan="6" style="text-align: center">
                    <asp:Label ID="Label1" runat="server" Font-Bold="True" Style="text-align: center"></asp:Label></td>
            </tr>
            <tr>
                <td class="style1">
                    受理進度</td>
                <td class="style2" colspan="5">
                    <asp:RadioButtonList ID="com_schedule" runat="server">
                        <asp:ListItem Value="01">實施檢查中</asp:ListItem>
                        <asp:ListItem Value="02">發函事業單位備妥資料至本局受檢</asp:ListItem>
                        <asp:ListItem Value="03">第2次發函事業單位備妥資料至本局受檢</asp:ListItem>
                        <asp:ListItem Value="04">已回復結案</asp:ListItem>
                    </asp:RadioButtonList>
            </tr>
            <tr>
                <td class="style1">
                    受理日期</td>
                <td class="style2" colspan="5">
                    <asp:TextBox ID="ins_date" runat="server" MaxLength="12" Width="120px" 
                        Enabled="False"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="style1">
                    受理時間</td>
                <td class="style2" colspan="5">
                    <asp:TextBox ID="cmp_date" runat="server" MaxLength="12" Width="120px" 
                        Enabled="False"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="style1">
                    案件編號<asp:Image ID="Image2" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2" colspan="5">
                    <asp:TextBox ID="txt_cmp_code" runat="server" MaxLength="12" Width="120px" 
                        Enabled="False"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="style1">
                    是否為代他人申訴，如是必須上傳委託函</td>
                <td class="style2" colspan="5">
                    <asp:RadioButton ID="rb_entrustY" runat="server" GroupName="entrust" Text="是"/>
                    <asp:RadioButton ID="rb_entrustN" runat="server" GroupName="entrust" Text="否"/>
                </td>
            </tr>            <tr>
                <td class="style1">
                    姓名</td>
                <td class="style2" colspan="5">
                    <asp:TextBox ID="txt_cmp_name" runat="server" MaxLength="16" Width="200px"></asp:TextBox>
                    <font color=red>(申訴類別為具名時必填)</font></td>
            </tr>
            <tr>
                <td class="style1">
                    性別</td>
                <td class="style2" colspan="5">
                <asp:RadioButton ID="msex" runat="server" text="男" value="m" groupname="csex"/>
                <asp:RadioButton ID="fsex" runat="server" text="女" value="f" groupname="csex"/><font color=red>(申訴類別為具名時必填)</font>
                </td>
            </tr>
           <tr>
                <td class="style1">
                    年齡</td>
                <td class="style2" colspan="5">
                <asp:TextBox ID="txt_cmp_age" runat="server" maxlength="3"/>
                <font color=red>(申訴類別為具名時必填)</font>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    電子信箱E-Mail</td>
                <td class="style2" colspan="5">
                    <asp:TextBox ID="txt_cmp_email" runat="server" MaxLength="100" Width="500px"></asp:TextBox>
                    <asp:Button ID="Bt_send" runat="server" Text="補發" Visible="False" />
                    <font color=red>(申訴類別為具名時必填)</font></td>
            </tr>
            <tr>
                <td class="style1">
                    聯絡電話</td>
                <td class="style2" colspan="5">
                    <asp:TextBox ID="txt_cmp_tel1" runat="server" MaxLength="16" Width="160px"></asp:TextBox>
                    <font color=red>(申訴類別為具名時必填)</font></td>
            </tr>
            <tr>
                <td class="style1">
                    申訴人部門</td>
                <td class="style2" colspan="5">
                    <asp:TextBox ID="txt_cmp_depart" runat="server" MaxLength="25" Width="200px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    申訴人職稱</td>
                <td class="style2" colspan="5">
                    <asp:TextBox ID="txt_cmp_job" runat="server" MaxLength="25" Width="200px"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="style1">
                    身份</td>
                <td class="style2" colspan="5">
                    <asp:RadioButton ID="rblist" runat="server" text="事業單位所屬勞工(目前尚在職者)" value="1" groupname="hide1" TabIndex="4"/><br>
                    <font color=red>*</font>到職日期：
                    <asp:TextBox ID="cmp_sdate" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                    <asp:ImageButton ID="img_date2" runat="server" 
                        ImageUrl="~/images/calendr.gif" />
                    <br>
                    <font color=red>*</font>擔任職務及工作內容：
                    <asp:TextBox ID="jobcontent" runat="server" maxlength="20" Width="200px" TabIndex="8" />
                    <br />
                    <asp:RadioButton ID="rblist2" runat="server" value="2" Text="之前曾經是事業單位勞工" groupname="hide1" TabIndex="9"/>，離職日期：
                    <asp:TextBox ID="cmp_edate" runat="server" MaxLength="10" Width="100px"></asp:TextBox>
                    <asp:ImageButton ID="img_date3" runat="server" 
                        ImageUrl="~/images/calendr.gif" />
                                          
                    <br />
                    <asp:RadioButton ID="rblist3" runat="server" text="民眾(非屬前二項者)" value="3" groupname="hide1" TabIndex="13"/>
                    <font color=red></font></td>
            </tr>
            <!--<tr>
                <td class="style1">
                    身份證字號</td>
                <td class="style2" colspan="5">
                                <asp:TextBox ID="txt_cmp_idno" runat="server" MaxLength="10" 
                        Width="100px"></asp:TextBox>
                    <font color=red>(申訴類別為具名時必填)</font></td>
            </tr>-->
            <tr>
                <td class="style1">
                    被申訴事業單位名稱<asp:Image ID="Image3" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2" colspan="5">
                                <asp:TextBox ID="txt_com_name" runat="server" MaxLength="100" 
                        Width="500px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    統一編號</td>
                <td class="style2" colspan="5">
                                <asp:TextBox ID="txt_com_idno" runat="server" MaxLength="10" 
                        Width="100px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    地址<asp:Image ID="Image4" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2" colspan="5">
                    <asp:DropDownList ID="ddl_att_area" runat="server" AutoPostBack="True">
                    </asp:DropDownList>
                    <asp:DropDownList ID="ddl_att_road" runat="server">
                    </asp:DropDownList>
                                <asp:TextBox ID="txt_com_add" runat="server" MaxLength="100" 
                        Width="500px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    事業單位行業別</td>
                <td class="style2" colspan="5">
                                <asp:TextBox ID="com_type" runat="server" MaxLength="100" 
                        Width="500px"></asp:TextBox>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    勞工勞務給付地<asp:Image ID="Image5" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2" colspan="5">
                    <asp:DropDownList ID="ddl_att_area1" runat="server" AutoPostBack="True">
                    </asp:DropDownList>
                    <asp:DropDownList ID="ddl_att_road1" runat="server">
                    </asp:DropDownList>
                                <asp:TextBox ID="txt_com_add1" runat="server" MaxLength="100" 
                        Width="500px"></asp:TextBox>
                                </td>
            </tr>
            
            <tr>
                <td class="style1">
                    申訴類別<asp:Image ID="Image6" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2" colspan="5">
                    <asp:DropDownList ID="ddl_cmp_type" runat="server">
                    </asp:DropDownList>
                                </td>
            </tr>
             <tr>
                <td class="style1">
                    勞動契約</td>
                <td class="style2" colspan="5">
              <asp:CheckBoxList ID="com_Contract" runat="server" RepeatColumns="2" 
                        Height="16px">
                    </asp:CheckBoxList>
                    <asp:CheckBox ID="com_Contract_chk" runat="server" TabIndex="30"/>
                    其他：<asp:TextBox ID="com_Contract_other" runat="server" maxlength="30" Width="350px" TabIndex="33" />
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    工資</td>
                <td class="style2" colspan="5">
                <asp:CheckBoxList ID="com_paymon" runat="server" RepeatColumns="2" 
                        Height="16px">
                    </asp:CheckBoxList>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    工作時間、休息、休假</td>
                <td class="style2" colspan="5">
              <asp:CheckBoxList ID="com_relax" runat="server" RepeatColumns="2" 
                        Height="16px">
                    </asp:CheckBoxList>
                    <asp:CheckBox ID="com_relax_chk" runat="server" TabIndex="30"/>
                    其他：<asp:TextBox ID="com_relax_other" runat="server" maxlength="30" Width="350px" TabIndex="33" />
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    童工、女工</td>
                <td class="style2" colspan="5">
                <asp:CheckBoxList ID="com_child" runat="server" RepeatColumns="2" 
                        Height="16px">
                    </asp:CheckBoxList>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    保障與福利</td>
                <td class="style2" colspan="5">
              <asp:CheckBoxList ID="com_Welfar" runat="server" RepeatColumns="2" 
                        Height="16px">
                    </asp:CheckBoxList>
                    <asp:CheckBox ID="com_Welfare_chk" runat="server" TabIndex="30"/>
                    其他：<asp:TextBox ID="com_Welfare_other" runat="server" maxlength="30" Width="350px"/>
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    本爭議事項是否曾向其他勞工主管機關申訴<asp:Image ID="Image7" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2" colspan="5">
                    <asp:RadioButton ID="DN" runat="server" GroupName="disputetype" Text="否" />
                <asp:RadioButton ID="DY" runat="server" GroupName="disputetype" Text="是"/>，請註明日期及機關
                <asp:TextBox ID="com_disputecon" runat="server" maxlength="20" Width="295px" />
                                </td>
            </tr>
            <tr>
                <td class="style1">
                    本爭議事項是否曾申請勞資爭議調解<asp:Image ID="Image8" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2" colspan="5">
                    <asp:RadioButton ID="MN" runat="server" GroupName="Mediationtype" Text="否" />
                <asp:RadioButton ID="MYY" runat="server" GroupName="Mediationtype" Text="是" />，請註明日期及機關
                <asp:TextBox ID="com_Mediationcon" runat="server" maxlength="20" Width="295px"/>
                                </td>
            </tr>
            <!--<tr>
                <td class="style1">
                    個資限制</td>
                <td class="style2" colspan="5">
                    <asp:DropDownList ID="ddl_is_agree" runat="server">
                    </asp:DropDownList>是否同意本處繼續使用資料
                    <font color=red>(申訴類別為具名時必填)</font></td>
            </tr>-->
            <tr>
                <td class="style1">
                    陳情事項<asp:Image ID="Image1" runat="server" ImageUrl="~/images/tick.png" /></td>
                <td class="style2" colspan="5">
                    <asp:TextBox ID="txt_cmp_memo" runat="server" Width="624px" Height="150px" 
                        TextMode="MultiLine"></asp:TextBox></td>
            </tr> 
            <tr>
                <td class="style1">
                    為利於協助追償受損權益，是否願意將您的姓名提供給事業單位</td>
                <td class="style2" colspan="5">
                <asp:RadioButton ID="SY" runat="server" GroupName="sectype" Text="我願意將姓名提供給事業單位，以利檢查時協助追償受損權益" TabIndex="42"/><br>
                <asp:RadioButton ID="SN" runat="server" GroupName="sectype" Text="目前仍在職，請將我的身分保密" TabIndex="43"/>
                                </td>
            </tr>
             <tr>
                <td class="style1">
                    亂數密碼</td>
                <td class="style2" colspan="5">
                    <asp:TextBox ID="cmp_pass" runat="server" MaxLength="12" Width="120px" 
                        Enabled="False"></asp:TextBox></td>
            </tr>             
            <tr >
                <td class="style1">
                    附檔</td>
                <td class="style2">
                    <asp:Label ID="lab_cmp_file1" runat="server" Text="附檔1" ></asp:Label>
                                </td>
                <td class="style2">
                    <asp:Label ID="lab_cmp_file2" runat="server" Text="附檔2"></asp:Label>
                                </td>
                <td class="style2">
                    <asp:Label ID="lab_cmp_file3" runat="server" Text="附檔3"></asp:Label>
                                </td>
                <td class="style2">
                    <asp:Label ID="lab_cmp_file4" runat="server" Text="附檔4"></asp:Label>
                                </td>
                <td class="style2">
                    <asp:Label ID="lab_cmp_file5" runat="server" Text="附檔5"></asp:Label>
                                </td>
            </tr>              
        </table>
    </div>
            <asp:Button ID="Button1" runat="server" Text="存檔" />
            <asp:Button ID="Bt_BackUp" runat="server" Text="離開" />
            <asp:TextBox ID="txtPgmType" runat="server" Visible="False" />
            <asp:TextBox ID="txtPrimary" runat="server" Visible="False" />
            <asp:TextBox ID="txtSubWhere" runat="server" Visible="False" />
            <asp:TextBox ID="txtOrder" runat="server" Visible="False" />
            <asp:TextBox ID="txtFileName1" runat="server" Visible="False" />
            <asp:TextBox ID="txtFileName2" runat="server" Visible="False" />
            <asp:TextBox ID="txtFileName3" runat="server" Visible="False" />
            <asp:TextBox ID="txtFileName4" runat="server" Visible="False" />
            <asp:TextBox ID="txtFileName5" runat="server" Visible="False" />
            <asp:TextBox ID="cmp_idclass" runat="server" Visible="False" />
            <asp:TextBox ID="com_disputetype" runat="server" Visible="False" />
            <asp:TextBox ID="com_Mediationtype" runat="server" Visible="False" />
            <asp:TextBox ID="txtVerCode" runat="server" Visible="False" />
            <asp:TextBox ID="com_sectype" runat="server" Visible="False" />
            <asp:TextBox ID="com_sex" runat="server" Visible="False" />
            <asp:TextBox ID="txt_cmp_entrust" runat="server" Visible="False" /> <!--是否為代他人申訴RadioButton-->
            <br />
            <asp:Label ID="l_ErrMsg" runat="server" Font-Bold="True" ForeColor="Red" />
            </td>
        </tr>
      </table>    
    </form>
</body>
</html>
