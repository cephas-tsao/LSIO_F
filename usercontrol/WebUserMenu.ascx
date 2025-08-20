<%@ Control Language="VB" AutoEventWireup="false" CodeFile="WebUserMenu.ascx.vb" Inherits="WebUserMenu" %>


<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>MENU</title>
<%--<link rel='stylesheet' href ="/../Setting/menu.css" type = 'text/css'/>--%>
<style type='text/css'>
/* common styling */
/* set up the overall width of the menu div, the font and the margins */
.menu {
font-family: arial, sans-serif; 
width:1000px; 
margin:0; 
margin:0px 0;
margin-top:0px;

}
/*margin-top 控制下拉選單的上邊界 可以讓選單是否再底圖 */
/* remove the bullets and set the margin and padding to zero for the unordered list */
.menu ul {
padding:0; 
margin:0;
list-style-type: none;
}
/* float the list so that the items are in a line and their position relative so that the drop down list will appear in the right place underneath each list item */
.menu ul li {
float:left; 
position:relative;
}
/* style the links to be 104px wide by 30px high with a top and right border 1px solid white. Set the background color and the font size. */
.menu ul li a, .menu ul li a:visited {
display:block; 
text-align:center; 
text-decoration:none; 
width:110px; 
height:30px; 
color:#5A1E00; 
//background:#99ffff;
background:#6699FF;
line-height:30px; 
font-size:14px;
font-weight: bold;
border:2px solid #fff;
border-width:2px 2px 0 0;
}
/* make the dropdown ul invisible */
.menu ul li ul {
display: none;
}
/* specific to non IE browsers */
/* set the background and foreground color of the main menu li on hover */
.menu ul li:hover a {
color:#fff; 
background:#1C86EE;
font-weight: bold;
}
/* make the sub menu ul visible and position it beneath the main menu list item */
.menu ul li:hover ul {
display:block; 
position:absolute; 
top:32px; 
left:0; 
width:112px;
text-align:left; 
font-weight: normal;
font-size:12px;
}
/* style the background and foreground color of the submenu links */
.menu ul li:hover ul li a {
display:block; 
background:#faeec7; 
color:#000;
text-align:left; 
font-weight: normal;
font-size:12px;
}
/* style the background and forground colors of the links on hover */
.menu ul li:hover ul li a:hover {
background:#dfc184; 
color:#000;
text-align:left; 
font-weight: normal;
font-size:12px;
}
</style>





<style type='text/css'>
@charset "utf-8";
/* CSS Document */

body{
padding: 0px;
}

/*^'^ Navigation Structure ^'^*/
.nav-container-outer{
background: #990000;
padding: 0px;
height: 74px;
background: url(/images/menu-<%=menu_style.text %>/nav-bg.jpg);
}
.float-left{
float: left;
}
.float-right{
float: right;
}
.nav-container .divider{
display:block;
font-size:1px;
border-width:0px;
border-style:solid;
}
.nav-container .divider-vert{
float:left;
width:0px;
display: none;
}
.nav-container .item-secondary-title{
display:block;
cursor:default;
white-space:nowrap;
}
.clear{
font-size:1px;
height:0px;
width:0px;
clear:left;
line-height:0px;
display:block;
float:none;
}
.nav-container{
position:relative;
zoom:1;
margin: 0 auto;
}
.nav-container a, .nav-container li{
float:left;
display:block;
white-space:nowrap;
}
.nav-container div a, .nav-container ul a, .nav-container ul li{
float:none;
}
.nav-container ul{
left:-10000px;
position:absolute;
}
.nav-container, .nav-container ul{
list-style:none;
padding:0px;
margin:0px;
}
.nav-container li a{
float:none
}
.nav-container li{
position:relative;
}
.nav-container ul{
z-index:10;
}
.nav-container ul ul{
z-index:20;
}
.nav-container ul ul ul{
z-index:30;
}
.nav-container ul ul ul ul{
z-index:40;
}
.nav-container ul ul ul ul ul{
z-index:50;
}
li:hover>ul{
left:auto;
}
#nav-container ul {
top:100%;
}
#nav-container ul li:hover>ul{
top:0px;
left:100%;
}

/*^'^ Primary Items ^'^*/	
#nav-container a{	
padding:7px 17px 7px 15px;
margin: 10px 0px 0px 0px;
color: <%=textcolor.Text %>;;
font-family: Trebuchet MS, Arial, sans-serif, Helvetica;
font-size:14px;
text-decoration:none;
font-weight: bold;
 /*background: url(/images/menu-<%=menu_style.text %>/item-primary-bg.gif);  */
background-repeat: no-repeat;
background-position: top;
}

#nav-container a:hover{
color: #CD3700;
 /*background: url(/images/menu-<%=menu_style.text %>/item-primary-bg.gif); */
background-repeat: no-repeat;
background-position: center;
}

/*^'^ Secondary Items Container ^'^*/	
#nav-container div, #nav-container ul{	
padding:10px 4px 10px 4px;
margin:0px 0px 0px 0px;
/*background:url(/images/menu-<%=menu_style.text %>/item-secondary-container-bg.jpg); */
/*background-attachment: fixed; */
background-repeat:no-repeat;
background-color:<%=color.Text %>;
border-bottom: 1px solid #CA6500;
}

/*^'^ Secondary Items Container ^'^*/
/* 第二層 */
#nav-container ul ul{	
padding:10px 4px 10px 4px;
margin:0px 0px 0px 0px;
/*background:url(/images/menu-<%=menu_style.text %>/item-secondary-container-bg.jpg); */
/*background-attachment: fixed; */
background-repeat:no-repeat;
background-color:<%=color2.Text %>;
border-bottom: 1px solid #CA6500;
}

/*^'^ Secondary Items ^'^ list底圖顏色*/	
#nav-container div a, #nav-container ul a{	
padding:3px 10px 3px 6px;
/*background-color: #FFFFFF; */
background-color:<%=color.Text %>;
/*background: url(/images/menu-<%=menu_style.text %>/item-secondary-bg.jpg); */
background-repeat: no-repeat;
background-position: auto auto; 
font-size:11px;
border-width:0px;
border-style:dotted;        /* 選擇框線 */
margin: 0px 0px 0px 0px;
}

/*^'^ Secondary Items ^'^ list底圖顏色*/
/* 第二層 */	
#nav-container ul ul a{	
padding:3px 10px 3px 6px;
/*background-color: #FFFFFF; */
background-color:<%=color2.Text %>;
/*background: url(/images/menu-<%=menu_style.text %>/item-secondary-bg.jpg); */
background-repeat: no-repeat;
background-position: auto auto; 
font-size:11px;
border-width:0px;
border-style:dotted;        /* 選擇框線 */
margin: 0px 0px 0px 0px;
}

/*^'^ Secondary Items Hover State ^'^*/	
#nav-container div a:hover, #nav-container ul a:hover{	
background-color: #FFFFFF;
/*background: url(/images/menu-<%=menu_style.text %>/item-secondary-bg.jpg); */
background-repeat: no-repeat;
color:<%=ontextcolor.Text %>;
}
/*^'^ Secondary Items Hover State ^'^*/	
#nav-container a:hover{	
background-color: #FFFFFF;
/*background: url(/images/menu-<%=menu_style.text %>/item-secondary-bg.jpg); */
background-repeat: no-repeat;
color:<%=ontextcolor.Text %>;
}

/*^'^ Secondary Item Titles ^'^*/	
#nav-container .item-secondary-title{	
cursor:default;
padding:4px 0px 3px 7px;
color: #6C3600;
font-family: Arial, Trebuchet MS, Arial, sans-serif, Helvetica;
font-size:11px;
/* background: url(images/menu-<%=menu_style.text %>/item-secondary-title-bg.jpg); */
background-repeat: no-repeat;
font-weight:bold;
}

/*^'^ Horizontal Dividers ^'^*/	
#nav-container .divider-horiz{	
border-top-width:1px;
margin:5px 5px;
border-color: #C16100;
}

/*^'^ Vertical Dividers ^'^*/	
#nav-container .divider-vert{	
border-left-width:1px;
height:15px;
margin:4px 2px 0px 2px;
border-color:#AAAAAA;
}
/*圖片中間加文字*/
.waicheng 
{
color: white; 
margin:5px auto; 
width:1024px; 
height:98px; 
line-height:120px; 
border:0px solid #000000; 
background:url(/images/menu-<%=menu_style.text %>/default-head2.jpg) no-repeat center; 
}
.itext
{
font-size:40px;
position:relative;
top:10px;
left:250px; 
color:Black;
}
.itext2
{
position:relative;
top:5px;
right:-500px; 
color:Black;
}
.img {vertical-align:middle;}
</style>


<!--[if lte IE 6]>
<style type='text/css'>
/* styling specific to Internet Explorer IE5.5 and IE6. Yet to see if IE7 handles li:hover */
/* Get rid of any default table style */
table {
border-collapse:collapse;
margin:0; 
padding:0;
}
/* ignore the link used by 'other browsers' */
.menu ul li a.hide, .menu ul li a:visited.hide {
display:none;
}
/* set the background and foreground color of the main menu link on hover */
.menu ul li a:hover {
color:#fff; 
background:#1C86EE;
font-weight: bold;
}
/* make the sub menu ul visible and position it beneath the main menu list item */
.menu ul li a:hover ul {
display:block; 
position:absolute; 
top:32px; 
left:0; 
width:105px;
font-weight: normal;
font-size:11px;
}
/* style the background and foreground color of the submenu links */
.menu ul li a:hover ul li a {
background:#faeec7; 
color:#000;
font-weight: normal;
font-size:11px;
}
/* style the background and forground colors of the links on hover */
.menu ul li a:hover ul li a:hover {
background:#dfc184; 
color:#000;
font-weight: normal;
font-size:11px;
}
</style>
<![endif]-->


</head>
<body>
<div class="waicheng" ><asp:Label ID="Label1" runat="server" Font-Bold="True" CssClass="itext" Font-Size="X-Large"></asp:Label>
    <asp:Label ID="usr_name" runat="server" CssClass="itext2" Font-Size="Medium"></asp:Label>
    <asp:Label ID="login_time" runat="server" CssClass="itext2" Font-Size="Medium"></asp:Label>
</div>
<asp:Literal ID="Literal1" runat="server"></asp:Literal>
    <asp:Label ID="menu_style" runat="server" Visible="False"></asp:Label>
    <asp:Label ID="color" runat="server" Visible="False"></asp:Label>
    <asp:Label ID="color2" runat="server" Visible="False"></asp:Label>
    <asp:Label ID="textcolor" runat="server" Visible="False"></asp:Label>
    <asp:Label ID="ontextcolor" runat="server" Visible="False"></asp:Label>
</body>
</html>