<%@ Control Language="VB" AutoEventWireup="false" CodeFile="Menu.ascx.vb" Inherits="usercontrol_Menu" %>
<script type="text/javascript">
<!--
function MM_swapImgRestore() { //v3.0
    var i, x, a = document.MM_sr; for (i = 0; a && i < a.length && (x = a[i]) && x.oSrc; i++) x.src = x.oSrc;
}

function MM_preloadImages() { //v3.0
    var d = document; if (d.images) {
        if (!d.MM_p) d.MM_p = new Array();
        var i, j = d.MM_p.length, a = MM_preloadImages.arguments; for (i = 0; i < a.length; i++)
            if (a[i].indexOf("#") != 0) { d.MM_p[j] = new Image; d.MM_p[j++].src = a[i]; } 
    }
}

function MM_findObj(n, d) { //v4.01
    var p, i, x; if (!d) d = document; if ((p = n.indexOf("?")) > 0 && parent.frames.length) {
        d = parent.frames[n.substring(p + 1)].document; n = n.substring(0, p);
    }
    if (!(x = d[n]) && d.all) x = d.all[n]; for (i = 0; !x && i < d.forms.length; i++) x = d.forms[i][n];
    for (i = 0; !x && d.layers && i < d.layers.length; i++) x = MM_findObj(n, d.layers[i].document);
    if (!x && d.getElementById) x = d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
    var i, j = 0, x, a = MM_swapImage.arguments; document.MM_sr = new Array; for (i = 0; i < (a.length - 2); i += 3)
        if ((x = MM_findObj(a[i])) != null) { document.MM_sr[j++] = x; if (!x.oSrc) x.oSrc = x.src; x.src = a[i + 2]; }
}
//-->
</script>
<table border="0" cellspacing="0" cellpadding="0" summary="排版表格">
  <tr>
    <td><img src="../images/title-1.gif" width="294" height="72" alt="申辦服務項目圖示"></td>
  </tr>
  <tr>
    <td><a href="index-b01.asp" title="連至線上復工申請" onmouseout="MM_swapImgRestore()"  onmouseover="MM_swapImage('Image16','','../images/title-2.gif',1)" onblur="MM_swapImgRestore()"  onfocus="MM_swapImage('Image16','','../images/title-2.gif',1)"><img src="../images/title-2.gif" name="Image16" width="294" height="25" border="0" id="Image16" alt="連至線上復工申請"/></a></td>
  </tr>
  <tr>
    <td><img src="../images/title-line.gif" width="294" height="3" alt="*"/></td>
  </tr>
  <tr>
    <td><a href="index-menu.asp" title="連至危評審查" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image17','','../images/title-3.gif',1)" onblur="MM_swapImgRestore()" onfocus="MM_swapImage('Image17','','../images/title-3.gif',1)"><img src="../images/title-3.gif" name="Image17" width="294" height="25" border="0" id="Image17" alt="連至危評審查"/></a></td>
  </tr>
  <tr>
    <td><img src="../images/title-line.gif" width="294" height="3" alt="*"/></td>
  </tr>
  <tr>
    <td><a href="index-h050.asp" title="勞動安全電子報訂閱" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image18','','../images/title-4.gif',0)" onblur="MM_swapImgRestore()" onfocus="MM_swapImage('Image18','','../images/title-4.gif',0)"><img src="../images/title-4.gif" name="Image18" width="294" height="25" border="0" id="Image18" alt="勞動安全電子報訂閱"/></a></td>
  </tr>
  <tr>
    <td><img src="../images/title-line.gif" width="294" height="3" alt="*"/></td>
  </tr>
  <tr>
    <td><a href="enewindex.asp" title="勞動安全電子報" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image23','','../images/title-9.gif',0)" onblur="MM_swapImgRestore()" onfocus="MM_swapImage('Image23','','../images/title-9.gif',0)"><img src="../images/title-9.gif" name="Image23" width="294" height="25" border="0" id="Image23" alt="勞動安全電子報"/></a></td>
  </tr>
  <tr>
    <td><img src="../images/title-line.gif" width="294" height="3" alt="*"/></td>
  </tr>
  <tr>
    <td><a href="index-f050.asp" title="裝修工程及危險性作業通報" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image19','','../images/title-5.gif',1)" onblur="MM_swapImgRestore()" onfocus="MM_swapImage('Image19','','../images/title-5.gif',1)"><img src="../images/title-5.gif" name="Image19" width="294" height="25" border="0" id="Image19" alt="裝修工程及危險性作業通報"/></a></td>
  </tr>
  <tr>
    <td><img src="../images/title-line.gif" width="294" height="3" alt="*"/></td>
  </tr>
  <tr>
    <td><a href="index-a040.asp" title="局限空間申報" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image20','','../images/title-6.gif',1)"  onblur="MM_swapImgRestore()" onfocus="MM_swapImage('Image20','','../images/title-6.gif',1)"><img src="../images/title-6.gif" name="Image20" width="294" height="25" border="0" id="Image20" alt="局限空間申報"/></a></td>
  </tr>
  <tr>
    <td><img src="../images/title-line.gif" width="294" height="3" alt="*"/></td>
  </tr>
  <tr>
    <td><a href="index-g050.asp" title="工地自主稽查申報" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image21','','../images/title-7.gif',1)" onblur="MM_swapImgRestore()" onfocus="MM_swapImage('Image21','','../images/title-7.gif',1)"><img src="../images/title-7.gif" name="Image21" width="294" height="25" border="0" id="Image21" alt="工地自主稽查申報"/></a></td>
  </tr>
  <tr>
    <td><img src="../images/title-line.gif" width="294" height="3" alt="*"/></td>
  </tr>
  <tr>
    <td><a href="edumain.asp" title="教育訓練" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image22','','../images/title-8.gif',1)" onblur="MM_swapImgRestore()" onfocus="MM_swapImage('Image22','','../images/title-8.gif',1)"><img src="../images/title-8.gif" name="Image22" width="294" height="25" border="0" id="Image22" alt="教育訓練"/></a></td>
  </tr>
  <tr>
    <td><img src="../images/title-line.gif" width="294" height="3" alt="*"/></td>
  </tr>
  <tr>
    <td><a href="index-i01.asp" title="一呼百應網路回報" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image25','','../images/title-11.jpg',1)" onblur="MM_swapImgRestore()" onfocus="MM_swapImage('Image25','','../images/title-11.jpg',1)"><img src="../images/title-11.jpg" name="Image25" width="294" height="25" border="0" id="Image25" alt="一呼百應網路回報"/></a></td>
  </tr>
  <tr>
    <td><img src="../images/title-line.gif" width="294" height="3" alt="*"/></td>
  </tr>
    <tr>
    <td><a href="index-k01.asp" title="空氣品質簡易試算" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('Image24','','../images/title-12.jpg',1)" onblur="MM_swapImgRestore()" onfocus="MM_swapImage('Image24','','../images/title-12.jpg',1)"><img src="../images/title-12.jpg" name="Image24" width="294" height="25" border="0" id="Image24" alt="空氣品質簡易試算"/></a></td>
  </tr>
  <tr>
    <td><img src="../images/title-line.gif" width="294" height="3" alt="*"/></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>