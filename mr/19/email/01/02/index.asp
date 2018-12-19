<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>如何使用Microsoft OutLook发送邮件</title>
<style type="text/css">
<!--
.style1 {color: #FF0000}
body,td,th {font-size: 12px;}
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
</head>
<body><center>
<script language="vbscript">
Sub sendmail()
Set outObj=CreateObject("Outlook.Application")
Set mailItem = outObj.CreateItem(0)
mailItem.Subject=document.all.Subject.value
mailItem.Body=document.all.content.value
mailItem.To=document.all.email.value
mailItem.send
document.write("<p align=center style='font-size:9pt'>邮件正在发送中,请及时查收!<br><a href=index.asp>返回</a></p>")
End Sub
</script>
<table width="774" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td width="783" height="134"><img src="images/banner.jpg" width="774" height="134" border="0" usemap="#Map"></td>
	</tr>
</table>
<table width="774" border="0" align="center" cellpadding="0" cellspacing="0">
     <tr align="center">
	 <td align="center" valign="top" background="images/style.jpg" width="22" height="426">&nbsp;</td>
      <td width="726" colspan="2" align="center" valign="middle">
        <table width="420" border="1" align="center" cellpadding="2" cellspacing="0" bordercolorlight="#FFFFFF" bordercolordark="#669933" bgcolor="#669933">
	    <form name="form1" method="post" action="">
        <tr align="center">
          <td height="30" colspan="2">使用Microsoft
            OutLook发送邮件</td>
        </tr>
        <tr bgcolor="#EDF8D0">
          <td height="26" align="right">收件人地址：</td>
          <td height="26" align="left" bgcolor="#EDF8D0"><input name="email" type="text" id="email">
              <span class="style1">*</span></td>
        </tr>
        <tr bgcolor="#EDF8D0">
          <td height="26" align="right">标题：</td>
          <td height="26" align="left" bgcolor="#EDF8D0"><input name="subject" type="text" id="subject">
              <span class="style1">*</span></td>
        </tr>
        <tr bgcolor="#EDF8D0">
          <td align="right">内容：</td>
          <td align="left" bgcolor="#EDF8D0"><textarea name="content" cols="40" rows="5" id="content"></textarea>
              <span class="style1">*</span></td>
        </tr>
        <tr align="center" bgcolor="#EDF8D0">
          <td height="33" colspan="2"><input name="send" type="submit" id="send" value="发送" onClick="vbscript:sendmail()">
      　
        <input type="reset" name="Submit" value="重置"></td>
        </tr>
        <tr align="center" bgcolor="#EDF8D0">
          <td height="22" colspan="2">注意:请填写完整信息内容</td>
        </tr>
	    </form>
      </table></td>
	  <td valign="top" width="26" height="426" background="images/style_1.jpg">&nbsp;</td>
    </tr>
</table> 
<table width="774" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td height="62" background="images/bottom.jpg"></td>
	</tr>
</table>
</center>
<map name="Map">
  <area shape="rect" coords="26,80,100,107" href="#">
  <area shape="rect" coords="109,82,175,109" href="index.asp" target="_self">
  <area shape="rect" coords="181,81,247,108" href="index.asp" target="_self">
  <area shape="rect" coords="255,80,318,110" href="#">
  <area shape="rect" coords="330,81,394,107" href="#">
  <area shape="rect" coords="403,82,466,109" href="#">
  <area shape="rect" coords="478,83,540,108" href="#">
  <area shape="rect" coords="549,84,636,107" href="#">
  <area shape="rect" coords="647,85,737,107" href="#">
</map>
</body>
</html>
