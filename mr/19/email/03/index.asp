<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Response.Flush()
If Not Isempty(Request("send")) Then
'运行程序前请到http://www.aspemail.com下载并安装AspEmail组件
Set Mail=Server.CreateObject("Persits.MailSender")
Mail.Host="192.168.1.41"  
Mail.Port=25
Mail.IsHTML=True
Mail.Priority=1
'Mail.CharSet="gb2312"
Mail.From=Request.Form("mailfrom")
Mail.FromName="happy"
Mail.AddAddress Request.Form("mailto"),"MRLX"
Mail.Subject = Request.Form("subject")
Mail.Body = Request.Form("content")
'Mail.Queue = True
Mail.Send
Response.Write("<script language='javascript'>alert('使用AspEmail组件发送邮件成功!');window.location.href='index.asp';</script>")
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>使用AspEmail组件发送邮件</title>
<script type="text/javascript">
function Mycheck(form){
  for(i=0;i<form.length;i++){
    if(form.elements[i].value==""){
	  alert("请填写完整的邮件信息!");return false;}		
  }
}
</script>
<style type="text/css">
<!--
body,td,th {
	font-size: 9pt;
}
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></head>
<body style="font-size:12px">
<center>
<table width="755" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td height="89"><img src="images/banner.jpg" width="755" height="89" border="0" usemap="#Map"></td>
	</tr>
</table>
<table width="755" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="404" align="center" background="images/middle.jpg">	<table width="550" border="0" align="center" cellpadding="3" cellspacing="1" bordercolorlight="#00FF00" bordercolordark="#000000" bgcolor="#F5FDFF">
      <form name="form1" method="post" action="">
        <tr align="center">
          <th height="22" colspan="2">利用AspEmail组件发送邮件</th>
        </tr>
        <tr>
          <td width="181" height="30" align="right">收件人Email地址：</td>
          <td width="354" height="30" align="left"><input name="mailto" type="text" id="mailto" size="30"></td>
        </tr>
        <tr>
          <td height="30" align="right">发件人Email地址：</td>
          <td height="30" align="left"><input name="mailfrom" type="text" id="mailfrom" size="30"></td>
        </tr>
        <tr>
          <td height="30" align="right">SMTP服务器地址：</td>
          <td height="30" align="left"><input name="smtp" type="text" id="smtp" size="30"></td>
        </tr>
        <tr>
          <td height="30" align="right">邮件标题：</td>
          <td height="30" align="left"><input name="subject" type="text" id="subject" size="30"></td>
        </tr>
        <tr>
          <td height="30" align="right">邮件内容：</td>
          <td height="30" align="left"><textarea name="content" cols="35" rows="4" id="content"></textarea></td>
        </tr>
        <tr>
          <td height="30" align="right">&nbsp;</td>
          <td height="30" align="left"><input name="send" type="submit" id="send" value="发 送" onClick="return Mycheck(this.form)"></td>
        </tr>
      </form>
    </table></td>
  </tr>
</table>
<table width="755" border="0" cellpadding="0" cellspacing="0">
<tr>
	<td height="69" background="images/bottom.jpg"></td>
</tr>
</table>
</center>
<map name="Map">
  <area shape="rect" coords="406,65,454,81" href="#">
  <area shape="rect" coords="473,65,521,81" href="index.asp" target="_self">
  <area shape="rect" coords="604,65,653,80" href="#">
  <area shape="rect" coords="670,63,718,83" href="#">
  <area shape="rect" coords="538,65,591,79" href="index.asp" target="_self">
</map>
</body>
</html>
