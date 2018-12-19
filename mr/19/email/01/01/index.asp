<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
If Not Isempty(Request("send")) Then
on error resume next
Dim Mail
Set Mail=Server.CreateObject("CDO.Message")
Mail.From=Request.Form("mailfrom")  '设置发件人的邮箱地址
Mail.To=Request.Form("mailto")  '设置收件人的邮箱地址
Mail.Subject=Request.Form("mailsubject") '设定邮件主题
Mail.HtmlBody=Request.Form("mailbody")  '设定邮件内容
Mail.Send  '执行发送命令
Set Mail=Nothing  '释放对象
If err<>0 Then
  Response.Write("<script language='JavaScript'>alert('邮件发送失败,请核实输入内容是否准确!');history.back();</script>")
Else
  Response.Write("<script language='JavaScript'>alert('利用CDOSYS组件发送邮件成功!');window.location.href='index.asp';</script>")
End If
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>利用CDOSYS组件发送邮件</title>
<style type="text/css">
<!--
.style1 {color: #FF0000}
body,td,th {
	font-size: 12px;
}
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<script type="text/javascript">
function Mycheck(form){
  for(i=0;i<form.length;i++){
    if(form.elements[i].value==""){
	  alert("请填写完整的邮件信息!");return false;}		
  }
}
</script>
</head>
<body style="font-size:12px">
<center>
<table width="758" height="381" border="0" cellpadding="0" cellspacing="0">
  <tr align="center" valign="top">
    <td height="198" colspan="3"  background="images/banner.jpg">&nbsp;</td>
  </tr>
  <tr>
    <td height="510" colspan="3" align="center" valign="top" background="images/main.jpg"><table width="100%" height="510"  border="0" cellpadding="0" cellspacing="0">
      <tr align="center" valign="top">
        <td width="286" height="510"><img src="images/space.gif" width="286" height="510" border="0" usemap="#Map"></td>
        <td width="462"><table width="100%" height="507"  border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td height="83">&nbsp;</td>
          </tr>
          <tr>
            <td height="247">
			<table width="100%" height="249" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CC9933">
  <form name="form1" method="post" action="">
    <tr>
      <td height="26" colspan="2"  align="center" valign="middle"><p align="center">利用CDOSYS组件发送邮件</p></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td  align="right" valign="middle"><span class="style4">收件人Email： </span></td>
      <td  align="left" valign="middle"><span class="style4">
        <input name="mailto" type="text" id="mailto" size="30">
        <span class="style1">*</span> </span></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="right" valign="middle"><span class="style4">发件人Email： </span></td>
      <td align="left" valign="middle"><span class="style4">
        <input name="mailfrom" type="text" id="mailfrom" size="30">
        <span class="style1">*</span> </span></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="right" valign="middle">标
        题：</td>
      <td align="left" valign="middle"><span class="style4">
        <input name="mailsubject" type="text" id="mailsubject" size="30">        
      <span class="style1">*</span> </span></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td align="right" valign="middle">内
        容：</td>
      <td align="left" valign="middle"><span class="style4">
        <textarea name="mailbody" cols="43" rows="6" wrap="PHYSICAL" id="mailbody"></textarea>
        <span class="style1">*</span> </span></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height="38" colspan="2" align="center" valign="middle">
        <input name="send" type="submit" id="send" value="发送" onClick="return Mycheck(this.form)">
&nbsp;
        <input type="reset" name="Submit2" value="重写"></td>
    </tr>
  </form>
</table>
			</td>
          </tr>
          <tr>
            <td height="176">&nbsp;</td>
          </tr>
        </table></td>
        <td width="3%">&nbsp;</td>
      </tr>
    </table></td>
    </tr>
  <tr>
    <td height="93" colspan="3" align="center" valign="top" background="images/bottom.jpg">&nbsp;</td>
    </tr>
</table>
</center>
<map name="Map">
  <area shape="rect" coords="29,33,176,68" href="index.asp" target="_self">
  <area shape="rect" coords="31,73,175,105" href="#">
  <area shape="rect" coords="31,114,174,148" href="#">
  <area shape="rect" coords="31,154,174,186" href="#">
  <area shape="rect" coords="31,193,176,228" href="#">
  <area shape="rect" coords="31,233,177,268" href="#">
</map>
</body>
</html>
