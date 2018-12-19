<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% 
If Not Isempty(Request("send")) Then 
Set msg = Server.CreateObject("JMail.Message") 
msg.silent = true 
msg.Logging = true 
msg.Charset = "gb2312" 
msg.MailServerUserName =request("mailfrom")  '输入smtp服务器验证登陆名 （邮局中任何一个用户的Email地址） 
msg.MailServerPassword = request("mima") '输入smtp服务器验证密码 （用户Email帐号对应的密码） 
msg.From =request("mailfrom")  '发件人Email地址 
msg.FromName = "pyj" '发件人姓名 
msg.AddRecipient(request("emails")) '收件人Email地址 
msg.Subject = request("mailsubject") '信件主题 
msg.Body = request("mailbody") '正文 
msg.Send (request("smtpaddress")) '发件smtp服务器地址（企业邮局地址） 
set msg = nothing 
response.Write "<br>"
response.Write "发送成功！！"
End If
%> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>博客----发送邮件</title>
</head>
<body>
<table width="500" height="281" border="0" align="center">
  <tr>
    <td height="29" colspan="2"  align="center" valign="middle"><p align="center">输入信件内容</p></td>
  </tr>
  <tr>
    <td width="127" height="22" align="right" valign="middle"><span class="style4">收件人email： </span></td>
    <td width="363" align="left" valign="middle"><span class="style4">
      <input name="emails" type="text" id="emails">
    </span></td>
  </tr>
  <tr>
    <td height="22" align="right" valign="middle"><span class="style4">发信人email： </span></td>
    <td align="left" valign="middle"><span class="style4">
      <input name="mailfrom" type="text" id="mailfrom">
    </span></td>
  </tr>
  <tr>
    <td height="18" align="right" valign="middle"><span class="style4">密码： </span></td>
    <td height="18" align="left" valign="middle"><span class="style4">
      <input name="mima" type="password" id="mima">
    </span></td>
  </tr>
  <tr>
    <td height="21" align="right" valign="middle"><span class="style4">SMTP服务器地址： </span></td>
    <td height="21" align="left" valign="middle"><span class="style4">
      <input name="smtpaddress" type="text" id="smtpaddress">
    </span></td>
  </tr>
  <tr>
    <td height="18" align="right" valign="middle"><span class="style4">标题：　 </span></td>
    <td height="18" align="left" valign="middle"><span class="style4">
      <input name="mailsubject" type="text" id="mailsubject">
    </span></td>
  </tr>
  <tr>
    <td height="83" align="right" valign="middle"><span class="style4">　内容：　 </span></td>
    <td height="83" align="left" valign="middle"><span class="style4">
      <textarea name="mailbody" cols="45" rows="6" wrap="PHYSICAL" id="mailbody"></textarea>
    </span></td>
  </tr>
  <tr>
    <td height="37" colspan="2" align="center" valign="middle">
      <input name="send" type="submit" id="send" value="发送">
&nbsp;
    <input type="reset" name="Submit2" value="重写">    </td>
  </tr>
</table>
</body>
</html>
