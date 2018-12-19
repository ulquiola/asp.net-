<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>自动输入用户名密码</title>
<style>
body{
margin:12px;
}
body,th,td{
	font-size:12px;
}
</style>
</head>
<body>
<%
	if Request.Cookies("login").HasKeys then
		lgname = Request.Cookies("login")("lgname")
		lgpwd = Request.Cookies("login")("lgpwd")
	end if 
%>
<table border="0" cellspacing="0" cellpadding="0" align="center">
<form method="post" action="index_chk.asp">
<tr>
	<td colspan="2" height="25px" align="center" valign="middle" >用户登录</td>
</tr>
<tr>
	<td>用户名：</td>
	<td><input name="lgname" type="text" value="<%=lgname%>" /></td>
</tr>
<tr>
	<td>密&nbsp;&nbsp;码：</td>
	<td><input name="lgpwd" type="password" value="<%=lgpwd%>" /></td>
</tr>
<tr>
	<td colspan="2" align="center" valign="middle" height="25px"><input name="autolg" type="checkbox" value="1" <% if Request.Cookies("login").HasKeys then response.Write("checked=checked") end if %> />记住密码&nbsp;&nbsp;<Input type="submit" value="登录" /></td>
</tr>
</form>
</table>
</body>
</html>
