<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<style>
body{
	margin:12px;
	font-size:12px;
}
</style>
</head>
<body>
<%
	if not IsEmpty(request.Form("lgname")) and not IsEmpty(request.Form("lgpwd")) then
		session("lgname") = request.Form("lgname")
		 session("lgpwd") = request.Form("lgpwd")
	end if
	if session("lgname") <> "" or session("lgpwd") <> "" then
%>

<div align="center"><a href="index_chk.asp?act=lgname">删除用户名</a>&nbsp;&nbsp;<a href="index_chk.asp?act=lgpwd">删除密码</a>&nbsp;&nbsp;<a href="index_chk.asp?act=logout">退出登录</a></div>
<%
	response.Write(session("lgname"))
		response.End()
  end if
%>
<table border="0" cellpadding="0" cellspacing="0" align="center">
<form method="post" action="index.asp">
	<tr>
		<td colspan="2" align="center" valign="middle" height="25px">用户登录</td>
	</tr>
	<tr>
		<td>用户名：</td>
		<td><input name="lgname" type="text" /></td>
	</tr>
	<tr>
		<td>密码：</td>
		<td><input name="lgpwd" type="password" /></td>
	</tr>
	<tr>
		<td colspan="2" height="25px" align="center" valign="middle"><input type="submit" value="登录" /></td>
	</tr>
</form>
</table>
</body>
</html>
