<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<style>
body{
	margin:12px;
}
body,th,td{
	font-size:12px;
}
input{
	height: 12px;
}

</style>
</head>
<body>
<%
	response.Expires = 1
%>
<table border="0" cellspacing="0" cellpadding="0" align="center">
<form method="post" action="index_chk.asp">
<tr>
	<td colspan="2" height="25px" align="center" valign="middle" >�û���¼</td>
</tr>
<tr>
	<td>�û�����</td>
	<td><input name="lgname" type="text"  /></td>
</tr>
<tr>
	<td>��&nbsp;&nbsp;�룺</td>
	<td><input name="lgpwd" type="password" /></td>
</tr>
<tr>
	<td>��֤�룺</td>
	<td><input name="lgchk" type="text" style="width: 50px;" />&nbsp;&nbsp;<span style=" background-color: #FFFFCC;"><!--#include file="chkcode.asp"--></span></td>
</tr>
<tr>
	<td colspan="2" align="center" valign="middle" height="25px"><Input type="submit" value="��¼" style="height:20px;" /></td>
</tr>
</form>
</table>
</body>
</html>
