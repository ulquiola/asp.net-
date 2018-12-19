<!--#include file="user.asp"-->
<%
if request.Form("user_name")<>"" and request.Form("user_pwd")<>"" then
	user_name=request.Form("user_name")
	user_pwd=request.Form("user_pwd")
	if user_name=default_user_name and user_pwd=default_user_pwd then
%>
<script language="javascript">alert("登录成功!");window.location.href="index.asp";</script>
<%
	else
%>
<script language="javascript">alert("用户名或密码错误!");window.location.href="login.asp";</script>
<%
	end if
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>管理员登录</TITLE>
<meta http-equiv="Content-Type" content="text/html;charset=gb2312">
<link rel="stylesheet" type="text/css" href="../css/css.css">
</HEAD>
<BODY>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<table width="520" height='360' border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#ccddee" background="../images/login.jpg">
  <tr class='toptitle1'> 
	<td width='47' height="157">&nbsp;</td>
	<td colspan=2>&nbsp;</td>
  </tr>
  <tr height='10px'><td colspan=3></td></tr>
  <form name='login' method='post' action='../manage/login.asp'>
  <tr height='34px'>
	<td height="30"></td>
	<td width='52' height="30">用户名：</td>
	<td width="421" height="30"><input type='text' id='user_name' name='user_name' size=25></td>
  </tr>
  <tr height='34px'>
	<td height="30"></td>
	<td height="30">密&nbsp;&nbsp;码：</td>
	<td height="30"><input type='password' name='user_pwd' size=25></td>
  </tr>
  <tr height='34px'>
	<td height="30"></td>
	<td height="30"></td>
	<td height="30">
	<input name="submit" type='submit' class="btn1" id="submit" style="line-height:18px;height:20px" value=" 确定 ">&nbsp;
	<input name="Submit1" type="reset" class="btn1" id="Submit1" style="line-height:18px;height:20px" value=" 清除 "></td>
  </tr>
  </form>
  <tr></tr>
</table>
<script language='JavaScript'>
//焦点定位到用户名输入框
login.user_name.focus();
</script>
</BODY>
</HTML>
