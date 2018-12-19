<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
session.Timeout=10    '设置Session变量的期限
flag=false
If Request.Form("Submit")="进入聊天室" or trim(request.form("UserName")) <> "" Then
	UserName=trim(Request.Form("UserName"))
	head=Request.Form("head")
	UserIP=Request.ServerVariables("REMOTE_ADDR")
	A_Username=split(Application("UserName"),",")
	for i=0 to uBound(A_Username)
		If A_username(i)=username then
			flag = true
			%>
	<script language="javascript">
		alert("该用户昵称已经存在，请换一个吧!");
		location="login.asp";
	</script>
	<%
			exit for
		End if
	Next
	if flag=false Then
		application.Lock()
		application("UserName")=application("UserName")+","+UserName
		session("UserName")=UserName
		Application("head")=application("head")+","+head
		Application("UserIP")=application("UserIP")+","+UserIP
		application.UnLock()
%>
<script>
window.open('chatroom.asp','chat','width=777,height=600',false);
</script>
<%
	End if
End if
%>
<script language="javascript">
	function check(){
		if(form1.Username.value==""){
			alert("昵称不能为空！");
			form1.Username.focus();
			return false;
		}
	}
	</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>聊天室</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin:0px;
	background-image: url(images/bg.jpg);
	background-position:center;
	background-repeat:no-repeat;
}
body,th,td{
	font-size: 12px;

}
.STYLE1 {color: #0849AD;}
-->
</style></head>
<body>
<table width="635" height="346"  border="0" cellpadding="0" cellspacing="0" align="center">
<form method="post" action="login.asp" name="form1">
	<tr>
		<td height="113">&nbsp;</td>
	</tr>
	<tr>
		<td height="205">
			<table border="0" cellpadding="0" cellspacing="0" align="center" width="100%" height="100%">
				<tr>
					<td width="90">&nbsp;</td>
					<td width="312">
					<table width="312" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td>
			<table  border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#C6E3F7">
				<tr>
					<td align="center">
						<input name="head" type="radio" value="1.GIF" checked="checked" />
					</td>
					<td align="center">
						<div align="center"> &nbsp;<img src="headICO/1.GIF" width="32" height="32" /></div>
					</td>
					<td align="center">
						<input name="head" type="radio" value="2.GIF" />
					</td>
					<td>
						<div align="center"><img src="headICO/2.GIF" width="32" height="32" /></div>
					</td>
					<td align="center">
						<input name="head" type="radio" value="3.GIF" />
					</td>
					<td>
						<div align="center"><img src="headICO/3.GIF" width="32" height="32" /></div>
					</td>
						<td align="center">
						<input name="head" type="radio" value="10.GIF" />
					</td>
					<td>
						<div align="center"><img src="headICO/10.GIF" width="32" height="32" /></div>
					</td>
				</tr>
				<tr>
					<td align="center">
						<input name="head" type="radio" value="5.GIF" />
					</td>
					<td align="center">
						<div align="center"><img src="headICO/5.GIF" width="37" height="34" /></div>
					</td>
					<td align="center">
						<input name="head" type="radio" value="4.GIF" />
					</td>
					<td>
						<div align="center"><img src="headICO/4.GIF" width="33" height="32" /></div>
					</td>
					<td align="center">
						<input name="head" type="radio" value="9.GIF" />
					</td>
					<td>
						<div align="center"><img src="headICO/9.GIF" width="32" height="32" /></div>
					</td>
					<td align="center">
						<input name="head" type="radio" value="11.GIF" />
					</td>
					<td>
						<div align="center"><img src="headICO/11.GIF" width="32" height="32" /></div>
					</td>
				</tr>
				<tr>
					<td align="center">
						<input name="head" type="radio" value="7.GIF" />
					</td>
					<td align="center">
						<div align="center"><img src="headICO/7.GIF" width="35" height="33" /></div>
					</td>
					<td align="center">
						<input name="head" type="radio" value="6.GIF" />
					</td>
					<td>
						<div align="center"><img src="headICO/6.GIF" width="32" height="32" /></div>
					</td>
					<td>
						<input name="head" type="radio" value="8.GIF" />
					</td>
					<td>
						<div align="center"><img src="headICO/8.GIF" width="35" height="34" /></div>
					</td>
					<td align="center">
						<input name="head" type="radio" value="12.GIF" />
					</td>
					<td>
						<div align="center"><img src="headICO/12.GIF" width="32" height="32" /></div>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="27">
			<div align="center"><span class="style1">昵称：</span>
			  <input name="Username" type="text" class="Sytle_text" id="Username3" onKeyDown="Javascript:if(event.keyCode==13){return check()}"></div>
		</td>
	</tr>
	<tr>
		<td height="27">
			<div align="center"><input name="Submit" type="Submit" class="Style_button_del" value="" onClick="return check();"></div>
		</td>
	</tr>
</table>			
					</td>
					<td width="233">&nbsp;</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
</form>
</table>
</body>
</html>
