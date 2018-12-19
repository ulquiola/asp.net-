<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="Conn/Conn.asp"-->
<%
Dim UserName,TrueName,PWD,Head,Sex,Email,Tel,homepage,QQ,Address,rs
UserName=Request.Form("UserName")
TrueName=Request.Form("TrueName")
PWD=Request.Form("PWD1")
ICO=Request.Form("ICO")
Sex=Request.Form("sex")
birthday=Request.Form("birthday")
Email=Request.Form("Email")
Tel=Request.Form("Tel")
homepage=Request.Form("homepage")
QQ=Request.Form("QQ")
Address=Request.Form("Address")
If UserName<>"" Then
	On error Resume Next   '设置错误陷阱
	Set rs=Server.CreateObject("ADODB.RecordSet")
	SQL1="Select * From tb_User Where UserName='"&UserName&"'"
	rs.Open SQL1,Conn,1,3
	If rs.BOF and rs.Eof Then   '判断用户是否存在
		SQL="Insert Into tb_User (UserName,PWD,TrueName,ICO,Sex,birthday,Email,"&_
		"Tel,homepage,QQ,Address) Values('"&UserName&"','"&PWD&"','"&TrueName&"','"&_
		ICO&"','"&Sex&"','"&birthday&"','"&Email&"','"&Tel&"','"&homepage&_
		"','"&QQ&"','"&Address&"')"
		Conn.Execute(SQL)
		If err<>0 Then  '错误处理
%>
		<script language="javascript">
			alert("<% Response.Write("错误提示: " & err.Description ) %>");
			window.location.href="register.asp";
		</script>
		<%Else
			Session("UserName")=UserName
		%>
			<script language="javascript">
				alert("用户注册成功！");
				parent.leftwindow.location.reload();
				window.location.href="main.asp";
			</script>
		<%
		End If
	Else
	%>
		<script language="javascript">
			alert("该用户名已经存在！");
			window.location.href="register.asp";
		</script>
	<% 
End If
	rs.Close()
	Set rs=Nothing
End If
%>

