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
On error Resume Next   '设置错误陷阱
SQL="Update tb_User Set PWD='"&PWD&"',TrueName='"&TrueName&"',ICO='"&ICO&"',Sex='"&_
Sex&"',birthday='"&birthday&"',Email='"&Email&"',Tel='"&Tel&"',homepage='"&homepage&_
"',QQ='"&QQ&"',Address='"&Address&"' Where UserName='"&UserName&"'"
Conn.Execute(SQL)
If err<>0 Then  '错误处理 
%>
	<script language="javascript">
		alert("<% Response.Write("错误提示: " & err.Description ) %>");
		window.location.href="Modify.asp";
	</script>
<%Else%>
	<script language="javascript">
		alert("用户信息修改成功！");
		window.location.href="modify.asp";
	</script>
<%End If
rs.Close()
Set rs=Nothing
%>

