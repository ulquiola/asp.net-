<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>公告信息浏览</title>
<link rel="stylesheet" type="text/css" href="Css/style.css">
<style>
<!--
.yellow02{color:#E56B04; font-weight:bold}
-->
</style>
</head>
<body>
<table width="640" height="360" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="40">&nbsp;</td>
  </tr>
  <tr>
    <td>
	<table width="558" border="0" align="center" cellpadding="0" cellspacing="0">
<%
id=Request.QueryString("id")
If Not IsNumeric(id) Then
Else
Set rs=conn.Execute("select Aftitle,Afcontent from tb_affiche where id="&id&"")
%>
  <tr>
    <td height="35" align="left" style="text-indent:35px ">公告信息浏览</td>
  </tr>
  <tr>
    <td height="35" align="center" class="yellow02"><%=rs("Aftitle")%></td>
  </tr>
  <tr>
    <td height="150" align="left" valign="top" style="text-indent:10px"><%=rs("Afcontent")%></td>
  </tr>
  <tr>
    <td height="22" align="center"><form name="form1" method="post" action="">
        <input type="button" name="Submit" value="关 闭" onClick="javascript:window.close();">
    </form></td>
  </tr>
  <%
Set rs=Nothing
End If%>
</table>
	</td>
  </tr>
  <tr>
    <td height="40">&nbsp;</td>
  </tr>
</table>
</body>
</html>
