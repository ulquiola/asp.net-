<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Conn/conn.asp" -->
<%
set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = MR_website_STRING
Rs.Source = "SELECT * FROM website ORDER BY idd DESC"
Rs.CursorType = 0
Rs.CursorLocation = 2
Rs.LockType = 3
Rs.Open()
Rs_numRows = 0
%>
<html>
<head>
<title>注册2</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link 
href="css/css1.css" rel=stylesheet type="text/css">
<style type="text/css">
<!--
.pt9 {  font-family: "宋体"; font-size: 9pt; line-height: 13pt}
.pt105 {  font-family: "宋体"; font-size: 10.5pt}
-->
</style>
</head>
<body bgcolor="#FF0000" text="#000000" leftmargin="0" topmargin="3">
<b><font color="#FFFFFF"> </font></b>
<table width="430" border="0" cellspacing="2" cellpadding="4" align="center">
  <tr>
    <td class="pt105"><b><font color="#FFFFFF">第二步：注册成功</font></b></td>
  </tr>
</table>
<table width="430" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#FFFFFF">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td height="80"><div align="center"><font size="3" color="#990000"><b><font size="5">注册成功</font></b></font></div></td>
  </tr>
  <tr>
    <td><table width="83%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <%
		dim userid
		userid=Rs.Fields.Item("id").Value
		set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MR_website_STRING
Recordset1.Source = "SELECT * FROM website WHERE id = '"&userid&"'"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 3
Recordset1.Open()
Recordset1_numRows = 0
%>
          <td class="pt9">您的网址是：<a href="../templet<%=(Recordset1.Fields.Item("xuanze").Value)%>/page1.asp?id=<%=(Recordset1.Fields.Item("id").Value)%>" target="_blank" class="pt105">www.mrbccd.com/website.asp?id=<%=(Rs.Fields.Item("id").Value)%></a></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="30">&nbsp;</td>
  </tr>
  <tr>
    <td><div align="center">
        <input type="button" name="B1" onClick="window.close();window.opener.location='index.asp'" value=" 立即登录 " class="pt9n" >
      </div></td>
  </tr>
  <tr>
    <td height="60">&nbsp;</td>
  </tr>
</table>
</body>
</html>
<%
Rs.Close()
%>
