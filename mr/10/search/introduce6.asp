<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp"-->
<%
 if request.QueryString("id")<>"" then
 id=request.QueryString("id")
 end if
set ac=server.CreateObject("adodb.recordset")
bc="select * from tb_programme where id="&id
ac.open bc,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>明日在线搜索</title>
</head>

<body bgcolor="#eeeeee">
<table width="777" height="793" border="0">
  <tr>
    <td valign="top"> <%=ac("name1")%><br><%=replace(ac("introduce"),chr(13),"<br>")%><br><%=ac("netaddress")%>
   </td>
  </tr>
</table>
</body>
</html>
