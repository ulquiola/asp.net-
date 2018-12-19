<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp"-->
<% if request.querystring("id")<>"" then
id=request.QueryString("id")
end if 
set cc=server.CreateObject("adodb.recordset")
kk="select * from tb_books where id="&id
cc.open kk,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>明日在线搜索</title>
<style type="text/css">
<!--
.STYLE1 {font-size: 9pt}
-->
</style>
</head>

<body onLoad="clockon()" bgcolor="#eeeeee">
<table width="777" height="535" border="0">
  <tr>
    <td height="499" valign="top"><div align="center" class="STYLE1"><%=cc("name1")%><br>
        <%=replace(cc("introduce"),chr(13),"<br>")%><br>
        <%=cc("netaddress")%></div></td>
  </tr>
  <tr>
    <td height="30"> 
   </td>
  </tr>
</table>
</body>
</html>
