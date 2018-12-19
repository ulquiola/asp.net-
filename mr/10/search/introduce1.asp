<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp"-->
<%
if request.QueryString("ID")<>"" then
ID=request.QueryString("ID")
end if 
set r=server.CreateObject("adodb.recordset")
c="select * from tb_ITman where ID="&ID
r.open c,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>明日在线搜索</title>
<style type="text/css">
<!--
.STYLE2 {font-size: 9pt}
-->
</style>
</head>
<body bgcolor="#EEEEEE">
<table width="777" height="532" border="0">
  <tr>
    <td height="503" valign="top"><span class="STYLE2"><%=r("name1")%><br>
      <%=r("project")%><br>
      <%=replace(r("introduce"),chr(13),"<br>")%></span></td>
  </tr>
  <tr>
    <td height="23">
    </td>
  </tr>
</table>

</body>
</html>
