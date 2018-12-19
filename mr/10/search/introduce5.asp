<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp"-->
<%
if request.QueryString("id")<>"" then 
id=request.QueryString("id")
end if 
set ab=server.CreateObject("adodb.recordset")
ab_1="select * from tb_open where id="&id
ab.open ab_1,conn,1,3
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

<body bgcolor="#eeeeee">
<table width="777" height="536" border="0">
  <tr>
    <td height="505" valign="top"><span class="STYLE2"><%=ab("name1")%><br>
      <%=replace(ab("introduce"),chr(13),"<br>")%></span></td>
  </tr>
  <tr>
    <td height="25">
   </td>
  </tr>
</table>

</body>
</html>
