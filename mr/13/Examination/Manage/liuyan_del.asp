<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../conn/conn.asp"-->
<%
if request.QueryString("id")<>"" then
	Del="Delete from tb_board where ID="&request.QueryString("id")
	conn.execute(Del)
%>
<script language="javascript">
window.location.href='liuyan.asp';
</script>
<%
end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
</head>

<body>
</body>
</html>
