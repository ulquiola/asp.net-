<!--#include file="conn/conn.asp"-->
<!--#include file="inc/func.inc"-->
<%
	name = request.QueryString("name")
	if not isEmpty(name) then
		sql = "select id from tb_member where name = '"&name&"'"
		chkform(sql)
		response.Write(reback)
	else
		response.Write("3")
	end if
%>