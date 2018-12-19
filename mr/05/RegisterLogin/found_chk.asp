<!--#include file="conn/conn.asp"-->
<!--#include file="inc/func.inc"-->
<%
	name = request.QueryString("foundname")
	question = request.QueryString("question")
	answer = request.QueryString("answer")
	sql = "select password from tb_member where name = '"&name&"' "
	call chkform(sql)
	if reback = 2 then
		response.Write(reback)
		response.End()
	end if
	sql = sql & "and question = '"&question&"' "
	sql = sql & "and answer = '"&answer&"' "
	call chkform(sql)
	if reback = 2 then
		response.Write(reback)
		response.End()
	end if
	call rebackrst(sql,"password")
	response.Write(reback)
%>