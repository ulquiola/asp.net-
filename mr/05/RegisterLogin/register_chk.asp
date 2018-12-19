<!--#include file="conn/conn.asp"-->
<!--#include file="inc/func.inc"-->
<%

	'获取添加数据
	name = request.QueryString("name")
	pwd = request.QueryString("pwd")
	question = request.QueryString("question")
	answer = request.QueryString("answer")
	realname = request.QueryString("realname")
	telephone = request.QueryString("telephone")
	qq = request.QueryString("qq")
	email = request.QueryString("email")
	
	'添加sql语句及字段
	sql = "select * from tb_member"
	dim flds(8),vle(8)
	flds(0) = "name"
	flds(1) = "password"
	flds(2) = "realname"
	flds(3) = "question"
	flds(4) = "answer"
	flds(5) = "email"
	flds(6) = "telephone"
	flds(7) = "qq"
	vle(0) = name
	vle(1) = pwd
	vle(2) = realname
	vle(3) = question
	vle(4) = answer
	vle(5) = email
	vle(6) = telephone
	vle(7) = qq
	call addrst(sql,flds,vle)
	session("name") = name
	response.Write(reback)
%>