<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="include/conn.asp" -->
<!--#include file="include/md5.asp" -->
<%
if request("login")="out" then	'笔者将用户登录和退出写在同一个文件内，以接收到的 login 值进行判断
	session("cishu")=""
	session("shijian")=""
	session("user")=""
	response.Redirect("index.asp")	'清除所有与用户有关的信息，并转向到首页
	response.End()
end if
sql="select * from [user] where name='"&trim(request("user"))&"';"	'按用户名进行查询
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
if not rs.eof then	'如果用户名存在
	if rs("pass")=md5(trim(request("pass"))) then
	'将数据库中存储的密码和用户提交的密码进行比较 
		session("shijian")=rs("shijian2")
		session("cishu")=rs("cishu")
		session("user")=trim(request("user"))	'将相关的信息写到 session 内，以便随时获取
		rs("shijian2")=now()	'最后一次登录时间，也就是当前时间
		rs("cishu")=rs("cishu")+1	'登录次数加1
		rs.update
		response.Redirect("index.asp")	'成功后转到首页
	else	'如果密码不一样
		session("user")=""
		session("cishu")=""
		session("shijian")=""
		response.Write("<script>alert('用户名或密码错误！');window.location.href='index.asp';</script>")
		response.End() '清除所有与用户有关的信息，并转向到首页
	end if
else	'如果用户名不存在
	session("user")=""
	session("cishu")=""
	session("shijian")=""
	response.Write("<script>alert('用户名或密码错误！');window.location.href='index.asp';</script>")
	response.End() '清除所有与用户有关的信息，并转向到首页
end if
rs.close
set rs=nothing
%>