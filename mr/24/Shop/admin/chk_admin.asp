<!--#include file="../include/md5.asp" -->
<%
if session("admin")="" then
	response.Write("<script>alert('非法操作！');window.location.href='login.asp';</script>")
else
	sql="select * from [admin] where username='"&session("admin")&"';"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	if not rs.eof then
		if rs("pass")<>session("admin_pass") then
			response.Write("<script>alert('非法操作！');window.location.href='login.asp';</script>")
		end if
	else
		response.Write("<script>alert('非法操作！');window.location.href='login.asp';</script>")
	end if
end if
%>