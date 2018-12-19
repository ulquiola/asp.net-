<%
	lgname = request.form("lgname")
	lgpwd = request.Form("lgpwd")
	
	if lgname <> "" and lgpwd <> "" then
		response.Write("用户名或密码错误，请重新登录!")
		response.Write("&nbsp;&nbsp;<a href=index.asp>返回</a>")
	else
		response.Write("<script>alert('用户名和密码不允许为空');history.go(-1);</script>")
		response.End()
	end if
%>