<%
	act = request.QueryString("act")
	Select Case act
		Case "lgname"
			if IsEmpty(session("lgname")) = false then
				session.Contents.Remove("lgname")
			end if
			response.Write("<script>alert('用户名删除成功');location='index.asp';</script>")
			
		Case "lgpwd"
			if IsEmpty(session("lgpwd")) = false then
				session.Contents.Remove("lgpwd")
			end if
			response.Write("<script>alert('密码删除成功');location='index.asp';</script>")
		Case "logout"
			session.Abandon()
			response.Write("<script>alert('欢迎下次再来');location='index.asp';</script>")
		
	End Select
%>
