<%
	act = request.QueryString("act")
	Select Case act
		Case "lgname"
			if IsEmpty(session("lgname")) = false then
				session.Contents.Remove("lgname")
			end if
			response.Write("<script>alert('�û���ɾ���ɹ�');location='index.asp';</script>")
			
		Case "lgpwd"
			if IsEmpty(session("lgpwd")) = false then
				session.Contents.Remove("lgpwd")
			end if
			response.Write("<script>alert('����ɾ���ɹ�');location='index.asp';</script>")
		Case "logout"
			session.Abandon()
			response.Write("<script>alert('��ӭ�´�����');location='index.asp';</script>")
		
	End Select
%>
