<%
	lgname = request.form("lgname")
	lgpwd = request.Form("lgpwd")
	
	if lgname <> "" and lgpwd <> "" then
		response.Write("�û�����������������µ�¼!")
		response.Write("&nbsp;&nbsp;<a href=index.asp>����</a>")
	else
		response.Write("<script>alert('�û��������벻����Ϊ��');history.go(-1);</script>")
		response.End()
	end if
%>