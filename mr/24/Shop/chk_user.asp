<%
if request("class")<>"class" then
	if session("user")="" then
		response.Write("<script>alert('���ȵ�¼��');window.location.href='index.asp';</script>")
	end if
else
	if session("user")="" then
		response.Write("<script>alert('���ȵ�¼��');parent.window.location.href='index.asp';</script>")	'��Ϊ����Ʒ��������ʹ�ÿ�ܽṹ
	end if
end if
%>