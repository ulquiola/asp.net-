<% if session("UserName") = "" then %>
<script type="text/javascript">
	alert("��ʱ�����ǷǷ���¼�������µ�¼һ�Σ�лл��");
	location="index.asp";
</script>
<% 
	response.End()
	end if
%>