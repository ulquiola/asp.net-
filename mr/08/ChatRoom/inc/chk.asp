<% if session("UserName") = "" then %>
<script type="text/javascript">
	alert("超时，或是非法登录，请重新登录一次，谢谢！");
	location="index.asp";
</script>
<% 
	response.End()
	end if
%>