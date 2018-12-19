<%
if request("class")<>"class" then
	if session("user")="" then
		response.Write("<script>alert('请先登录！');window.location.href='index.asp';</script>")
	end if
else
	if session("user")="" then
		response.Write("<script>alert('请先登录！');parent.window.location.href='index.asp';</script>")	'因为在商品分类我们使用框架结构
	end if
end if
%>