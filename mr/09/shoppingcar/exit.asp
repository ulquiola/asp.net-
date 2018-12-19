<%
session.Abandon()					'应用Abandon方法来清空Session对象
response.Redirect("index.asp")		'应用Redirect方法跳转到指定的动态页面
%>