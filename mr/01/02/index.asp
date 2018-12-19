<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<style type="text/css">
body{
	margin:12px;
	font-size: 12px;
}

</style>
</head>
<body>
<%
dim ary(7)
ary(0)="朋友下午好，敲几下键盘，打上N个你好，是否在聊天，灌水~~~~"
ary(1)="老朋友，想你的伙伴了吧！赶快来社区留言吧！"
ary(2)="不管你走到哪里，这里永远是你最温暖的家！"
ary(3)="新来的朋友，你们好！阳光社区欢迎你们。"
ary(4)="想成功吗？集体的智慧是你坚强的依靠！"
ary(5)="还是单身吗？这里可以找到你的另一半！"
ary(6)="工作累吗？有时间的话就来这里放松放松！"
Randomize     '初始化随机数生成器，该生成器具有基于系统计时器的种子
num = int(rnd() * 10)
select case num  
	case 0
%>
		<%=ary(0) %>
	<% case 1 %>
		<%=ary(1) %>
	<% case 2 %>
		<%=ary(2) %>
	<% case 3 %>
		<%=ary(3) %>
	<% case 4 %>
		<%=ary(4) %>
	<% case 5 %>
		<%=ary(5) %>
	<% case else %>
		<%=ary(6) %>	 
<% end select %>
</body>
</html>
