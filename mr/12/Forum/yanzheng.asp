<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
typeid=request.querystring("TypeID")					'获取表单元素ID的值并赋给ID变量
If Session("UserName")="" or Session("UserName")="Friend" Then%>
<script language="javascript">
alert("先注册成为用户，才可以发表主题信息！");
window.location.href="register.asp";
</script>
<%else%>
<script language="javascript">
window.location.href="Add_Topic.asp?typeid=<%=typeid%>";
</script>
<%end if%>
<%If Session("UserName")="" or Session("UserName")="Friend"  Then%>
<a href="#" onClick="Myopen(User,form_U)">用户登录</a>
<%Else%>
<a href="Modify.asp"> 修改资料</a>
<%End If%> 
┊ <a href="index.asp">查看主题分类</a> ┊ <a href="#" onClick="window.location.reload();">刷新页面</a> ┊ 
<%If Session("UserName")="" or Session("UserName")="Friend"  Then%>
<a href="#" onClick="Myopen(Manager,form_M)">管理员登录</a> ┊ <a href="friend_Add_Topic.asp">发表主题</a> 
<%Else%>
<a href="Logout.asp"> 注销用户</a>
<%End If%>