<!--#include file="Conn/conn.asp" -->
<!--#include file="js/Function.asp" -->
<%
GroupID = Request.Form("GroupID") 					'接收分组ID值
MemberID = Request.Form("MemberID") 				'接收好友的ID值
 Set rs = Server.CreateObject("ADODB.Command")		'创建Command对象
 rs.ActiveConnection =conn 							'定义Command对象的连接信息
 sql = "delete from tb_GroupMember where id in("&MemberID&")"	'应用delete语句实现删除
 rs.CommandText = sql								'指定将要对数据库执行的操作
 rs.Execute											'执行sql语句
%>
<script language="javascript">
alert("删除成功!")										//弹出提示对话框
window.location = "Group_xianxi.asp?ID=<%=GroupID%>"	//跳转到指定的页面
window.parent.parent.opener.location.reload() 			//重新加载页面
</script>
