<!--#include file="Conn/conn.asp" -->
<%'数据库连接文件%>
<!--#include file="js/Function.asp" -->
<%
GroupID = Request.Form("GroupID")							'接收GroupID的值
 Set rs = Server.CreateObject("ADODB.Command")				'创建Command对象
 rs.ActiveConnection =conn 									'定义Command对象的连接信息
 sql = "delete from tb_GroupMain where id in("&GroupID&")"	'应用delete语句实现记录的删除
 rs.CommandText = sql										'指定将要对数据库执行的操作
 rs.Execute													'执行sql语句
%>
<script>
alert("删除成功!")											//弹出提示对话框
window.location = "GList.asp"								//跳转到指定的页面
window.parent.parent.opener.location.reload() 				//重新加载页面
</script>