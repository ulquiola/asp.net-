<!--#include file="conn/conn.asp" -->
<%'数据库连接文件%>
<!--#include file="js/Function.asp" -->
<%'包含数据验证的页面%>
<%
Set rs = Server.CreateObject("ADODB.RecordSet")				'创建记录集
rs.ActiveConnection = conn 									'
rs.Open "tb_talk",,1,3										'打开数据源
rs.AddNew 
rs.fields("fuserid").value = request.form("fuserid")				'发送信息用户的ID
rs.fields("fusernickname").value = request.form("fusernickname")	'发送信息用户的昵称
rs.fields("touserid").value = request.form("touserid") 				'接受信息的用户ID
rs.fields("message").value = request.form("Message")				'获取聊天信息
rs.Update 															'进行数据更新
%>
<script language="javascript">
alert('信息发送成功');												//弹出提示信息
window.close();														//关闭弹出页面
</script>