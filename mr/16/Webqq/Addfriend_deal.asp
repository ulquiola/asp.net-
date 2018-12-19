<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<!--#include file="js/Function.asp" -->
<%
UserID =split(Request.QueryString("ID"),",")		'应用split函数返回基于0的一维数组，其中包含指定数目的子字符串
GroupID = Request.QueryString("GroupID")			'获取分组名称的ID号
for i=0 to ubound(UserID)							'应用ubound函数获取最大可用下标
	Set rs = Server.CreateObject("ADODB.RecordSet")	'创建记录集
	rs.ActiveConnection = conn 						'记录集
	rs.Source = "SELECT  * FROM tb_groupMember WHERE UserID="&UserID(i)&" and GroupID="&GroupID&""	'数据信息查询
	rs.CursorType = 0
	rs.CursorLocation = 2
	rs.LockType =2
	rs.Open()
	if rs.EOF then
		rs.AddNew 									'应用AddNew方法实现数据添加
		rs.fields("UserID").value = UserID(i) 
		rs.fields("GroupID").value = GroupID
		rs.Update 
	end if
next
%>
<script language="javascript">
window.location="Group_xianxi.asp?ID=<%=GroupID%>"		//跳转到分组列表页面
window.parent.parent.opener.location.reload() 			//刷新父窗口
window.close();											//用户添加完成后，关闭指定的窗口
</script>
