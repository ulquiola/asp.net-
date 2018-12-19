<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn/conn.asp" -->
<%'包含数据库的连接文件%>
<!--#include file="js/Function.asp" -->
<%'包含数据验证文件%>
<%
Set rs = Server.CreateObject("ADODB.RecordSet")			'创建数据源
rs.ActiveConnection = conn 								'定义Command对象的连接信息
rs.Source = "SELECT  * FROM tb_groupMain where UserID="&Request.Cookies("UserID")&" and GroupName='"&request.form("GroupName")&"'"
'查询数据信息
rs.CursorType = 0										'定义游标类型
rs.CursorLocation = 2									'定义游标访问区域
rs.LockType =2											'只能同时被一个用户修改
rs.Open()												'打开数据源
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="javascript">
<%if not rs.EOF then %>
	alert("新添加的分组名称已存在!")
	//弹出提示对话框
<%
else
	rs.AddNew 												'应用Addnew方法实现数据信息的添加
	rs.fields("UserID").value=request.cookies("UserID") 	'获取用户ID值
	rs.fields("GroupName").value=request.form("GroupName") 	'获取要添加群组的名称
	rs.Update 												'进行数据更新
%>
	window.opener.frames("mainfrm").frames("left").location.reload() //跳转到指定的页面，并重新加载页面
	window.opener.opener.location.reload() 								
<%
end if
%>
window.close();														//关闭弹出提示对话框
</script>