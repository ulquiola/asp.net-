<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="Conn/Conn.asp"--><!--包含数据库连接文件-->
<%
If Request.QueryString("UID")<>"" Then					'判断UID是否有值
	Set rs=Server.CreateObject("ADODB.RecordSet")		'创建记录集
	SQL="Select * From tb_Topic Where Author="&Request.QueryString("UID")	'查询数据
	rs.Open SQL,conn,1,3													'打开记录集
	If Not(rs.Eof and rs.Bof) Then											'判断是否有记录信息
%>
		<script language="javascript">
		if(!confirm("该用户已经发帖，真的要删除该用户吗？")){//弹出提示对话框
			window.location.href="User_Manage.asp";//跳转到指定的ASP页面
		}
		</script>
	<%
	End If
		On Error Resume Next											'设置错误陷阱
		SQL="Delete From tb_User Where ID="&Request.QueryString("UID")'删除指定的用户信息
		Conn.Execute(SQL)												'执行SQL语句
		If Err<>0 Then
	%>
			<script language="javascript">
				alert("用户删除失败！");					//弹出提示对话框
				window.location.href="User_Manage.asp";		//跳转到指定的ASP页面
			</script>
		<%Else%>
			<script language="javascript">
				alert("用户删除成功！");				//弹出提示对话框
				window.location.href="User_Manage.asp";	//跳转到指定的ASP页面
			</script>
		<%End If
End If
%>