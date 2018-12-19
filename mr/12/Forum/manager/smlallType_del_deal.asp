<!--#include file="Conn/Conn.asp"--><!--包含数据库连接文件-->
<%
If Request.QueryString("ID")<>"" Then				'判断接收的ID值是否等于空
	On Error Resume Next							'设置错误陷阱
	SQL="Delete From tb_smalltype Where ID="&Request.QueryString("ID")	'删除指定的记录
	Conn.Execute(SQL)								'执行SQL语句
	If Err<>0 Then						
%>
		<script language="javascript">
			alert("信息删除失败！");				//弹出提示对话框
			window.location.href="addsmall_manage.asp";//跳转到指定的ASP页面
		</script>
	<%Else%>
		<script language="javascript">
			alert("信息删除成功！");				//弹出提示对话框
		</script>
	<%
		response.Redirect("addsmall_manage.asp")	'跳转到指定的ASP页面
	End If
End If
%>