<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="Conn/Conn.asp"--><!--包含数据库连接文件-->
<%
ID=Request.QueryString("ID")							'获取ID的值
If ID<>"" Then											'判断ID的值是否空
	Set rs=Server.CreateObject("ADODB.RecordSet")		'创建记录集
	SQL="Select * From tb_smalltype Where ID="&ID		'查询数据
	rs.Open SQL,conn,1,3								'打开记录集
	If Not(rs.Eof and rs.Bof) Then'判断是否有记录信息%>					
		<script language="javascript">
		if(!confirm("已经存在该版块区的主题信息，\n如果删除该版块，此时的主题也将一同删除，您真的要删除该版块吗？")){
			window.location.href="addsmall_manage.asp";//弹出提示对话框
		}else{
		window.location.href="smlallType_del_deal.asp?ID=<%=ID%>";//跳转到指定的ASP页面
		}
		</script>
	<%
	Else
		response.Redirect("smlallType_del_deal.asp?ID="&ID)			'跳转到指定的ASP页面
	End IF
	
End If
%>