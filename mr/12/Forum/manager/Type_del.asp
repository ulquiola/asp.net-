<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="Conn/Conn.asp"--><!--包含数据库连接文件-->
<%
QID=Request.QueryString("ID")							'为QID变量赋值
If QID<>"" Then											'判断ID值是否等于空
	Set rs=Server.CreateObject("ADODB.RecordSet")		'创建记录集
	SQL="Select * From tb_Topic Where TypeID="&QID		'查询数据
	rs.Open SQL,conn,1,3								'打开记录集
	If Not(rs.Eof and rs.Bof) Then'是否有记录信息%>					
		<script language="javascript">
		if(!confirm("已经存在该讨论区的主题信息，\n如果删除则该讨论区的主题也将一同删除，您真的要删除该类别吗？")){
			window.location.href="Type_Manage.asp";
			//跳转到指定的ASP页面
		}else{
		window.location.href="Type_del_deal.asp?ID=<%=QID%>";//跳转到指定的ASP页面
		}
		</script>
	<%
	Else
		response.Redirect("Type_del_deal.asp?ID="&QID)		'跳转到指定的ASP页面
	End IF
	
End If
%>