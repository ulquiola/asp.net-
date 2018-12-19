<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="Conn/Conn.asp"-->
<%
TopicID=Request.Form("HTopicID")
subject=Request.Form("subject")
content=replace(Request.Form("content"),"'","''")
username=Request.Form("UID")
IP=Request.Form("HIP")
Set rs=server.CreateObject("ADODB.RecordSet")
sql="select * from tb_topic where ID="&TopicID
rs.open sql,conn,1,3
rs.close
set rs=nothing
If subject<>"" Then
	On Error resume Next
	Ins="Insert Into tb_Reply (TopicID,author,subject,Content,IP) values("&TopicID&","&username&",'"&subject&"','"&content&"','"&IP&"')"

	Conn.Execute(Ins)
	If err<>0 Then
	%>
		<script language="javascript">
			alert("信息回复失败！<% Response.Write("错误提示: " & err.Description ) %>");
			history.back();
		</script>
	<%Else%>
		<script language="javascript">
			alert("信息回复成功！");
			parent.mainwindow.location.href="main.asp";
			</script>
<%
End If
End If
%>
