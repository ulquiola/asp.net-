<%
Set rs=Server.CreateObject("ADODB.RecordSet")
SQL="Select * From tb_User Where UserName='"&Session("UserName")&"'"
rs.Open SQL,Conn,1,3
If rs.Eof and rs.Bof Then%>
	<script language="javascript">
		alert("�����û��Ѿ����ڣ������µ�¼��");
		window.location.href="index.asp";
	</script>
<%
Response.End()
End If
rs.close()
Set rs=Nothing
%>