<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="Conn/Conn.asp"-->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
SQL="Select ID,status From tb_User Where UserName='"&Session("UserName")&"'"
rs.Open SQL,conn,1,3
Author=rs("ID")
flag=rs("status")
Face="face"&Request.Form("emote")&".gif"
TypeID=Request.Form("TypeName")
subject=Request.Form("subject")
content=replace(Request.Form("content"),"'","''")
IP=Request.Form("HIP")
If content<>"" Then
	On Error resume Next
	if flag="����" then
		Ins_T="Insert Into tb_topic (Author,Face,TypeID,subject,content,IP,state) values("&Author&",'"&Face&"',"&TypeID&",'"&subject&"','"&content&"','"&IP&"','��������')"
	else
		Ins_T="Insert Into tb_topic (Author,Face,TypeID,subject,content,IP) values("&Author&",'"&Face&"',"&TypeID&",'"&subject&"','"&content&"','"&IP&"')"
	end if
	Conn.Execute(Ins_T)	
	upd_T="Update tb_User Set SendRatio=SendRatio+1 Where ID="&Author  '�����û��ķ�����
	Conn.Execute(Upd_T)		
	If err<>0 Then
	%>
		<script language="javascript">
			alert("����ʧ�ܣ�");
			history.back();
		</script>
	<%Else%>
		<script language="javascript">
			alert("�����ɹ���");
			parent.mainwindow.location.href="main.asp";
		</script>
	<%End If
End If
%>