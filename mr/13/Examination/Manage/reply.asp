<!--#include file="../conn/conn.asp"-->
<%
ID=request.Form("ID")
if ID<>"" then
	reply=replace(request.Form("content1"),"'","''")
	on error resume next	
	sql="update tb_board set reply='"&reply&"' where ID="&ID
	conn.execute(sql)
	if err<>0 then
%>
<script type="text/javascript">
alert('�ظ�ʧ��');
history.go(-1);
</script>
<%
	else
		response.Write("<script language='javascript'>alert('�ظ��ɹ�!');opener.location.reload();window.close();</script>")
	end if
else
	response.Write("error")
end if
%>