<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn/conn.asp" -->
<%'�������ݿ�������ļ�%>
<!--#include file="js/Function.asp" -->
<%'����������֤�ļ�%>
<%
Set rs = Server.CreateObject("ADODB.RecordSet")			'��������Դ
rs.ActiveConnection = conn 								'����Command�����������Ϣ
rs.Source = "SELECT  * FROM tb_groupMain where UserID="&Request.Cookies("UserID")&" and GroupName='"&request.form("GroupName")&"'"
'��ѯ������Ϣ
rs.CursorType = 0										'�����α�����
rs.CursorLocation = 2									'�����α��������
rs.LockType =2											'ֻ��ͬʱ��һ���û��޸�
rs.Open()												'������Դ
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="javascript">
<%if not rs.EOF then %>
	alert("����ӵķ��������Ѵ���!")
	//������ʾ�Ի���
<%
else
	rs.AddNew 												'Ӧ��Addnew����ʵ��������Ϣ�����
	rs.fields("UserID").value=request.cookies("UserID") 	'��ȡ�û�IDֵ
	rs.fields("GroupName").value=request.form("GroupName") 	'��ȡҪ���Ⱥ�������
	rs.Update 												'�������ݸ���
%>
	window.opener.frames("mainfrm").frames("left").location.reload() //��ת��ָ����ҳ�棬�����¼���ҳ��
	window.opener.opener.location.reload() 								
<%
end if
%>
window.close();														//�رյ�����ʾ�Ի���
</script>