<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
'��ȡ�û���Ϣ
UserName=request.Form("UserName")				'ΪUserName������ֵ
Password1=request.Form("Password1")				'ΪPassword1������ֵ
sex=request.form("sex")							'Ϊsex������ֵ
birthday=request.form("birthday")
touxiang=request.form("touxiang")
if request.form("touxiang")="" then
	touxiang=0
end if
Email=request.form("Email")
'�ж��û������Ƿ�Ϊ��
if UserName<>"" then
    '������¼��
	set rs=server.createobject("adodb.recordset")
	rs1="select * from tb_user where UserName='"&UserName&"'"
	rs.open rs1,conn,1,3
	if rs.eof and rs.bof then 
	'Ӧ��insert into ���ʵ���û���Ϣ�����
	rs1="insert into tb_user (UserName,Password1,sex,birthday,touxiang,Email) values('"&UserName&"','"&Password1&"','"&sex&"','"&birthday&"','"&touxiang&"','"&Email&"')"
		conn.execute(rs1)
		%>
			<script language="javascript">
				alert("�û�ע��ɹ�����"); //������ʾ�Ի���
				window.location.href="index1.asp";//��ת��ָ����ҳ��
				window.close();
			</script>
<%		end if 
end if
%>