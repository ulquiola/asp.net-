<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
response.ContentType="text/html; charset=gb2312"			'����ת��
Dim conn													'��������
set conn=Server.CreateObject("ADODB.Connection")			'��������Դ
sql="provider=sqloledb;data source=(local);initial catalog=db_QQ;user id=sa;password=;"		'�������ݿ�����
conn.open(sql)												'ִ��SQL���
if request("UserName")<>"" then								'�ж��û����Ƿ�Ϊ��
	session("UserName")=request("UserName")					'Ϊsession("UserName")������ֵ
end if
Set rs=Server.CreateObject("ADODB.RecordSet")				'������¼��
sql1="select * from tb_user where UserName='"&session("UserName")&"'"	'��ѯ����
conn.execute(sql1)											'ִ��sql1���
rs.open sql1,conn,1,3										'�򿪼�¼��
if rs.eof and rs.bof then							
	response.Write("ף�������û���["&session("UserName")&"]û�б�ע�ᣡ")
	'������ʾ��Ϣ
else
	response.Write("�ܱ�Ǹ���û���["&session("UserName")&"]�Ѿ���ע�ᣡ")
	'������ʾ��Ϣ
end if 
%>
