<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="include/conn.asp" -->
<!--#include file="include/md5.asp" -->
<%
if request("login")="out" then	'���߽��û���¼���˳�д��ͬһ���ļ��ڣ��Խ��յ��� login ֵ�����ж�
	session("cishu")=""
	session("shijian")=""
	session("user")=""
	response.Redirect("index.asp")	'����������û��йص���Ϣ����ת����ҳ
	response.End()
end if
sql="select * from [user] where name='"&trim(request("user"))&"';"	'���û������в�ѯ
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
if not rs.eof then	'����û�������
	if rs("pass")=md5(trim(request("pass"))) then
	'�����ݿ��д洢��������û��ύ��������бȽ� 
		session("shijian")=rs("shijian2")
		session("cishu")=rs("cishu")
		session("user")=trim(request("user"))	'����ص���Ϣд�� session �ڣ��Ա���ʱ��ȡ
		rs("shijian2")=now()	'���һ�ε�¼ʱ�䣬Ҳ���ǵ�ǰʱ��
		rs("cishu")=rs("cishu")+1	'��¼������1
		rs.update
		response.Redirect("index.asp")	'�ɹ���ת����ҳ
	else	'������벻һ��
		session("user")=""
		session("cishu")=""
		session("shijian")=""
		response.Write("<script>alert('�û������������');window.location.href='index.asp';</script>")
		response.End() '����������û��йص���Ϣ����ת����ҳ
	end if
else	'����û���������
	session("user")=""
	session("cishu")=""
	session("shijian")=""
	response.Write("<script>alert('�û������������');window.location.href='index.asp';</script>")
	response.End() '����������û��йص���Ϣ����ת����ҳ
end if
rs.close
set rs=nothing
%>