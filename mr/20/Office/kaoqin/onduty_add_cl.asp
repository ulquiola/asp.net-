<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!--#include file="../conn/conn.asp"-->
<%
if request.Form("name1")<>"" then
name1=request.Form("name1")
department=request.Form("department")
enroltype=request.Form("enroltype")
enrolremark=request.Form("enrolremark")
ss=now()
'��ȡ�Ǽ�����
select case enroltype
       case "�ϰ�Ǽ�"
	   'Ӧ��cdate�������ַ���ת��Ϊ������������
	   ee=cdate(date()&" 08:20:00")
	   if ee-ss>=0 then
	   	state1="����"
	   else
	   	state1="�ٵ�"
	   end if
		sql="insert into Tab_onduty(name1,department,enroltype,enrolremark,definetime,enroltime,state) values('"&name1&"','"&department&"','"&enroltype&"','"&enrolremark&"','"&ee&"','"&ss&"','"&state1&"')"
		conn.execute(sql)
		case "�°�Ǽ�"
		'Ӧ��cdate�������ַ���ת��Ϊ������������
		ee=cdate(date()&" 17:10:00")
		ee1=cdate(date()&" 20:30:00")
	   if ee-ss<=0 and ee1-ss<=0 then
	   state1="�Ӱ�"
	   else if ee-ss<=0 and ee1-ss>0 then
	   	state1="����"
	   else
	   	state1="����"
	    end if
		end if	   
			sql="insert into Tab_onduty(name1,department,enroltype,enrolremark,definetime,enroltime,state) values('"&name1&"','"&department&"','"&enroltype&"','"&enrolremark&"','"&ee&"','"&ss&"','"&state1&"')"
		conn.execute(sql)
end select
%>
<script language="javascript">
alert("�Ǽ���Ϣ��ӳɹ���");
opener.location.reload();
window.close();
</script>
<%end if%>