<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<style type="text/css">
body{
	margin:12px;
	font-size: 12px;
}

</style>
</head>
<body>
<%
dim ary(7)
ary(0)="��������ã��ü��¼��̣�����N����ã��Ƿ������죬��ˮ~~~~"
ary(1)="�����ѣ�����Ļ���˰ɣ��Ͽ����������԰ɣ�"
ary(2)="�������ߵ����������Զ��������ů�ļң�"
ary(3)="���������ѣ����Ǻã�����������ӭ���ǡ�"
ary(4)="��ɹ��𣿼�����ǻ������ǿ��������"
ary(5)="���ǵ�������������ҵ������һ�룡"
ary(6)="����������ʱ��Ļ�����������ɷ��ɣ�"
Randomize     '��ʼ��������������������������л���ϵͳ��ʱ��������
num = int(rnd() * 10)
select case num  
	case 0
%>
		<%=ary(0) %>
	<% case 1 %>
		<%=ary(1) %>
	<% case 2 %>
		<%=ary(2) %>
	<% case 3 %>
		<%=ary(3) %>
	<% case 4 %>
		<%=ary(4) %>
	<% case 5 %>
		<%=ary(5) %>
	<% case else %>
		<%=ary(6) %>	 
<% end select %>
</body>
</html>
