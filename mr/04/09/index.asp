<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
  Dim Conn,Str									'�������
  Set Conn=Server.CreateObject("ADODB.Connection") 		'����Connection����ʵ��
  Str="Driver={Microsoft Access Driver (*.mdb)};DBQ="&Server.mappath("DataBase/db_8.mdb")&"" 
'�����������ݿ��ַ���
  Conn.Open(Str) 									'��������
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��ȡ������µ��ֶ�</title>
<style>
body{
	margin:12px;
}
body,th,td{
	font-size:12px;
}
</style>
</head>
<body>
<%
set rs=server.CreateObject("adodb.recordset")				'����recordset����ʵ��
sql="select * from tb_user"							'��������
rs.open sql,conn,1,3									'�򿪼�¼��
%>
<table width="437" height="56" border="1" align="center" cellpadding="0" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#CCCCCC">
  <tr>
    <td bgcolor="#FFCC33"><div align="center">�û�����</div></td>
    <td bgcolor="#FFCC33"><div align="center">�û�����</div></td>
    <td bgcolor="#FFCC33"><div align="center">ͨѶ��ַ</div></td>
    <td bgcolor="#FFCC33"><div align="center">��������</div></td>
  </tr>
  <tr>
    <td><div align="center"><%=Rs.Fields("uname").value%></div></td><!--��ȡ�ֶ���Ϣ-->
    <td><div align="center"><%=Rs.Fields("pwd").value%></div></td><!--��ȡ�ֶ���Ϣ-->
    <td><div align="center"><%=Rs.Fields("address").value%></div></td><!--��ȡ�ֶ���Ϣ-->
    <td><div align="center"><%=Rs.Fields("postcode").value%></div></td><!--��ȡ�ֶ���Ϣ-->
  </tr>
</table>

</body>
</html>
