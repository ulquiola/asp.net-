<%
  Dim Conn,Str											'�������
  Set Conn=Server.CreateObject("ADODB.Connection") 		'����Connection����ʵ��
  Str="Driver={Microsoft Access Driver (*.mdb)};DBQ="&Server.mappath("DataBase/db_09.mdb")&"" 
'�����������ݿ��ַ���
  Conn.Open(Str) 										'��������
%>