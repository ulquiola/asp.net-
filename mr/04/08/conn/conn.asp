<%
  Dim Conn,Str											'定义变量
  Set Conn=Server.CreateObject("ADODB.Connection") 		'创建Connection对象实例
  Str="Driver={Microsoft Access Driver (*.mdb)};DBQ="&Server.mappath("DataBase/db_09.mdb")&"" 
'定义连接数据库字符串
  Conn.Open(Str) 										'建立连接
%>