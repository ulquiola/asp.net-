<%
dim conn,db
dim connstr
db="./dataBase/db_jmail.mdb" '数据库文件位置
connstr="DBQ="+server.mappath(""&db&"")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set conn=server.createobject("ADODB.CONNECTION")
conn.open connstr
sub CloseConn()
	conn.close
	set conn=nothing
end sub
%>