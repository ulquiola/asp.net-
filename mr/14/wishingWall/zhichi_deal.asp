<!--#include file="conn.asp" -->
<%	
	response.ContentType="text/html; charset=GBK"
	id=request("scripId")
	Set rs=server.CreateObject("Adodb.RecordSet")
	UP="Update tb_wish set hits=hits+1 where ID="&ID
	conn.execute(UP)
	sql="Select * From tb_wish where id="&id
	rs.open sql,conn,1,3
	id1=rs("id")
	hits=rs("hits")
	idvalue=id1&"#"&hits
	response.Write(idvalue)
	rs.close
	Set rs = Nothing
	conn.close		
	Set Conn = Nothing
%>
