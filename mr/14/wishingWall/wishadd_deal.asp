<!--#include file="conn.asp" -->
<%	response.ContentType="text/html; charset=GBK"
	wishMan=request("wishMan")
	wellWisher=request("wellWisher")
	content=request("content")
	color=request("color")
	face=request("face")
	sql="Select * From tb_wish"
	Set rs=server.CreateObject("Adodb.RecordSet")
	rs.open sql,conn,1,3	
	rs.addnew
	rs("wishMan")=wishMan
	rs("wellWisher")=wellWisher
	rs("content")=content
	rs("color")=color
	rs("face")=face
	rs.update
	wishid=rs("id")
	rs.close
	Set rs = Nothing
	conn.close		
	Set Conn = Nothing
	response.write "恭喜您，贴字条成功，请记住您的字条编号：["&wishid&"]"
%>
