<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../conn/conn.asp"-->
<%
set rs=server.CreateObject("adodb.recordset")
sql="select * from tb_board order by id desc"
rs.open sql,conn,1,3
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Css/style.css" rel="stylesheet">
<title>�ޱ����ĵ�</title>
</head>
<body>
<div style=" height: 50px;"></div>
<%
If(rs.Eof)Then
	Response.Write("����������Ϣ!")
	response.End()
Else
%>
<table width="95%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" bordercolordark="#1985DB" bordercolorlight="#FFFFFF">
	<tr>
		<td height="25" align="center" nowrap="nowrap">���</td>
		<td height="25" align="center" nowrap="nowrap">������</td>
		<td height="25" align="center" nowrap="nowrap">ԤԼ��</td>
		<td height="25" align="center" nowrap="nowrap">��ѯ����</td>
		<td height="25" align="center" nowrap="nowrap">��ѯ����</td>
		<td height="25" colspan="2" align="center" nowrap="nowrap">����</td>
	</tr>
<%
'��ҳ
rs.pagesize=8
page1=clng(request("page1"))
if page1<1 then page1=1
rs.absolutepage=page1
for i=1 to rs.pagesize
	content=rs("content")
%>
	<tr>
		<td align="center"><%=rs("ID")%>&nbsp;</td>
		<td align="center"><%=rs("person")%>&nbsp; </td>
		<td align="center"><%=rs("yuyueren")%></td>
		<td align="center">
<%
	if len(rs("title"))>5 then
		response.Write(replace(left(rs("title"),5)&"��",chr(13),"<br>"))
	else
		response.Write(rs("title"))
	end if
%>
		</td>
		<td align="center">
			<% if len(content)>10 then
				response.Write(replace(left(content,10)&"....",chr(13),"<br>"))
			else
				response.Write(content)
			end if %>
		</td>
		<td align="center"><a href="#" onClick="JScript:window.open('liuyan_chakan.asp?id=<%=rs("id")%>','','width=500,height=400')">�鿴</a></td>
<td align="center"><a href="#" onClick="if(confirm('�Ƿ�ȷ��ɾ��?')){window.location.href='liuyan_del.asp?id=<%=rs("id")%>';}">ɾ��</a></td>
	</tr>
<%
rs.movenext
if rs.eof then exit for
next
%>
</table>
<table width="95%" height="25" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr><td align="right" valign="middle">
[<%=page1%>/<%=rs.pagecount%>]&nbsp;[��<font color="red"><%=rs.recordcount%></font>����¼]
<% if page1<>1 then%>
<a href="<%=path1%>?page1=1">��һҳ</a> <a href="<%=path1%>?page1=<%=(page1-1)%>"> ��һҳ</a>
<%end if %>
<%if page1<>rs.pagecount then%>
<a href="<%=path1%>?page1=<%=(page1+1)%>">��һҳ</a> <a href="<%=path1%>?page1=<%=rs.pagecount%>" >���һҳ</a>
<%end if%>
		</td></tr>
</table>
<%end if%>
</body>
</html>
