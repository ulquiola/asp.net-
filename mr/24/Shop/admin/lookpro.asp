<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="../include/include.asp" -->
<!--#include file="chk_admin.asp" -->
<%
if request("action")="del" then
	conn.execute("delete from [shangpin]  where id="&request("id")&"")
	response.Write("<script>alert('ɾ���ɹ���');window.location.href='lookpro.asp';</script>")
end if
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
	<tr> 
    <td align="center"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
<form name="myform" action="lookpro.asp" method="post">  <tr> 
        <tr>
          <td width="320">&nbsp;</td>
          <td width="120"><div align="center"><font color="#FFFFFF">��Ʒ��Ϣ����</font> </div></td>
          <td width="160">
<font color="#FFFFFF">�������ࣺ</font>
<select name="bigclassid">
<%
sql="select * from [bigclass] order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof
%>
    <option value="<%=rs("id")%>"><%=rs("mingcheng")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
</select></td>
          <td width="160">
<font color="#FFFFFF">����С�ࣺ</font>
<select name="classid">
<%
sql="select * from [class] order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof
%>
    <option value="<%=rs("id")%>"><%=rs("mingcheng")%></option>
<%
rs.movenext
loop
rs.close
set rs=nothing
%>
</select></td>
          <td width="40"><input type="submit" value="�鿴"></td>
        </tr>
	</form>
      </table></td>
  </tr>
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="234"><font color="#FF0000">��Ʒ����</font></td>
          <td width="167"><div align="center"><font color="#FF0000">�ϼ�ʱ��</font></div></td>
          <td width="139"><font color="#FF0000">�г��۸�</font></td>
          <td width="132"><font color="#FF0000">��Ա�۸�</font></td>
          <td width="100"><font color="#FF0000">�޸�</font></td>
          <td width="96"><font color="#FF0000">����</font></td>
        </tr>
<%
if request("bigclassid")<>"" and request("classid")<>"" then	'�ж��Ƿ�ָ�������в�ѯ
	wh="where bigclassid="&request("bigclassid")&" and classid="&request("classid")&""
end if
sql="select * from [shangpin] "&wh&" order by id desc"
'���������õ��� wh �����������ȷ���������в�ѯ��ʵ�ʵ� SQL ��������У�����ֵΪ�գ���Ӱ���κβ���,�������Ա������β�ѯ���������鷳
'sql="select * from [shangpin] where bigclassid="&request("bigclassid")&" and classid="&request("classid")&" order by id desc" 
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	Response.Write "<p align='center' class='contents'> �Բ����������ݣ�</p>"
else
	rs.pagesize=10
	SafeRequest(request("page"))
	page=clng(request("page"))
	if page<1 then page=1
	if page>rs.pagecount then page=rs.pagecount
	show rs,page
	sub show(rs,page)
	rs.absolutepage=page
	for i=1 to rs.pagesize
		if rs("shuliang")<20 then
			response.Write("<script>alert('��Ʒ "&rs("mingcheng")&" �����Ѿ����� 20��');</script>")	'�ж�ÿ����Ʒ�Ŀ�����Ƿ���20�����������ʾ
		end if
%>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td><a href="pro.asp?id=<%=rs("id")%>"><%=rs("mingcheng")%></a></td>
          <td><%=rs("riqi")%></td>
          <td><%=rs("shichang")%></td>
          <td><%=rs("huiyuan")%></td>
          <td><a href="uppro.asp?id=<%=rs("id")%>">�޸�</a></td>
          <td><a href="lookpro.asp?id=<%=rs("id")%>&action=del&bigclassid=<%=request("bigclassid")%>&classid=<%=request("classid")%>">ɾ��</a></td>
        </tr>
<%
	rs.movenext
	if rs.eof then exit for
	next
	end sub
%>
        <tr height="20" bgcolor="#FFFFFF" align="center">
		<form name="form" action="?" method="get">
          <td colspan="6">
<%	
	if page<>1 then
		response.Write("&nbsp;&nbsp;<a href="&path&"?page=1>��һҳ</a>")
		response.Write("&nbsp;&nbsp;<a href="&path&"?page="&(page-1)&" >��һҳ</a>")
	end if 
	response.Write("&nbsp;&nbsp;��ǰ <font color='#FF0000'>"&page&"</font> ҳ/<font color='#FF0000'>"&rs.pagecount&"</font> ҳ")
	response.Write("&nbsp;&nbsp;�� <font color='#FF0000'>"&rs.recordcount&"</font>��")
	if page<>rs.pagecount then
		response.Write("&nbsp;&nbsp;<a href="&path&"?page="&(page+1)&">��һҳ</a>")
		response.Write("&nbsp;&nbsp;<a href="&path&"?page="&rs.pagecount&">��ĩҳ</a>")
	end if
	response.Write("&nbsp;&nbsp;��ת��<input type='text' size='2' name='page'>ҳ<input type='submit' value='GO'>")
end if
rs.close
set rs=nothing
%>
   		  </td>
		  </form>
        </tr>
      </table>
		<br></td>
  </tr>
</table>