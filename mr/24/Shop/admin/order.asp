<!--#include file="../include/conn.asp" -->
<!--#include file="chk_admin.asp" -->
<!--#include file="../include/include.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�����û�����</title>
</head>
<body>
<%
if request("action")="del" then	'�ж��Ƿ�ִ������ɾ��
	conn.execute("delete from [dingdan]  where id="&request("id")&"")	'����ȡ�õ���Ʒ ID ����ɾ��
	response.Write("<script>alert('ɾ���ɹ���');window.location.href='order.asp';</script>")
end if
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
  <tr> 
    <td align="center"><font color="#FFFFFF">�����û�����</font></td>
  </tr>
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="151"><font color="#FF0000">�������</font></td>
          <td width="147"><font color="#FF0000">�ջ���ʽ</font></td>
          <td width="168"><font color="#FF0000">���ʽ</font></td>
          <td width="156"><font color="#FF0000">�µ��û�</font></td>
          <td width="158"><font color="#FF0000">�µ�ʱ��</font></td>
          <td width="88"><font color="#FF0000">�� ��</font></td>
        </tr>
<%
if request("tiaojian")="1" then
	chaxun="where didanhao='"&request("guanjian")&"'"	'��������Ų�ѯ
end if
if request("tiaojian")="2" then
	chaxun="where name='"&request("guanjian")&"'"	'���û�����ѯ
end if
if request("tiaojian")="3" then
	if request("guanjian")="" then	'�Ƿ���Ҫ��������ѯ
		chaxun="where zhuangtai='1'"	'���Ѹ���Ķ�����ѯ
	else
		chaxun="where zhuangtai='1' and  name='"&request("guanjian")&"'"	'���Ѹ���Ķ������Ҷ�Ӧ�û�����ѯ
	end if
end if
if request("tiaojian")="4" then
	if request("guanjian")="" then	'�Ƿ���Ҫ��������ѯ
		chaxun="where zhuangtai='2'"	'���ѷ����Ķ�����ѯ
	else
		chaxun="where zhuangtai='2' and  name='"&request("guanjian")&"'"	'���ѷ����Ķ������Ҷ�Ӧ�û�����ѯ
	end if
end if
sql="select * from [dingdan] "&chaxun&" order by id desc;"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.eof And rs.bof then	'�ж����û�����ݵĻ�
	Response.Write "<p align='center' class='contents'> �Բ����������ݣ�</p>"
else
	rs.pagesize=10	'ÿҳ��ʾ������
	SafeRequest(request("page"))
	page=clng(request("page"))
	if page<1 then page=1	'�����ǰҳ��С�� 1 ��ҳ������ 1������ͬ��
	if page>rs.pagecount then page=rs.pagecount	'�����ǰҳ��������ҳ������ǰҳ��������ҳ����
	'����һ����10ҳ������ǰҳ����100����ʱ��ǰҳ��������ҳ������û����Ϣ���ߵ��³������
	show rs,page
	sub show(rs,page)	'�������
	rs.absolutepage=page
	for i=1 to rs.pagesize
%>
        <tr bgcolor="#FFFFFF" align="center"> 
            <td><a href="lookorder.asp?dingdan=<%=rs("didanhao")%>"><%=rs("didanhao")%></a></td>
            <td><%=rs("songhuo")%></td>
            <td><%=rs("zhifu")%></td>
            <td><%=rs("name")%></td>
            <td><%=rs("shijian")%></td>
            <td><a href="order.asp?action=del&id=<%=rs("id")%>">ɾ��</a></td>
        </tr>
<%
	rs.movenext
	if rs.eof then exit for	'������ݵ������һ������������ѭ��
	next
	end sub	'���̽���
%>      
        <tr bgcolor="#FFFFFF" align="center">
		<form name="form" action="?" method="get">		
          <td colspan="6">
<%	
	if page<>1 then
		response.Write("&nbsp;&nbsp;<a href="&path&"?page=1>��һҳ</a>")
		response.Write("&nbsp;&nbsp;<a href="&path&"?page="&(page-1)&" >��һҳ</a>")
	end if 
	response.Write("&nbsp;&nbsp;��ǰ <font color='#FF0000'>"&page&"</font> ҳ/<font color='#FF0000'>"&rs.pagecount&"</font> ҳ")
	response.Write("&nbsp;&nbsp;�� <font color='#FF0000'>"&rs.recordcount&"</font>&nbsp;��")
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
<br>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
<form action="order.asp" method="post">
  <tr>
    <td align="center"><font color="#FFFFFF">�û�������ѯ</font></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><div align="center">��ѯ����&nbsp;&nbsp;
        <input name="guanjian" type="text" size="30">
        &nbsp;&nbsp;
	      <select name="tiaojian">
            <option value="1">������Ų�ѯ</option>
            <option value="2">�û�������ѯ</option>
            <option value="3">�������ѯ</option>
            <option value="4">����������ѯ</option>
          </select>
    &nbsp;&nbsp;<input type="submit" value="��ѯ"></div></td>
  </tr>
  <input type="hidden" name="chaxun" value="chaxun">
 </form>
</table>
</body>
</html>