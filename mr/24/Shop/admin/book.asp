<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="chk_admin.asp" -->
<!--#include file="../include/include.asp" -->
<script language = "JavaScript">
function addbook()
{
	if(document.myform.user.value=="") 
	{
		document.myform.user.focus();
		alert("�����뷢���ˣ�");
		return false;
	}
	if(document.myform.shijian.value=="") 
	{
		document.myform.shijian.focus();
		alert("������ʱ�䣡");
		return false;
	}
	if(document.myform.biaoti.value=="") 
	{
		document.myform.biaoti.focus();
		alert("��������⣡");
		return false;
	}
	if(document.myform.neirong.value=="") 
	{
		document.myform.neirong.focus();
		alert("���������ݣ�");
		return false;
	}
}
</script>
<%
if request("action")="del" then
	conn.execute("delete from [liuyan]  where id="&request("id")&"")
	response.Write("<script>alert('ɾ���ɹ���');window.location.href='book.asp';</script>")
end if
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">

  <tr> 
    <td align="center"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
	  <tr>
        <td width="46%">&nbsp;</td>
        <td width="29%"><font color="#FFFFFF">�鿴������Ϣ</font></td>
        <td width="25%">
		</td>
      </tr>
    </table></td>
  </tr>
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="128"><font color="#FF0000">������</font></td>
          <td width="420"><font color="#FF0000">������</font></td>
          <td width="153"><font color="#FF0000">ʱ ��</font></td>
          <td width="85"><font color="#FF0000">��  ��</font></td>
        </tr>
<%
sql="select * from [liuyan] order by id desc;"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,1
if rs.eof And rs.bof then
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
%>
        <tr bgcolor="#FFFFFF" align="center"> 
            <td><%=rs("user")%></td>
            <td><a href="#" onClick='javascript:window.open("lookbook.asp?id=<%=rs("id")%>")'><%=rs("biaoti")%></a></td>
            <td><%=rs("shijian")%></td>
            <td><a href="book.asp?action=del&id=<%=rs("id")%>">ɾ��</a></td>
        </tr>
<%
	rs.movenext
	if rs.eof then exit for
	next
	end sub
%>
      <tr>
	  <form action='' method='get' name='form'>
        <td colspan="6" bgcolor="#FFFFFF">
          <div align="center">
<%	
	if page<>1 then
		response.Write("&nbsp;&nbsp;<a href="&path&"?page=1>��һҳ</a>")
		response.Write("&nbsp;&nbsp;<a href="&path&"?page="&(page-1)&" >��һҳ</a>")
	end if 
	response.Write("&nbsp;&nbsp;��ǰ <font color='#FF0000'>"&page&"</font> ҳ")
	response.Write("&nbsp;&nbsp;�� <font color='#FF0000'>"&rs.recordcount&"</font>/<font color='#FF0000'>"&rs.pagecount&"</font> ҳ")
	if page<>rs.pagecount then
		response.Write("&nbsp;&nbsp;<a href="&path&"?page="&(page+1)&">��һҳ</a>")
		response.Write("&nbsp;&nbsp;<a href="&path&"?page="&rs.pagecount&">��ĩҳ</a>")
	end if
	response.Write("&nbsp;&nbsp;��ת��<input type='text' size='2' name='page'>ҳ<input type='submit' value='GO'>")
end if
rs.close
set rs=nothing
%>
		    </div></td>
        </form></tr>
     </table>
      <br></td>
  </tr>
</table>