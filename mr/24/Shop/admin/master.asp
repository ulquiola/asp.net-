<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="chk_admin.asp" -->
<!--#include file="../include/include.asp" -->
<script language = "JavaScript">
function addadmin()
{
	if(document.myform.user.value=="")
	{
		document.myform.user.focus();
		alert("�������û�����");
		return false;
	}
	if(document.myform.pass.value=="")
	{
		document.myform.pass.focus();
		alert("���������룡");
		return false;
	}
	if(document.myform.mail.value=="")
	{
		document.myform.mail.focus();
		alert("����������ʼ���");
		return false;
	}
	if(document.myform.xingming.value=="")
	{
		document.myform.xingming.focus();
		alert("������������");
		return false;
	}
	if(document.myform.tel.value=="")
	{
		document.myform.tel.focus();
		alert("������绰��");
		return false;
	}
	if(document.myform.dizhi.value=="")
	{
		document.myform.dizhi.focus();
		alert("�������ַ��");
		return false;
	}
}
</script>
<%
if request("action")="add" then
	if trim(request("user"))="" or trim(request("pass"))="" or trim(request("mail"))="" or trim(request("xingming"))="" or trim(request("tel"))="" or trim(request("dizhi"))="" then
		response.Write("<script>alert('����ϸ��д��');history.back();</script>")
		response.End()
	end if
	sql="select * from [admin]"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs.addnew
	rs("username")=trim(request("user"))
	rs("pass")=md5(trim(request("pass")))
	rs("mail")=trim(request("mail"))
	rs("xingming")=trim(request("xingming"))
	rs("tel")=trim(request("tel"))
	rs("dizhi")=trim(request("dizhi"))
	rs("vip")="0"
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('��ӳɹ���');window.location.href='master.asp';</script>")
end if
%>
<%
if request("action")="del" then
	conn.execute("delete from [admin]  where id="&request("id")&"")
	response.Write("<script>alert('ɾ���ɹ���');window.location.href='master.asp';</script>")
end if
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
  <tr> 
    <td align="center"><font color="#FFFFFF">�鿴��̨�û�</font></td>
  </tr>
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="133"><font color="#FF0000">�û�I D</font></td>
          <td width="122"><font color="#FF0000">����</font></td>
          <td width="163"><font color="#FF0000">�����ʼ�</font></td>
          <td width="149"><font color="#FF0000">�绰</font></td>
          <td width="233"><font color="#FF0000">��ַ</font></td>
          <td width="68"><font color="#FF0000">�� ��</font></td>
        </tr>
<%
sql="select * from [admin] order by id desc;"
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
            <td><a href="#" onClick='javascript:window.open("lookadmin.asp?id=<%=rs("id")%>");'><%=rs("username")%></a></td>
            <td><%=rs("xingming")%></td>
            <td><%=rs("mail")%></td>
            <td><%=rs("tel")%></td>
            <td><%=rs("dizhi")%></td>
            <td><a href="master.asp?action=del&id=<%=rs("id")%>">ɾ��</a></td>
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
	response.Write("&nbsp;&nbsp;��ǰ <font color='#FF0000'>"&page&"</font> ҳ/<font color='#FF0000'>"&rs.pagecount&"</font> ҳ")
	response.Write("&nbsp;&nbsp;�� <font color='#FF0000'>"&rs.recordcount&"</font> ��")
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

<br>


<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
 <tr> 
    <td align="center"><font color="#FFFFFF">��Ӻ�̨�û�</font></td>
  </tr>
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
<form name="myform" action="master.asp" method="post" onSubmit="return addadmin();">
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="25%">�û�����</td>
          <td width="75%"><div align="left">
            <input name="user" type="text">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>��&nbsp;&nbsp;�룺</td>
          <td><div align="left">
            <input name="pass" type="password" size="21">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>E-MAIL��</td>
          <td><div align="left">
            <input name="mail" type="text">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>��&nbsp;&nbsp;����</td>
          <td><div align="left">
            <input name="xingming" type="text">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>��&nbsp;&nbsp;����</td>
          <td><div align="left">
            <input name="tel" type="text">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>��&nbsp;&nbsp;ַ��</td>
          <td><div align="left">
            <input name="dizhi" type="text" size="40">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td colspan="2">
<input name="action" type="hidden" value="add">
<input name="submit" type="submit" value="���">
&nbsp;&nbsp;<input type="reset" name="reset" value="��д">
&nbsp;&nbsp;</td>
          </tr>
</form>
      </table>
      <br></td>
  </tr>
</table>