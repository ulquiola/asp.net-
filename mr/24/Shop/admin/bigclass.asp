<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="../include/include.asp" -->
<!--#include file="chk_admin.asp" -->
<script language = "JavaScript">
function addbigclass()
{
	if(document.myform.mingcheng.value=="") 
	{
		document.myform.mingcheng.focus();
		alert("������������ƣ�");
		return false;
	}
	if(document.myform.paixu.value=="") 
	{
		document.myform.paixu.focus();
		alert("�������������");
		return false;
	}
}
</script>
<%
if request("action")="del" then
	conn.execute("delete from [bigclass]  where id="&request("id")&"")	'ɾ������
	conn.execute("delete from [class]  where bigclassid="&request("id")&"")	'ɾ����Ӧ����ķ���
	conn.execute("delete from [shangpin]  where bigclassid="&request("id")&"")	'ɾ����Ӧ�������Ʒ
	response.Write("<script>alert('ɾ���ɹ���');window.location.href='bigclass.asp';</script>")
end if

if request("action")="update" then
	if request("paixu")="" or request("mingcheng")="" then
		response.Write("<script>alert('����ϸ��д��');history.back();</script>")
		response.End()
	end if
	SafeRequest(trim(request("paixu")))
	sql="select * from [bigclass] where id="&request("id")&""
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs("mingcheng")=trim(request("mingcheng"))
	rs("paixu")=trim(request("paixu"))
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('�޸ĳɹ���');window.location.href='bigclass.asp';</script>")
end if

if request("action")="add" then
	if request("paixu")="" or request("mingcheng")="" then
		response.Write("<script>alert('����ϸ��д��');history.back();</script>")
		response.End()
	end if
	SafeRequest(trim(request("paixu")))
	sql="select * from [bigclass]"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs.addnew
	rs("mingcheng")=trim(request("mingcheng"))
	rs("paixu")=request("paixu")
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('��ӳɹ���');window.location.href='bigclass.asp';</script>")
end if
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
  <tr> 
    <td align="center"><font color="#FFFFFF">������Ʒ����</font></td>
  </tr>
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
		<tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="254"><font color="#FF0000">��������</font></td>
          <td width="152"><font color="#FF0000">��������</font></td>
          <td width="155"><font color="#FF0000">�޸�</font></td>
          <td width="150"><font color="#FF0000">����</font></td>
          <td width="160"><font color="#FF0000">�������</font></td>
        </tr>
<%
sql="select * from [bigclass] order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof
%>
        <form name="form" action="bigclass.asp" method="post">
        <tr bgcolor="#FFFFFF" align="center"> 
            <td><input name="mingcheng" type="text" value="<%=rs("mingcheng")%>" size="32"></td>
            <td><input name="paixu" type="text" size="10" value="<%=rs("paixu")%>"></td>
            <td><input name="submit" type="submit" value="�޸�"></td>
            <td><a href="bigclass.asp?action=del&id=<%=rs("id")%>">ɾ��</a></td>
            <td><a href="class.asp?bigclassid=<%=rs("id")%>">�������</a></td>
        </tr>
		<input name="id" type="hidden" value="<%=rs("id")%>">
		<input name="action" type="hidden" value="update">
       </form> 
<%
rs.movenext
loop
%>
	</table>
      <br></td>
  </tr>
</table>
<br>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
  <tr>
    <td align="center"><font color="#FFFFFF">�����Ʒ����</font></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><br>
        <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
        <form name="myform" action="bigclass.asp" method="post" onSubmit="return addbigclass();">
          <tr height="20" bgcolor="#FFFFFF" align="center">
            <td width="425"><font color="#FF0000">��������&nbsp;&nbsp;              
              <input name="mingcheng" type="text" value="" size="32">
            </font></td>
            <td width="287"><font color="#FF0000">��������&nbsp;&nbsp;              
              <input name="paixu" type="text" size="10" value="">
            </font></td>
            <td width="165"><input name="submit" type="submit" value="���"></td>
          </tr>
		  <input type="hidden" name="action" value="add">
		  </form>
        </table>
        <br></td>
  </tr>
</table>
