<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="../include/include.asp" -->
<!--#include file="chk_admin.asp" -->
<script language = "JavaScript">
function addclass()
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
	conn.execute("delete from [class]  where id="&request("id")&"")
	conn.execute("delete from [shangpin]  where classid="&request("id")&"")
	response.Write("<script>alert('ɾ���ɹ���');window.location.href='class.asp?bigclassid="&request("bigclassid")&"';</script>")
end if

if request("action")="update" then
	if request("paixu")="" or request("mingcheng")="" then
		response.Write("<script>alert('����ϸ��д��');history.back();</script>")
		response.End()
	end if
	SafeRequest(trim(request("paixu")))
	sql="select * from [class] where id="&request("id")&""
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs("mingcheng")=trim(request("mingcheng"))
	rs("paixu")=trim(request("paixu"))
	if clng(rs("bigclassid"))<>clng(request("bigclassid")) then
		sql2="select * from [shangpin] where bigclassid="&rs("bigclassid")&";"
		set rs2=Server.CreateObject("ADODB.Recordset")
		rs2.open sql2,conn,3,3
		do while not rs2.eof
			rs2("bigclassid")=request("bigclassid")
			rs2("classid")=rs("id")
			rs2.update
		rs2.movenext
		loop
		rs2.close
		set rs2=nothing
		rs("bigclassid")=request("bigclassid")
	end if
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('�޸ĳɹ���');window.location.href='class.asp?bigclassid="&request("bigclassid")&"';</script>")
end if

if request("action")="add" then
	if request("paixu")="" or request("mingcheng")="" then
		response.Write("<script>alert('����ϸ��д��');history.back();</script>")
		response.End()
	end if
	SafeRequest(trim(request("paixu")))
	sql="select * from [class]"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs.addnew
	rs("mingcheng")=request("mingcheng")
	rs("paixu")=request("paixu")
	rs("bigclassid")=request("bigclassid")
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('��ӳɹ���');window.location.href='class.asp?bigclassid="&request("bigclassid")&"';</script>")
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
            <td width="155"><font color="#FF0000">��������</font></td>
            <td width="155"><font color="#FF0000">�޸�</font></td>
            <td width="150"><font color="#FF0000">����</font></td>
            <td width="160"><font color="#FF0000">�����Ʒ</font></td>
          </tr>
<%
if request("bigclassid")<>"" then	'�ж��Ƿ�ָ������
	bigclassid="where bigclassid="&request("bigclassid")&""	'���Ӳ�ѯ�����SQL��丳ֵ������ bigclassid ��
end if
sql="select * from [class] "&bigclassid&" order by paixu"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
do while not rs.eof	'������������ķ���
%>
<form name="form" action="class.asp" method="post">
            <tr bgcolor="#FFFFFF" align="center">
              <td><input name="mingcheng" type="text" value="<%=rs("mingcheng")%>" size="32"></td>
              <td><input name="paixu" type="text" size="10" value="<%=rs("paixu")%>"></td>
              <td>
<select name="bigclassid">
<%
sql2="select * from [bigclass] where id="&rs("bigclassid")&" order by paixu"
set rs2=Server.CreateObject("ADODB.Recordset")
rs2.open sql2,conn,3,3	'����÷��������Ĵ���
%>
  <option value="<%=rs2("id")%>"><%=rs2("mingcheng")%></option>
<%
rs2.close
set rs2=nothing
%>
<%
sql3="select * from [bigclass] where id<>"&rs("bigclassid")&" order by paixu"
set rs3=Server.CreateObject("ADODB.Recordset")
rs3.open sql3,conn,3,3
do while not rs3.eof	'������˸÷������������д���
%>
  <option value="<%=rs3("id")%>"><%=rs3("mingcheng")%></option>
<%
rs3.movenext
loop
rs3.close
set rs3=nothing
%>
</select>

			  </td>
              <td><input name="submit" type="submit" value="�޸�"></td>
              <td><a href="class.asp?action=del&id=<%=rs("id")%>&bigclassid=<%=rs("bigclassid")%>">ɾ��</a></td>
              <td><a href="addpro.asp?bigclassid=<%=rs("bigclassid")%>&classid=<%=rs("id")%>">�����Ʒ</a></td>
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
          <form name="myform" action="class.asp" method="post" onSubmit="return addclass();">
            <tr height="20" bgcolor="#FFFFFF" align="center">
              <td width="285"><font color="#FF0000">��</font><font color="#FF0000">������&nbsp;&nbsp;
                    <input name="mingcheng" type="text" value="" size="26">
              </font></td>
              <td width="198"><font color="#FF0000">��������&nbsp;&nbsp;
                    <input name="paixu" type="text" size="10" value="">
              </font></td>
              <td width="304">
  <font color="#FF0000">��������;</font>&nbsp;&nbsp;
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
</select>			  </td>
              <td width="87"><input name="submit2" type="submit" value="���"></td>
            </tr>
            <input type="hidden" name="action" value="add">
          </form>
        </table>
        <br></td>
  </tr>
</table>