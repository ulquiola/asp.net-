<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="chk_admin.asp" -->
<!--#include file="../include/include.asp" -->
<script language = "JavaScript">
function addnotify()
{
	if(document.myform.shijian.value=="") 
	{
		document.myform.shijian.focus();
		alert("������ʱ�䣡");
		return false;
	}
	if(document.myform.neirong.value=="") 
	{
		document.myform.neirong.focus();
		alert("���乫�����ݣ�");
		return false;
	}
}
</script>
<%
if request("action")="update" then
	if trim(request("shijian"))="" or trim(request("neirong"))="" then
		response.Write("<script>alert('����ϸ��д��');history.back();</script>")
		response.End()
	end if
	sql="select * from [gonggao]"	'��δָ����������ʱ����Ĭ�϶Ա��ڵ��������ݽ��в���
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs("shijian")=trim(request("shijian"))
	rs("neirong")=trim(request("neirong"))
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('�������óɹ���');window.location.href='notify.asp';</script>")
end if
%>
<br>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
  <tr>
    <td align="center"><font color="#FFFFFF">վ�ڹ�������</font></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><br>
        <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
<%
sql="select * from [gonggao]"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
if not rs.eof then
%>
          <form name="myform" action="notify.asp" method="post" onSubmit="return addnotify();">
            <tr height="20" bgcolor="#FFFFFF" align="center">
              <td width="25%">ʱ&nbsp;&nbsp;�䣺</td>
              <td width="75%"><div align="left">
                  <input name="shijian" type="text" value="<%=now()%>">
              </div></td>
            </tr>
            <tr height="20" bgcolor="#FFFFFF" align="center">
              <td>��&nbsp;&nbsp;�ݣ�</td>
              <td><div align="left">
                  <textarea name="neirong" cols="60" rows="8"><%=rs("neirong")%></textarea>
              </div></td>
            </tr>
            <tr height="20" bgcolor="#FFFFFF" align="center">
              <td colspan="2">
                <input name="action" type="hidden" value="update">
                <input name="submit" type="submit" value="����">
&nbsp;&nbsp;
        <input type="reset" name="reset" value="��д">
&nbsp;&nbsp;</td>
            </tr>
          </form>
<%
end if
rs.close
set rs=nothing
%>
        </table>
        <br></td>
  </tr>
</table>