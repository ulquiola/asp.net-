<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="chk_admin.asp" -->
<!--#include file="../include/include.asp" -->
<script language = "JavaScript">
function addnews()
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
		alert("��������ʱ�䣡");
		return false;
	}
	if(document.myform.biaoti.value=="") 
	{
		document.myform.biaoti.focus();
		alert("�������ű��⣡");
		return false;
	}
	if(document.myform.neirong.value=="") 
	{
		document.myform.neirong.focus();
		alert("�����������ݣ�");
		return false;
	}
}
</script>
<%
if request("action")="add" then
	if trim(request("user"))="" or trim(request("shijian"))="" or trim(request("biaoti"))="" or trim(request("neirong"))="" then
		response.Write("<script>alert('����ϸ��д��');history.back();</script>")
		response.End()
	end if
	sql="select * from [news]"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs.addnew
	rs("user")=trim(request("user"))
	rs("shijian")=trim(request("shijian"))
	rs("biaoti")=trim(request("biaoti"))
	rs("neirong")=trim(request("neirong"))
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('վ��������ӳɹ���');window.location.href='addnews.asp';</script>")
end if
%>
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
 <tr> 
    <td align="center"><font color="#FFFFFF">���վ������</font></td>
  </tr>
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
<form name="myform" action="addnews.asp" method="post" onSubmit="return addnews();">
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="25%">�����ˣ�</td>
          <td width="75%"><div align="left">
            <input name="user" type="text" value="<%=session("admin")%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>ʱ&nbsp;&nbsp;�䣺</td>
          <td><div align="left">
            <input name="shijian" type="text" value="<%=now()%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>��������</td>
          <td><div align="left">
            <input name="biaoti" type="text" size="40">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>��&nbsp;&nbsp;�ݣ�</td>
          <td><div align="left"><textarea name="neirong" cols="80" rows="18"></textarea>
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