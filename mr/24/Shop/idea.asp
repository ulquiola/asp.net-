<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.style1 {color: #f2ab5b}
-->
</style>
<!--#include file="include/conn.asp" -->
<!--#include file="chk_user.asp" -->
<!--#include file="include/include.asp" -->
<script language = "JavaScript">
function adddismess()
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
if request("action")="add" then
	if trim(request("user"))="" or trim(request("shijian"))="" or trim(request("biaoti"))="" or trim(request("neirong"))="" then
		response.Write("<script>alert('����ϸ��д��');history.back();</script>")
		response.End()
	end if
	sql="select * from [fankui]"
	set rs=Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,3,3
	rs.addnew
	rs("user")=trim(request("user"))
	rs("shijian")=trim(request("shijian"))
	rs("biaoti")=trim(request("biaoti"))
	rs("neirong")=trim(request("neirong"))
	rs("leixing")=trim(request("leixing"))
	rs.update
	rs.close
	set rs=nothing
	response.Write("<script>alert('лл������ı��������');window.location.href='index.asp';</script>")
end if
%>
<table width="792" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="3"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="792" height="165" background="images/index_r1_c1.jpg"><!--#include file="top.asp" --></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td width="197"><!--#include file="left.asp" --></td>
    <td width="590" valign="top"><table border="0" cellpadding="0" cellspacing="0" width="590">
      <tr>
        <td colspan="3"><img name="index_7_r1_c1" src="images/yjfk.jpg" width="590" height="49" border="0" alt=""></td>
      </tr>
      <tr>
        <td colspan="3"><img name="index_7_r2_c1" src="images/index_7_r2_c1.jpg" width="590" height="9" border="0" alt=""></td>
      </tr>
      <tr>
        <td width="16" height="643" background="images/index_7_r3_c1.jpg">&nbsp;</td>
        <td width="564" valign="top">
          <table width="100%"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="30">&nbsp;</td>
            </tr>
          </table>
          <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1">
<form name="myform" action="idea.asp" method="post" onSubmit="return adddismess();">
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="25%" height="30">�����ˣ�</td>
          <td width="75%"><div align="left">
            <input name="user" type="text" value="<%=session("user")%>">&nbsp;&nbsp;&nbsp;&nbsp;
            <select name="leixing">
              <option value="1">����վ�Ľ���</option>
              <option value="2">�Թ�˾�Ľ���</option>
              <option value="3">�Բ�Ʒ��Ͷ��</option>
              <option value="4">�Է����Ͷ��</option>
            </select>
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">ʱ&nbsp;&nbsp;�䣺</td>
          <td><div align="left">
            <input name="shijian" type="text" value="<%=now()%>">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="30">��������</td>
          <td><div align="left">
            <input name="biaoti" type="text" size="40">
          </div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>��&nbsp;&nbsp;�ݣ�</td>
          <td><div align="left"><textarea name="neirong" cols="60" rows="16"></textarea>
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
      </table>		</td>
        <td width="10" background="images/index_7_r3_c3.jpg">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="3"><img name="index_7_r4_c1" src="images/index_7_r4_c1.jpg" width="590" height="7" border="0" alt=""></td>
      </tr>
    </table></td>
    <td width="8"><img name="index_r2_c3" src="images/index_r2_c3.jpg" width="5" height="753" border="0" alt=""></td>
  </tr>
  <tr>
    <td colspan="3"><!--#include file="foot.asp" --></td>
  </tr>
</table>