<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="top.asp"--><%'������վ��ͷ�ļ�%>
<!--#include file="conn/conn.asp"--><%'�������ݿ������ļ�%>
<%
if request.Form("username")<>"" then	'�ж��û����Ƿ�Ϊ��
Set rs= Server.CreateObject("ADODB.Recordset")	'������¼��
	sql="select username from tb_user where username='"&Request.Form("username")&"'"	'��ѯ����
	rs.open sql,conn,1,3																'�򿪼�¼��
	if rs.eof then 																				
	username=request.Form("username")										'��ȡ��Ԫ��username��ֵ������username����
	name1=request.Form("name1")												'��ȡ��Ԫ��name1��ֵ������name1����
	pwd=request.Form("pwd")													'��ȡ��Ԫ��pwd��ֵ������pwd����
	city=request.form("city")												'��ȡ��Ԫ��city��ֵ������city����
	address=request.Form("address")											'��ȡ��Ԫ��address��ֵ������address����
	postcode=request.Form("postcode")										'��ȡ��Ԫ��postcode��ֵ������postcode����
	telephone=request.Form("telephone")										'��ȡ��Ԫ��telephone��ֵ������telephone����
	email=request.Form("email")												'��ȡ��Ԫ��email��ֵ������email����
	sql="insert into tb_user(username,name1,pwd,city,address,postcode,telephone,email) values('"&username&"','"&name1&"','"&pwd&"','"&city&"','"&address&"','"&postcode&"','"&telephone&"','"&email&"')"	'��Ӽ�¼��Ϣ
	conn.execute(sql)														'ִ��SQL���
%>
	<script language="javascript">
	alert("�û���Ϣ��ӳɹ�!!");											//������ʾ��Ϣ
	window.location.href='index.asp';										//��ת��ָ����ҳ��
	</script>
<%else%>
	<script language="javascript">
	alert("���û��Ѿ����ڣ�");												//������ʾ�Ի���
	window.location.href='reg.asp';											//��ת��ָ����ҳ��
	</script>
<%
	end if
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}

-->
</style></head>
<script language="javascript">
function checkemail(email)
{
var str=email;
//��JavaScript�У�������ʽֻ��ʹ��"/"��ͷ�ͽ���������ʹ��˫����
var Expression=/\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*/;
var objExp=new RegExp(Expression);
if(objExp.test(str)==true)
{return true;}
else
{return false;}
}
</script>
<script language="javascript">
function Mycheck()			//�����Զ��庯������֤������û���Ϣ�Ƿ�Ϸ�
{
if(form1.username.value=="")	//�ж��û������ı����ֵ�Ƿ�Ϊ��
{alert("�������û���!!");form1.username.focus();return;}	
if(form1.name1.value=="")
{alert("������������ʵ����!!");form1.name1.focus();return;}
if(form1.pwd.value=="")
{alert("���벻��Ϊ��!!");form1.pwd.focus();return;}
if(form1.pwd1.value=="")
{alert("ȷ�����벻��Ϊ��!!");form1.pwd1.focus();return;}
if(form1.pwd.value!=form1.pwd1.value)
{alert("������������벻����ȷ��!");form1.pwd.focus();return;}
if(form1.city.value=="")
{alert("���������ڳ�������!!");form1.city.focus();return;}
if(form1.address.value=="")
{alert("��������ϵ��ַ!!");form1.address.focus();return;}
if(form1.postcode.value=="")//�ж���������ŵ��ı����ֵ�Ƿ�Ϊ��
{alert("���������������!!");form1.postcode.focus();return;}
if(isNaN(form1.postcode.value))//�ж����������Ƿ���������
{alert("��������ű���Ϊ����!!");form1.postcode.focus();return;}
if(form1.telephone.value=="")
{alert("��������ϵ�绰!!");form1.telephone.focus();return;}
if(form1.email.value=="")
{alert("������E-mail!!");form1.email.focus();return;}
if(!checkemail(form1.email.value))
{alert("�����ַ��ʽ����ȷ������������!!");form1.email.focus();return;}
form1.submit();
}
</script>
<body>
<table width="800" height="520" border="0" align="center" cellpadding="0" cellspacing="0" background="images/new16.gif">
  <tr>
    <td height="86" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td width="46" height="289">&nbsp;</td>
    <td width="713" valign="top"><table width="693" height="70" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="693" height="70">
		<table width="670" height="294" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>
	<form action="" method="post" name="form1">
	<table width="581" height="285" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="125"><div align="center"><span class="STYLE2">�û���:</span></div></td>
        <td width="615"><input name="username" type="text" id="username" onKeyDown="if(event.keyCode==13){form1.name1.focus();}"></td>
      </tr>
      <tr>
        <td><div align="center" class="STYLE2">��ʵ����:</div></td>
        <td><input name="name1" type="text" id="name1" onKeyDown="if(event.keyCode==13){form1.pwd.focus();}"></td>
      </tr>
      <tr>
        <td><div align="center" class="STYLE2">��&nbsp;&nbsp;&nbsp;��:</div></td>
        <td><input name="pwd" type="password" id="pwd" onKeyDown="if(event.keyCode==13){form1.pwd1.focus();}"></td>
      </tr>
      <tr>
        <td><div align="center" class="STYLE2">ȷ������:</div></td>
        <td><input name="pwd1" type="password" id="pwd1" onKeyDown="if(event.keyCode==13){form1.city.focus();}"></td>
      </tr>
      <tr>
        <td><div align="center" class="STYLE2">���ڳ���:</div></td>
        <td><input name="city" type="text" id="city" onKeyDown="if(event.keyCode==13){form1.address.focus();}"></td>
      </tr>
      <tr>
        <td><div align="center" class="STYLE2">��ϵ��ַ:</div></td>
        <td><input name="address" type="text" id="address" size="30" onKeyDown="if(event.keyCode==13){form1.postcode.focus();}"></td>
      </tr>
      <tr>
        <td><div align="center" class="STYLE2">��������:</div></td>
        <td><input name="postcode" type="text" id="postcode" onKeyDown="if(event.keyCode==13){form1.telephone.focus();}"></td>
      </tr>
      <tr>
        <td><div align="center" class="STYLE2">��ϵ�绰:</div></td>
        <td><input name="telephone" type="text" id="telephone" onKeyDown="if(event.keyCode==13){form1.email.focus();}"></td>
      </tr>
      <tr>
        <td><div align="center" class="STYLE2">E-mail:</div></td>
        <td><input name="email" type="text" id="email" size="30" onKeyDown="if(event.keyCode==13){form1.Submit.focus();}"></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><input type="button" name="Submit" value="ȷ������" onClick="Mycheck();">
          <input type="reset" name="Submit2" value="������д">
          <input type="button" name="Submit3" value="����" onClick="JScript:window.location.href='index.asp';"></td>
      </tr>
    </table>
	</form>
	</td>
  </tr>
</table>
</td>
      </tr>
    </table></td>
    <td width="41">&nbsp;</td>
  </tr>
  <tr>
    <td height="13" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td height="127" colspan="3">&nbsp;</td>
  </tr>
</table>
</body>
</html>
