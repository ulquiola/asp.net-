<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Request���󼯺�֮��Form���Ϻ�QueryString����</title>
<style>
body{
	margin:12px;
	font-size:12px;

}
</style>
</head>
<body>
<%'���һ��Form��%>
<table width="399" height="166" border="0" cellpadding="0" cellspacing="0">
<form name="form1" method="post" action="index.asp">
  <tr>
    <td colspan="2"><div align="center">��Ա��Ϣ</div>
        <div align="center"></div></td>
  </tr>
  <tr>
    <td width="79"><div align="center">��Ա�ǳƣ�</div></td>
    <td width="203"><input name="name1" type="text" id="name1" value=<%=request.Form("name1")%> /></td>
  </tr>
  <tr>
    <td><div align="center">��Ա���ƣ�</div></td>
    <td><input name="name2" type="text" id="name2" value="<%=request.Form("name2")%>" /></td>
  </tr>
  <tr>
    <td><div align="center">����������</div></td>
    <td><select name="dq" id="dq">
      <option value="������">������</option>
      <option value="�׳���">�׳���</option>
      <option value="��Դ��">��Դ��</option>
      <option value="��ɽ��">��ɽ��</option>
      <option value="������">������</option>
    </select>
    </td>
  </tr>
  <tr>
    <td><div align="center">ͨѶ��ַ��</div></td>
    <td><input name="address" type="text" id="address" value="<%=request.Form("address")%>" /></td>
  </tr>
  <tr>
    <td colspan="2"><div align="center">
      <input type="submit" name="Submit" value="�ύ" />
      <%'���һ���ύ��ť%>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      <input type="reset" name="Submit2" value="����" />
      <%'���һ�����ð�ť%>
    </div></td>
  </tr>
</form>
</table>

<hr />
<%
	if request.Form("name1") <> "" then
%>
&nbsp;&nbsp;&nbsp;&nbsp;��Ա��Ϣ<br />
��Ա�ǳƣ�<%=request.Form("name1")%><br> 	<%'Ӧ��Form���ݼ��ϻ�ȡ������%>
��Ա���ƣ�<%=request.Form("name2")%><br>	<%'Ӧ��Form���ݼ��ϻ�ȡ������%>
����������<%=request.Form("dq")%><br> 		<%'Ӧ��Form���ݼ��ϻ�ȡ������%>
ͨѶ��ַ��<%=request.Form("address")%> 		<%'Ӧ��Form���ݼ��ϻ�ȡ������%>
<%
	end if
%>


</body>
</html>
