<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../conn/conn.asp"-->
<!--#include file="../inc/customFunc.asp"-->
<%
if request.QueryString("ID")<>"" then
ID=request.QueryString("ID")
end if
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from tb_board where ID="&ID
rs.open sql,conn,1,3
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
<script language="javascript">
function check(myform)
{
	if (myform.content1.value=="")
	{
		alert("������ظ�����!");
		myform.content1.focus();
		return false;
	}
}
</script>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
}
a:hover {
	text-decoration: none;
}
a:active {
	text-decoration: none;
}
-->
</style></head>
<body>
<table width="403" height="230" align="center" cellpadding="0" cellspacing="0">
<form name="myform" method="post" action="reply.asp" />
  <tr>
    <td width="401" height="208" valign="top"><table width="430" height="181" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="12" colspan="2" bgcolor="#327387"><div align="center" class="STYLE2"><span class="STYLE1"><font color="#FFFFFF">���Բ���Ϣ�鿴ҳ��</font></span></div></td>
      </tr>
      <tr>
        <td width="87" height="8"><div align="center"><span class="STYLE2">�����ˣ�</span></div></td>
        <td width="343"><input name="person" type="text" id="person" value="<%=rs("person")%>" readonly="readonly" /></td>
      </tr>
      <tr>
        <td width="87" height="8"><div align="center"><span class="STYLE2">��ϵ�绰��</span></div></td>
        <td width="343"><input name="person" type="text" id="person" value="<%=rs("tel")%>" readonly="readonly" /></td>
      </tr>
      <tr>
        <td width="87" height="8"><div align="center"><span class="STYLE2">QQ��</span></div></td>
        <td width="343"><input name="person" type="text" id="person" value="<%=rs("qq")%>" readonly="readonly" /></td>
      </tr>
      <tr>
        <td width="87" height="8"><div align="center"><span class="STYLE2">���䣺</span></div></td>
        <td width="343"><input name="person" type="text" id="person" value="<%=rs("email")%>" readonly="readonly" /></td>
      </tr>
      <tr>
        <td width="87" height="8"><div align="center"><span class="STYLE2">ԤԼ�ˣ�</span></div></td>
        <td width="343"><input name="person" type="text" id="person" value="<%=rs("yuyueren")%>" readonly="readonly" /></td>
      </tr>
      <tr>
        <td width="87" height="8"><div align="center"><span class="STYLE2"≯��ʱ��Σ�</span></div></td>
        <td width="343"><input name="person" type="text" id="person" value="<%=rs("talktime")%>" readonly="readonly" /></td>
      </tr>
      <tr>
        <td height="17"><div align="center" class="STYLE2">��ѯ���⣺</div></td>
        <td><input name="title" type="text" id="title" value="<%=rs("title")%>" size="30" readonly="readonly" /></td>
      </tr>
      <tr>
        <td height="30"><div align="center" class="STYLE2">��ѯ���ˣ�</div></td>
        <td><textarea name="textarea" cols="30" rows="5" readonly="readonly" ><%if rs("reply")<>"" then%><%=rs("content")%><%else%>���޻ظ���Ϣ��<%end if%>
        </textarea></td>
      </tr>
      <tr>
        <td height="30"><div align="center" class="STYLE2">�ظ����ԣ�</div></td>
        <td><textarea name="content1" cols="30" rows="5" ><%=rs("reply")%></textarea></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="20"><div align="center">
	<Input name="ID" type="hidden" value="<%=rs("id")%>" /><input type="submit" value="�ظ�" style="border: 0px; background-color:#FFFFFF;" onClick="return check(myform)" />&nbsp;&nbsp;<input  type="button" value="�رմ���" onClick="javascript:window.close()"  style="border: 0px; background-color:#FFFFFF;"></div></td>
  </tr>
</table>
</body>
</html>