<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../../Conn/conn.asp"-->
<%
getid = Replace(Request("id"),"'"," ")
getcondition = Replace(Request("condition"),"'"," ")
getkey = Replace(Request("key"),"'"," ")
getpageno = Replace(Request("pageno"),"'"," ")
Sql = "Select * from tb_Administrator Where ID = "&getid
Set rs = conn.Execute(Sql)
If(rs.Eof)Then
%>
<script type="text/javascript">
	alert('û�д˹���Ա');
	history.go(-1);
</script>
<%
	Response.End()
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�޸Ĺ���Ա��Ϣ</title>
<link rel="stylesheet" type="text/css" href="../../Css/style.css">
<style type="text/css">
<!--
.style1 {color: #FFFFFF}
-->
</style>
</head>
<script>
function Check()
{
	if(UpdateAdmin.AdminName.value == ""){alert("����Ա���Ʋ���Ϊ��!");UpdateAdmin.AdminName.focus();return false;}
	if(UpdateAdmin.AdminPass.value == ""){alert("����Ա���벻��Ϊ��!");UpdateAdmin.AdminPass.focus();return false;}
}
</script>
<body topmargin="0" leftmargin="0">
	<table cellpadding="0" cellspacing="0" border="0" align="center" width="100%" height="100%">
		<tr>
			<td height="30" align="left">&nbsp;</td>
		</tr>
		<tr>
			<td align="center">
			<form name="UpdateAdmin" action="tre_/tre_adm_UpdateAdmin.asp" method="post" onSubmit="return Check()">
				<table cellpadding="0" cellspacing="0" border="0" align="center" width="400" height="65">
					<tr>
						<td width="174" height="30" align="right">
						����Ա���ƣ�&nbsp;&nbsp;
						</td>
						<td width="226" align="left">
						<input type="text" name="AdminName" class="txt_grey" value="<%=Server.HtmlenCode(rs("Name"))%>">					  </td>
					</tr>
					<tr>
						<td width="174" height="30" align="right">
						����Ա���룺&nbsp;&nbsp;
						</td>
						<td width="226" align="left">
						<input type="password" name="AdminPass" class="txt_grey" value="<%=Server.HtmlenCode(rs("PWD"))%>">
						<input type="hidden" name="id" value="<%=rs("ID")%>">
						<input type="hidden" name="condition" value="<%=getcondition%>">
						<input type="hidden" name="key" value="<%=getkey%>">
					  <input type="hidden" name="pageno" value="<%=getpageno%>">					  </td>
					</tr>
					<tr>
						<td height="5" colspan="2">&nbsp;
						
						</td>
					</tr>
					<tr>
						<td height="30" colspan="2" align="center">
						<input type="submit" name="�޸�" value="�޸�" class="btn_grey">
						<input type="reset" name="����" value="����" class="btn_grey">
						<input name="����" type="button" class="btn_grey" onClick="Javascript:history.go(-1);" value="����">
						</td>
					</tr>
			  </table>
			</form>
			</td>
		</tr>
	</table>
</body>
</html>
