<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�����Ϣ</title>
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
	if(AddAdmin.AdminName.value == ""){alert("����Ա���Ʋ���Ϊ��!");AddAdmin.AdminName.focus();return false;}
	if(AddAdmin.AdminPass.value == ""){alert("����Ա���벻��Ϊ��!");AddAdmin.AdminPass.focus();return false;}
}
</script>
<body topmargin="0" leftmargin="0">
	<table cellpadding="0" cellspacing="0" border="0" align="center" width="100%" height="100%">
		<tr>
			<td height="30" align="left">&nbsp;</td>
		</tr>
		<tr>
			<td align="center">
			<form name="AddAdmin" action="tre_/tre_adm_AddAdmin.asp" method="post" onsubmit="return Check()">
				<table cellpadding="0" cellspacing="0" border="0" align="center" width="400" height="65">
					<tr>
						<td width="174" height="30" align="right">
						����Ա���ƣ�&nbsp;&nbsp;
						</td>
						<td width="226" align="left">
						<input type="text" name="AdminName" class="txt_grey">					  </td>
					</tr>
					<tr>
						<td width="174" height="30" align="right">
						����Ա���룺&nbsp;&nbsp;
						</td>
						<td width="226" align="left">
						<input type="password" name="AdminPass" class="txt_grey">					  </td>
					</tr>
					<tr>
						<td height="5" colspan="2">&nbsp;
						
						</td>
					</tr>
					<tr>
						<td height="30" colspan="2" align="center">
						<input type="submit" name="���" value="���" class="btn_grey">
						<input type="reset" name="����" value="����" class="btn_grey">
						<input name="����" type="button" class="btn_grey" onclick="Javascript:history.go(-1);" value="����">
						</td>
					</tr>
			  </table>
			</form>
			</td>
		</tr>
	</table>
</body>
</html>
