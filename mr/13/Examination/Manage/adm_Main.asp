<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Validate.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��̨����</title>
<link rel="stylesheet" type="text/css" href="../Css/style.css">
</head>
<script>
function goto()
{
if(document.all("leftmenu").style.display == "")
	{
	document.all("leftmenu").style.display = "none";
	gotobar.innerHTML = "<img src='images/biao2.gif' alt='��С' style='cursor:hand'>";
	}
else
	{
	document.all("leftmenu").style.display = "";
	document.getElementById("gotobar").alt="�Ŵ�";
	gotobar.innerHTML = "<img src='images/biao.gif' alt='�Ŵ�' style='cursor:hand'>";
	}
}
</script>
<script src="Js/Check.js"></script>
<body topmargin="0" leftmargin="0">
<table height="100" cellpadding="0" cellspacing="0" border="0" align="center" width="782">
	<tr>
		<td style=" background-image:url(images/banner.jpg); background-position:center; background-repeat:no-repeat; height:129px;">&nbsp;</td>
	</tr>
</table>
<table cellpadding="0" cellspacing="0" border="0" align="center" width="782">
	<tr>
		<td width="185px" id="leftmenu">
			<iframe src="adm_Mainleft.asp" frameborder="0" width="100%" height="380" name="mainl" align="middle" scrolling="auto">�Բ�������������֧�ֿ�ܣ�</iframe>
		</td>
		<td width="45px" align="center" valign="middle">
		<a style="cursor:hand;" id="gotobar" onClick="goto()"><img src="images/biao.gif" alt="�Ŵ�"  width=20 height="19"></a>	
		</td>
		<td align="center" valign="top">
		<iframe src="adm_Mainright.asp" frameborder="0" width="100%" height="360" name="mainr" scrolling="auto">�Բ�������������֧�ֿ�ܣ�</iframe>	
		</td>
	</tr>
</table>
<table cellpadding="0" cellspacing="0" border="0" align="center" width="778">
<tr>
<td height="40" align="center" bgcolor="#E9E9E9">ȫ����СѧӢ��ѧϰ�ɼ����ԣ�NEAT������ʡ���԰칫��</td>
</tr>
</table>
</body>
</html>
