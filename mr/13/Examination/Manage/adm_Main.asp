<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Validate.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>后台管理</title>
<link rel="stylesheet" type="text/css" href="../Css/style.css">
</head>
<script>
function goto()
{
if(document.all("leftmenu").style.display == "")
	{
	document.all("leftmenu").style.display = "none";
	gotobar.innerHTML = "<img src='images/biao2.gif' alt='缩小' style='cursor:hand'>";
	}
else
	{
	document.all("leftmenu").style.display = "";
	document.getElementById("gotobar").alt="放大";
	gotobar.innerHTML = "<img src='images/biao.gif' alt='放大' style='cursor:hand'>";
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
			<iframe src="adm_Mainleft.asp" frameborder="0" width="100%" height="380" name="mainl" align="middle" scrolling="auto">对不起，你的浏览器不支持框架！</iframe>
		</td>
		<td width="45px" align="center" valign="middle">
		<a style="cursor:hand;" id="gotobar" onClick="goto()"><img src="images/biao.gif" alt="放大"  width=20 height="19"></a>	
		</td>
		<td align="center" valign="top">
		<iframe src="adm_Mainright.asp" frameborder="0" width="100%" height="360" name="mainr" scrolling="auto">对不起，你的浏览器不支持框架！</iframe>	
		</td>
	</tr>
</table>
<table cellpadding="0" cellspacing="0" border="0" align="center" width="778">
<tr>
<td height="40" align="center" bgcolor="#E9E9E9">全国中小学英语学习成绩测试（NEAT）吉林省考试办公室</td>
</tr>
</table>
</body>
</html>
