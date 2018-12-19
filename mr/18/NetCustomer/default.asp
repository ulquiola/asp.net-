<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% session.Timeout=120 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>网上客户管理系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	background-image: url(Images/bg.gif);
}
-->
</style></head>
<body>
<table height="242" border="0" cellpadding="-2" cellspacing="-2">
  <tr>
    <td valign="top"><img src="Images/default.gif" width="595" height="399" border="0" usemap="#Map">
      <map name="Map">
        <area shape="rect" coords="70,324,232,359" href="login_manager.asp" 
		alt="如果您是系统管理员，请点击！">
        <area shape="rect" coords="350,323,506,357" href="login_worker.asp" 
		alt="如果您是业务员，请点击！">
      </map></td>
  </tr>
</table>
</body>
</html>
