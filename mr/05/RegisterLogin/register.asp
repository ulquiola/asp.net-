<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>用户注册</title>
<link type="text/css" rel="stylesheet" href="css/style.css" />
<script type="text/javascript" src="js/xmlhttprequest.js"></script>
<script type="text/javascript" src="js/register.js"></script>
</head>
<body>
<div id="contain">
  <div id="rgbgdiv">
	<div id="registerdiv">
		<ul>
			<li>用户名称：<input id="regname" type="text" />&nbsp;<span id="namediv" class="morediv">&nbsp;</span>
			</li>
			<li>用户密码：<input id="regpwd1" type="password" />&nbsp;<span id="pwddiv1" class="morediv">&nbsp;</span>
			</li>
			<li>确认密码：<input id="regpwd2" name="regpwd2" type="password" />
			<span id="pwddiv2" class="morediv"></span>
			</li>
		</ul>
		<div id="morediv" style="display:none;" >
		<hr id="part" />
		<ul>
			<li>密保问题：<input id="question" type="text" />&nbsp;<span class="morediv">&nbsp;</span>
			</li>
			<li>密保答案：<input id="answer" type="text" />&nbsp;<span class="morediv">&nbsp;</span>
			</li>
			<li>真实姓名：<input id="realname" type="text" />&nbsp;<span class="morediv">&nbsp;</span>
			</li>
			<li>联系电话：<input id="telephone" type="text" />&nbsp;<span class="morediv">&nbsp;</span>
			</li>
			<li>电子邮件：<input id="email" type="text" />&nbsp;<span id="emaildiv" class="morediv">&nbsp;</span>
			</li>
			<li>QQ号&nbsp;&nbsp;码：<input id="qq" type="text" />&nbsp;<span class="morediv">&nbsp;</span>
			</li>
		</ul>
		</div>
		<div><button id="regbtn" disabled="true"></button>&nbsp;<button id="morebtn"></button>&nbsp;<button id="logbtn"></button></div>
		<div><img id="imgdiv" src="images/loading.gif" style=" visibility: hidden;" /></div>
	</div>
  </div>
</div>
</body>
</html>
