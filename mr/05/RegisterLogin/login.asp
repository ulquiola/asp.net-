<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>注册与登录</title>
<link type="text/css" rel="stylesheet" href="css/style.css" />
<script type="text/javascript" src="js/login.js"></script>
<script type="text/javascript" src="js/xmlhttprequest.js"></script>
</head>
<body>
<div id="contain">
  <div id="bgdiv">
	<div id="logindiv">
		<ul>
			<li>用户名：<input id="lgname" type="text" /></li>
			<li>密&nbsp;&nbsp;码：<input id="lgpwd" type="password" /></li>
			<li>验证码：<input id="lgchk" type="text" maxlength="4" /><span id="chknm" style="background-color:#f0f0f0; font-size: 18px; font-weight: bolder; color:#990000; margin: 0px 10px;"> </span> <a id="changea">看不清</a></li>
			<li><button id="lgbtn"></button>
			<button id="rgbtn"></button>
			<button id="fdbtn"></button>
			</li>
			<li><img id="regimg" src="images/loading.gif" style=" visibility: hidden;" /></li>
		</ul>
	</div>
  </div>
</div>
</body>
</html>
