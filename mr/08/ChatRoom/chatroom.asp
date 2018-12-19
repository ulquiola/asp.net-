<!--#include file="inc/chk.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>欢迎光临聊天室</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></head>
<script language="javascript">
function enterkey(){
	if(event.keyCode == 116){				//如果按键是〈F5〉键
		alert('禁止刷新');				//弹出警告框
		event.keyCode = 0;				//将按键归零
		return false;
	}
}
function Exit(){
	window.location.href="exit.asp";
	alert("欢迎您下次光临！");
}
document.onkeydown=enterkey;				//将函数赋给onkeydown事件
</script>
<script language="javascript" src="JS/xmlHttpRequest.js"></script>
</head>
<frameset rows="150,*,120" cols="*" frameborder="yes" border="3" framespacing="0" onUnload="Exit()">
  <frame src="top.asp" name="topFrame" scrolling="NO" noresize title="topFrame">
  <frameset cols="160,*" frameborder="NO" border="0" framespacing="0" onUnload="Exit()">
    <frame src="left.asp" name="leftFrame" scrolling="yes" noresize title="leftFrame">
    <frame src="main.asp" name="mainFrame" title="mainFrame">
  </frameset>
  <frame src="bottom.asp" frameborder="yes" bordercolor="#666600" name="bottomFrame" scrolling="NO" noresize title="bottomFrame">
</frameset>
<noframes><body>
</body></noframes>
</html>
