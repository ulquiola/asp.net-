<!--#include file="inc/chk.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ӭ����������</title>
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
	if(event.keyCode == 116){				//��������ǡ�F5����
		alert('��ֹˢ��');				//���������
		event.keyCode = 0;				//����������
		return false;
	}
}
function Exit(){
	window.location.href="exit.asp";
	alert("��ӭ���´ι��٣�");
}
document.onkeydown=enterkey;				//����������onkeydown�¼�
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
