<%@LANGUAGE="VBSCRIPT"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>用户登录页面</title>
<script language="javascript">
self.resizeTo("320","220")		//设置窗口的大小
width=screen.width				//表示用户屏幕，提供屏幕宽度
height=screen.height			//表示用户屏幕，提供屏幕高度
self.moveTo((width-320)/2,(height-220)/2)//将窗口移动到指定坐标处
</script>
<link href="./CSS/style.css" rel="stylesheet" type="text/css">
<script src="JS/check.js"></script>
</head>
<style>
<!--
.wenben{
border:1px solid #9933FF;}
.STYLE2 {
	font-size: 9pt;
	color: #000000;
}
-->
</style>
<body topmargin="0" leftmargin="0" style="border:none" scroll=no>
<table width="321" height="199" border="0" align="center" cellpadding="0" cellspacing="0" background="Images/1.gif">
  <tr>
    <td>
	<form name="form1" method="POST">
  <p>&nbsp;</p>
  <p>&nbsp;</p><br>
  <table width="67%" height="33%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr height="51"> 
      <td height="2" colspan="2"></td>
    </tr>
    <tr> 
      <td width="42%" height="22" align="right"><p class="STYLE2">用户名：</p></td>
      <td width="58%" align="left"><font size="2"> 
        <input name="UserName" type="text" class="wenbenkuang" id="UserName" size="19">
      </font></td>
    </tr>
    <tr> 
      <td width="42%" height="43" align="right"><p class="STYLE2">密&nbsp;&nbsp;码：</p>
      </td>
      <td width="58%" align="left"><p><font size="2">
        <input name="Password1" type="password" class="wenbenkuang" id="Password1" size="19">
        </font></p>
        </td>
    </tr>
    <tr> 
      <td height="21" colspan="2" align="right"> 
        <div align="center">
          <input type="submit" value="登录" name="B3" onClick="mycheck1()">
          &nbsp;
          <input type="button" value="取消" name="B3" onClick="window.close()">
          &nbsp;
          <input type="button" value="注册" name="B3" onClick="window.open('registr.asp','','width=550,height=429')">
        </div></td>
  </tr></table>
</form>
	</td>
  </tr>
</table>
</body>
</html>