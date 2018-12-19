<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp"-->
<% 
	dim rndnum,verifycode
	Randomize
	Do While Len(rndnum)<4
	num1=CStr(Chr((57-48)*rnd+48))
	rndnum=rndnum&num1
	loop
	session("verifycode")=rndnum
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<style>
body{
	margin:0px;
	background-color:#FFFFFF;
}
body,th,td{
	font-size:12px;

}

</style>
</head>
<script src="Js/Check.js"></script>
<body topmargin="0" leftmargin="0">
<div style=" margin:0 auto; text-align: center;"> 
<div style="text-align: center; border-left: 1px #bfbfbf solid; border-right: 1px #bfbfbf solid; border-bottom: 3px #000000 solid; width: 782px;"><img src="images/banner.jpg" border="0" width="782" height="129" /></div>
<table border="1" cellpadding="0" cellspacing="0" align="center" width="782">
	<tr>
    	<td width="231" style=" text-align:center;">
   	     <table border="0" cellpadding="0" cellspacing="0" style=" margin:15px 10px;" align="center">
             <form name="Login" action="tre_/tre_Login.asp" method="post" onSubmit="return Check(Login)"> 
            	<tr>
                	<td width="211px" height="27px" style=" background-image:url(Images/lgtop.gif); background-position:center; background-repeat:no-repeat;">&nbsp;</td>
                </tr>
                <tr>
                    <td height="25px" style=" text-indent: 15px;">考&nbsp;&nbsp;号：&nbsp;&nbsp;<input name="用户名" type="text" style=" width:100px; height:18px; borlder: 1px #000000 solid;"></td>
                </tr>
                <tr>
                    <td height="25px" style=" text-indent: 15px;">密&nbsp;&nbsp;码：&nbsp;&nbsp;<input name="密码" type="password" style=" width:100px; height:18px; borlder: 1px #000000 solid;"></td>
				</tr>
                <tr>
                	<td height="25px"  style=" text-indent: 15px;">验证码：&nbsp;&nbsp;<input name="验证码" id="verifycode" size="6" maxlength="4" style=" width:50px; height:18px; borlder: 1px #000000 solid;"><font color="#000000" size="2">&nbsp;<%=session("verifycode")%><input type="hidden" name="验证码2" value="<%=session("verifycode")%>">
          </font></td>
                </tr>
                <tr>
                	<td height="25px" align="center" valign="middle">
                    <input type="submit" name="登录" value="登录"> &nbsp;&nbsp;&nbsp;<input type="button" name="忘记密码" value="获取密码" onClick="Javascript:window.open('addPass/Default.asp','','width=300,height=150')">
                    </td>
                </tr>
              </form>
            </table>
            <a href="http://www.mrbccd.com" target="_blank"><img src="Images/ad1.jpg" width="169" height="41" style=" border: 1px #000000 solid;" /></a>
            <a href="http://www.mingribook.com" target="_blank"><img src="Images/ad2.jpg" width="169" height="41" style=" border: 1px #000000 solid;" /></a>
            <div style=" height: 100px;">&nbsp;</div>
      </td>
      <td style=" text-align:center; vertical-align: top;">
  <table border="0" cellpadding="0" cellspacing="0" style=" margin:15px 10px;" align="center" width="90%" bgcolor="#FFFFFF">
            	<tr>
                	<td>
                    <marquee behavior="scroll"  direction="left" bgcolor="#FFFFFF" width="100%" hspace="1" vspace="1" scrollamount="2" scrolldelay="0" class="yellow">
             欢迎您进入英语成绩查询服务！
          </marquee>                    </td>
                </tr>
            	<tr>
                	<td style=" background-image:url(Images/centertop.gif); background-position: left; background-repeat:no-repeat; height:30px;">&nbsp;</td>
          </tr>
                <tr>
                	<td><hr style=" width:100%; height:4px; background-color: #70ba7b;" /></td>
                </tr>
                <tr>
                	<td align="left" valign="top"><!--#include file="Affiche.asp"--></td>
          </tr>
            </table>
      </td>
    </tr>
</table>
<div style="text-align: center; border-left: 1px #bfbfbf solid; border-right: 1px #bfbfbf solid; border-bottom: 1px #bfbfbf solid; width: 782px; height:50px; line-height: 50px; background-color:#E9E9E9">全国中小学英语学习成绩测试（NEAT）吉林省考试办公室</div>
</div>
</body>
</html>