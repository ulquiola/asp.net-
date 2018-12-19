<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" --><!--包含数据库连接文件-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>管理员登录</title>
<link rel="stylesheet" type="text/css" href="../css/style.css">
<style type="text/css">
<!--
.STYLE1 {
	font-size: 9pt;
	color: #000000;
}
.STYLE15 {color: #000066; font-size: 12px; }
-->
</style>
</head>
<script language="javascript">
function mycheck(myform,tool)//创建自定义函数
{
if (myform.UID.value=="")//判断用户名文本框是否有值
{alert("请输入"+tool+"！");myform.UID.focus();return;}//弹出提示对话框
if(myform.PWD.value=="")//判断密码文本框是否有值
{alert("请输入密码！");myform.PWD.focus();return;}		//弹出提示对话框
if(myform.yanzheng.value=="")//判断验证码文本框是否有值
{alert("请输入验证码!");myform.yanzheng.focus();return;}//弹出提示对话框
if(myform.yanzheng.value!=myform.verifycode2.value)//判断两次输入的验证码是否相同
{alert("请输入正确的验证码!!");myform.yanzheng.focus();return;}//弹出提示对话框
myform.submit();//提交表单
}
</script>
<script src="../js/fun.js"></script>
<body bgcolor="#B9DFFF">
<%   
    '自动生成4位的随机验证码
	randomize
	i=0					'为变量i赋值
	do while i<4		
	num1=int(9*rnd)		'对生成的数字取整
	numimage="<img src=../images/num/"&num1&".gif>" '获取图片
	numi=numi&numimage	'为变量numi赋值
	num=num&num1		'为变量num赋值
	i=i+1				'i循环加1
	loop
	session("numi")=numi	'为session("numi")变量赋值
	session("num")=num		'为session("num")变量赋值
%>
<table width="500" height="400" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td background="../images/login.gif"><table width="350" height="200" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="54">&nbsp;</td>
        <td width="296">
		<table width="259" height="139"  border="0" align="center" cellpadding="-2" cellspacing="-2">
  <tr>
    <td height="139" align="center"><form name="form_M" method="post" action="Login_M.asp">
        <table width="262" height="120" border="0" align="center" cellpadding="-2" cellspacing="-2" class="tabBorder">
          <tr align="center" valign="middle">
            <td height="24" colspan="2" valign="bottom" background="Images/bg_Login.GIF"><font color="#505875">=== 管理员登录 === </font> </td>
          </tr>
          <tr>
            <td width="95" height="27" align="right" class="STYLE1">&nbsp;&nbsp;&nbsp;管理员名称：</td>
            <td width="167"><div align="left">
              <input name="UID" type="text" class="inputcss1" maxlength="20">
            </div></td>
          </tr>
          <tr>
            <td height="16" align="right" valign="middle" class="STYLE1">密&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;码：</td>
            <td valign="middle"><input name="PWD" type="password" class="inputcss1" onKeyDown="if(event.keyCode==13) mycheck(form_M,'版主名称')"  maxlength="20"></td>
          </tr>
          <tr>
            <td height="13" align="right" valign="middle" class="STYLE1"><span class="STYLE1">验&nbsp;&nbsp;证&nbsp;&nbsp;码：</span></td>
            <td valign="middle"><span class="STYLE16">
              <input name="yanzheng" type="text" class="inputcss1" id="yanzheng" size="10">
			  <input type="hidden" name="verifycode2" value="<%=session("num")%>"><%=session("numi")%>
            </span></td>
          </tr>
          <tr align="center" valign="middle">
            <td height="27" colspan="2"><input name="Submit" type="button" class="btn_grey" value="登录" onClick="mycheck(form_M,'版主名称')">
&nbsp;
            <input name="Submit2" type="reset" class="btn_grey" value="重置"></td>
          </tr>
        </table>
    </form></td>
  </tr>
</table>
		
		</td>
      </tr>
    </table></td>
  </tr>
</table>
<table width="500" height="30" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><div align="center" class="STYLE1" style=" line-height:2">版权所有：吉林省明日科技有限公司<br>Copyright&nbsp;&copy;&nbsp;www.mingrisoft.com All Rights Reserved!
</div></td>
  </tr>
</table>
</body>
</html>
