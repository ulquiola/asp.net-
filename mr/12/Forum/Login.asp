<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>用户登录</title>
<link rel="stylesheet" type="text/css" href="css/style.css">
<script language="javascript">
function mycheck(myform,tool)								//创建自定义函数
{
if (myform.UID.value=="")									//判断用户ID是否为空
{alert("请输入"+tool+"！");myform.UID.focus();return;}		//弹出提示对话框
if(myform.PWD.value=="")									//判断密码文本框的值是否为空
{alert("请输入密码！");myform.PWD.focus();return;}			//弹出提示对话框
if(myform.yanzheng.value=="")								//判断验证码的文本框的值是否为空
{alert("请输入验证码!");myform.yanzheng.focus();return;}	//弹出提示对话框
if(myform.yanzheng.value!=myform.verifycode2.value)
{alert("请输入正确的验证码!!");myform.yanzheng.focus();return;}
myform.submit();											//提交表单
}
</script>
<style type="text/css">
<!--
.STYLE15 {color: #000066; font-size: 12px; }
.STYLE16 {color: #316BD6; font-size: 9pt; }
.input2 {	font-size: 12px;
	border: 3px double #A8D0EE;
	color: #344898;
}
-->
</style>
</head>
<script src="js/fun.js"></script>
<body bgcolor="#B9DFFF" class="scrollbar">
<%   
    '自动生成4位的随机验证码
	randomize
	i=0
	do while i<4
	num1=int(9*rnd)
	numimage="<img src=images/num/"&num1&".gif>"
	numi=numi&numimage
	num=num&num1
	i=i+1
	loop
	session("numi")=numi
	session("num")=num
%>
<table width="500" height="400" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td background="images/login.gif"><table width="350" height="200" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="54">&nbsp;</td>
        <td width="296"><table width="240" height="139"  border="0" align="center" cellpadding="-2" cellspacing="-2">
  <tr>
    <td height="139" align="center"><form name="form_U" method="post" action="Login_deal.asp">
        <table width="220" height="120" border="0" align="center" cellpadding="-2" cellspacing="-2" class="tabBorder">
          <tr align="center" valign="middle">
            <td height="24" colspan="2" valign="bottom"><font color="#505875">=== 用户登录 === </font> </td>
          </tr>
          <tr>
            <td width="61" height="27" align="right" valign="middle">用户名：</td>
            <td width="157" valign="middle"><input name="UID" type="text" class="inputcss1" maxlength="20"></td>
          </tr>
          <tr>
            <td height="16" align="right" valign="middle">密 &nbsp;码：</td>
            <td valign="middle"><input name="PWD" type="password" class="inputcss1"  maxlength="20"></td>
          </tr>
          <tr>
            <td height="13" align="right" valign="middle"><span class="STYLE15">验证码：</span></td>
            <td valign="bottom"><span class="STYLE16">
              <input name="yanzheng" type="text" class="inputcss1" id="yanzheng" size="10">
			  <input type="hidden" name="verifycode2" value="<%=session("num")%>"><%=session("numi")%>
            </span></td>
          </tr>
          <tr align="center" valign="middle">
            <td height="27" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              <input name="Submit" type="button" class="btn_grey" value="登录" onClick="mycheck(form_U,'用户名')">
  &nbsp;<input name="Submit" type="reset" class="btn_grey" value="重置">
&nbsp;            </td></tr>
        </table>
    </form></td>
  </tr>
</table></td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
