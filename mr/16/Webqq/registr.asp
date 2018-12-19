<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>用户注册</title>
<link href="./Css/style.css" type="text/css" rel="stylesheet">
<script src="JS/check.js"></script>
<style type="text/css">
<!--
body {
	margin-top: 0px;
	margin-bottom: 0px;
	margin-left: 0px;
	margin-right: 0px;
}
.STYLE2 {
	font-size: 9pt;
	color: #000000;
}
-->
</style>
<link href="../Css/style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
}
a:hover {
	text-decoration: none;
}
a:active {
	text-decoration: none;
}
.STYLE3 {
	font-size: 9pt;
	color: #666666;
}
-->
</style></head>
<script language="javascript">
var http_request = false;
function createRequest(url) 
{
	//初始化对象并发出XMLHttpRequest请求
	http_request = false;
	if (window.XMLHttpRequest) { // Mozilla或其他除IE以外的浏览器
		http_request = new XMLHttpRequest();
		if (http_request.overrideMimeType) {
			http_request.overrideMimeType("text/xml");
		}
	} else if (window.ActiveXObject) { // IE浏览器
		try {
			http_request = new ActiveXObject("Msxml2.XMLHTTP");
		} catch (e) {
			try {
				http_request = new ActiveXObject("Microsoft.XMLHTTP");

		   } catch (e) {}
		}
	}
	if (!http_request) {
		alert("不能创建XMLHTTP实例!");
		return false;
	}
	http_request.onreadystatechange = alertContents;    //指定响应方法
	//发出HTTP请求
	http_request.open("GET", url, true);
	http_request.send(null);
}
function alertContents() {    //处理服务器返回的信息
	if (http_request.readyState == 4) {//表示数据接收完毕
		if (http_request.status == 200) {//返回当前请求的Http状态码
			alert(http_request.responseText);
			//aaaa.innerHTML="内容："+http_request.responseText;
		} else {
			alert('您请求的页面发现错误');
		}
	}

}
</script>
<script language="javascript">
function checkName() {//创建自定义函数
	var UserName = myform.UserName.value;//为UserName变量赋值
	if(UserName=="") {//判断用户名是否为空
		window.alert("请添写用户名!");//弾出提示对话框
		myform.UserName.focus();//获取焦点
		return false;
	}
	else {
		createRequest('yan.asp?UserName='+UserName);//跳转到验证页面
	}
}
</script>
<body>
<table width="560" height="429"  border="0" cellpadding="-2" cellspacing="-2" background="Images/re.jpg">
  <tr>
    <td width="604" valign="top"><table width="100%" height="205"  border="0" cellpadding="-2" cellspacing="-2" class="tableBorder">
      <tr>
        <td width="10" height="32">&nbsp;</td>
        <td width="532" style="color:#FFFFFF">&nbsp;</td>
        <td width="63">&nbsp;</td>
      </tr>
      <tr valign="top">
        <td height="171" colspan="3"><table width="103%" height="129"  border="0" cellpadding="-2" cellspacing="-2">
          <tr>
            <td valign="top"><table width="103%" height="377"  border="0" cellpadding="-2" cellspacing="-2" class="tableBorder_Top">
              <tr>
                <td height="5"></td>
              </tr>
              <tr>
                <td width="349" height="263" valign="top">
				  <div align="right"></div>
				  <table width="100" border="0" align="left" cellpadding="-2" cellspacing="-2">
<tr>
<td><br><br><br><form action="register_deal.asp" method="post" name="myform">
<table width="349" height="301" border="0" align="right" cellpadding="-2" cellspacing="-2">
<tr>
<td width="21%" height="30" align="center" class="STYLE2">&nbsp;&nbsp;用户名：</td>
<td width="79%"><input name="UserName" type="text" id="UserName4" maxlength="20" class="wenbenkuang">
<a href="#" class="STYLE3" onClick="checkName()">[检测用户名]</a> <span class="STYLE2">*</span> </td>
</tr>
<tr>
<td height="28" align="center" class="STYLE2">&nbsp;&nbsp;密&nbsp;&nbsp;码：</td>
<td height="28"><input name="Password1" type="password" class="wenbenkuang" id="PWD14" size="20" maxlength="20">
<span class="STYLE2">*</span></td>
</tr>
<tr>
<td height="28" align="center" class="STYLE2">&nbsp;&nbsp;确认密码：</td>
<td height="28"><input name="Password2" type="password" class="wenbenkuang" id="PWD25" size="20" maxlength="20">
<span class="STYLE2"> *</span> </td>
</tr>
<tr>
<td height="28" align="center" class="STYLE2">&nbsp;&nbsp;性&nbsp;&nbsp;&nbsp;别：</td>
<td class="STYLE2"><input name="sex" type="radio" class="noborder" value="男" checked>
<span class="STYLE2">男</span>&nbsp;
<input name="sex" type="radio" class="noborder" value="女">
女</td>
</tr>
<tr>
<td height="28" align="center" class="STYLE2">&nbsp;&nbsp;生&nbsp;&nbsp;&nbsp;&nbsp;日：</td>
<td class="word_grey"><input name="birthday" type="text" class="wenbenkuang" id="Tel">
<span class="STYLE2"> *(格式为:yyyy-mm-dd) </span></td>
</tr>
<tr>
<td height="97" align="center" class="STYLE2">&nbsp;&nbsp;选择头像：</td>
<td class="word_grey"><table width="172" cellpadding="0" cellspacing="0">
<tr>
<td width="8" height="47"><script language="javascript">
//查看全部头像并选择时应用该函数
function deal(){
var someValue;
someValue=window.showModalDialog('headbrowse.asp','','dialogWidth=510px;\
dialogHeight=370px;status=no;help=no;scrollbars=no');
if (someValue == undefined){ //当用户在弹出的网页对话框中没有选择头像时
someValue="0";
} 
document.images.img.src="images/head/"+someValue+".gif";
myform.touxiang.value=someValue;//添加隐藏域来获取头像值
}
</script>
</td>
<td width="97"><img src="Images/head/0.gif" name="img" width="50" height="50"></td>
<td width="245">&nbsp;</td>
</tr>
<tr>
<td>&nbsp;</td>
<td>[<a href="#" class="STYLE3" onClick="deal()">选择头像</a>] <input name="touxiang" type="hidden" id="toxian"></td>
<td>&nbsp;</td>
</tr>
</table></td>
</tr>
<tr>
<td height="28" align="center" class="STYLE2" style="padding-left:10px">E-mail：</td>
<td class="word_grey"><input name="Email" type="text" class="wenbenkuang" id="PWD224" size="30">
<span class="STYLE2">*</span></td>
</tr>
<tr>
<td height="34">&nbsp;</td>
<td class="word_grey"><input name="Button" type="button" class="btn_grey" value="确定保存" onClick="mycheck()">
<input name="Submit2" type="reset" class="btn_grey" value="重新填写"></td>
</tr>
</table>
</form></td>
</tr>
</table> 
				</td>
                <td width="7"></td>
                <td width="204" valign="top"><p>&nbsp;</p><br>
                  <table width="149" height="173"  border="0" align="left" cellpadding="-2" cellspacing="-2">

                  <tr>
                    <td width="100%" height="173" valign="top" class="word_grey">⊙ 用户名：可使用英文字母、数字或英文字母、数字、下划线的组合，长度控制在3-20个字符之内。<br>⊙生日：输入您的生日，如果您的生日是1900年3月13日则输入：1900-03-13。<br>⊙头像：通过头像下拉列表框选择头像。<br>⊙E-mail：请填写有效的E-mail地址，以便于与您联系。</td>
                  </tr>
                  <tr align="center">
                    <td valign="middle" class="word_grey"></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td height="5"></td>
              </tr>
            </table></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
