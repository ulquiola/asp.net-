<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="top.asp"--><%'包含网站的头文件%>
<!--#include file="conn/conn.asp"--><%'包含数据库连接文件%>
<%
if request.Form("username")<>"" then	'判断用户名是否为空
Set rs= Server.CreateObject("ADODB.Recordset")	'创建记录集
	sql="select username from tb_user where username='"&Request.Form("username")&"'"	'查询数据
	rs.open sql,conn,1,3																'打开记录集
	if rs.eof then 																				
	username=request.Form("username")										'获取表单元素username的值并赋给username变量
	name1=request.Form("name1")												'获取表单元素name1的值并赋给name1变量
	pwd=request.Form("pwd")													'获取表单元素pwd的值并赋给pwd变量
	city=request.form("city")												'获取表单元素city的值并赋给city变量
	address=request.Form("address")											'获取表单元素address的值并赋给address变量
	postcode=request.Form("postcode")										'获取表单元素postcode的值并赋给postcode变量
	telephone=request.Form("telephone")										'获取表单元素telephone的值并赋给telephone变量
	email=request.Form("email")												'获取表单元素email的值并赋给email变量
	sql="insert into tb_user(username,name1,pwd,city,address,postcode,telephone,email) values('"&username&"','"&name1&"','"&pwd&"','"&city&"','"&address&"','"&postcode&"','"&telephone&"','"&email&"')"	'添加记录信息
	conn.execute(sql)														'执行SQL语句
%>
	<script language="javascript">
	alert("用户信息添加成功!!");											//弹出提示信息
	window.location.href='index.asp';										//跳转到指定的页面
	</script>
<%else%>
	<script language="javascript">
	alert("该用户已经存在！");												//弹出提示对话框
	window.location.href='reg.asp';											//跳转到指定的页面
	</script>
<%
	end if
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}

-->
</style></head>
<script language="javascript">
function checkemail(email)
{
var str=email;
//在JavaScript中，正则表达式只能使用"/"开头和结束，不能使用双引号
var Expression=/\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*/;
var objExp=new RegExp(Expression);
if(objExp.test(str)==true)
{return true;}
else
{return false;}
}
</script>
<script language="javascript">
function Mycheck()			//创建自定义函数来验证输入的用户信息是否合法
{
if(form1.username.value=="")	//判断用户名的文本框的值是否为空
{alert("请输入用户名!!");form1.username.focus();return;}	
if(form1.name1.value=="")
{alert("请输入您的真实姓名!!");form1.name1.focus();return;}
if(form1.pwd.value=="")
{alert("密码不能为空!!");form1.pwd.focus();return;}
if(form1.pwd1.value=="")
{alert("确认密码不能为空!!");form1.pwd1.focus();return;}
if(form1.pwd.value!=form1.pwd1.value)
{alert("两次输入的密码不致请确认!");form1.pwd.focus();return;}
if(form1.city.value=="")
{alert("请输入所在城市名称!!");form1.city.focus();return;}
if(form1.address.value=="")
{alert("请输入联系地址!!");form1.address.focus();return;}
if(form1.postcode.value=="")//判断邮政编码号的文本框的值是否为空
{alert("请输入邮政编码号!!");form1.postcode.focus();return;}
if(isNaN(form1.postcode.value))//判断邮政编码是否是数字型
{alert("邮政编码号必须为数字!!");form1.postcode.focus();return;}
if(form1.telephone.value=="")
{alert("请输入联系电话!!");form1.telephone.focus();return;}
if(form1.email.value=="")
{alert("请输入E-mail!!");form1.email.focus();return;}
if(!checkemail(form1.email.value))
{alert("邮箱地址格式不正确，请重新输入!!");form1.email.focus();return;}
form1.submit();
}
</script>
<body>
<table width="800" height="520" border="0" align="center" cellpadding="0" cellspacing="0" background="images/new16.gif">
  <tr>
    <td height="86" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td width="46" height="289">&nbsp;</td>
    <td width="713" valign="top"><table width="693" height="70" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="693" height="70">
		<table width="670" height="294" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>
	<form action="" method="post" name="form1">
	<table width="581" height="285" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="125"><div align="center"><span class="STYLE2">用户名:</span></div></td>
        <td width="615"><input name="username" type="text" id="username" onKeyDown="if(event.keyCode==13){form1.name1.focus();}"></td>
      </tr>
      <tr>
        <td><div align="center" class="STYLE2">真实姓名:</div></td>
        <td><input name="name1" type="text" id="name1" onKeyDown="if(event.keyCode==13){form1.pwd.focus();}"></td>
      </tr>
      <tr>
        <td><div align="center" class="STYLE2">密&nbsp;&nbsp;&nbsp;码:</div></td>
        <td><input name="pwd" type="password" id="pwd" onKeyDown="if(event.keyCode==13){form1.pwd1.focus();}"></td>
      </tr>
      <tr>
        <td><div align="center" class="STYLE2">确认密码:</div></td>
        <td><input name="pwd1" type="password" id="pwd1" onKeyDown="if(event.keyCode==13){form1.city.focus();}"></td>
      </tr>
      <tr>
        <td><div align="center" class="STYLE2">所在城市:</div></td>
        <td><input name="city" type="text" id="city" onKeyDown="if(event.keyCode==13){form1.address.focus();}"></td>
      </tr>
      <tr>
        <td><div align="center" class="STYLE2">联系地址:</div></td>
        <td><input name="address" type="text" id="address" size="30" onKeyDown="if(event.keyCode==13){form1.postcode.focus();}"></td>
      </tr>
      <tr>
        <td><div align="center" class="STYLE2">邮政编码:</div></td>
        <td><input name="postcode" type="text" id="postcode" onKeyDown="if(event.keyCode==13){form1.telephone.focus();}"></td>
      </tr>
      <tr>
        <td><div align="center" class="STYLE2">联系电话:</div></td>
        <td><input name="telephone" type="text" id="telephone" onKeyDown="if(event.keyCode==13){form1.email.focus();}"></td>
      </tr>
      <tr>
        <td><div align="center" class="STYLE2">E-mail:</div></td>
        <td><input name="email" type="text" id="email" size="30" onKeyDown="if(event.keyCode==13){form1.Submit.focus();}"></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><input type="button" name="Submit" value="确定保存" onClick="Mycheck();">
          <input type="reset" name="Submit2" value="重新填写">
          <input type="button" name="Submit3" value="返回" onClick="JScript:window.location.href='index.asp';"></td>
      </tr>
    </table>
	</form>
	</td>
  </tr>
</table>
</td>
      </tr>
    </table></td>
    <td width="41">&nbsp;</td>
  </tr>
  <tr>
    <td height="13" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td height="127" colspan="3">&nbsp;</td>
  </tr>
</table>
</body>
</html>
