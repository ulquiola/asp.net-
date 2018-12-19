<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>网上客户管理系统--管理员登录！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	background-image:  url(../Images/bg.gif);
}
.style2 {color: #a2bcc5}
.style3 {color: #990000}
-->
</style></head>

<body>
<script language="javascript">
function mycheck(){
if (form1.Name.value=="")
{alert("请输入管理员姓名！");form1.Name.focus();return;}
if(form1.PWD.value=="")
{alert("请输入管理员密码！");form1.PWD.focus();return;}
form1.submit();
}
</script>
<table width="400" height="242" border="0" cellpadding="-2" cellspacing="-2" background="../Images/add_manager.gif">
  <tr>
    <td valign="top"><table width="400" height="271" cellpadding="-2" cellspacing="-2">
      <tr>
        <td width="11" height="85">&nbsp;</td>
        <td width="373">&nbsp;</td>
        <td width="14"><span class="style2"></span></td>
      </tr>
      <tr>
        <td height="144">&nbsp;</td>
        <td valign="top">
  
          <div align="center" class="style3">
            <table width="80%" height="70" cellpadding="-2"  cellspacing="-2">
              <tr>
                <td><div align="center">该管理员姓名已经存在，请点击[返回]按钮重新添加！</div></td>
              </tr>
            </table>
            <table width="80%"  cellspacing="-2" cellpadding="-2">
              <tr>
                <td><div align="center">
                  <input name="Submit" type="submit" class="Style_button" value="返回" onClick="JScript:window.location='add_manager.asp'">
                </div></td>
              </tr>
            </table>
          </div></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
