<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="shangbiao.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<script language="javascript">
function Mycheck()
{
if (form1.content.value=="")
{ alert("请输入要搜索的部分内容！");form1.content.focus();return false;}
form1.submit();
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>搜索引擎</title>
<link href="style1.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body {
	background-color: #bdcff7;
	background-image: url(images/bg.gif);
}
.style3 {font-size: 9pt}
-->
</style>
</head>
<body onload="clockon()" background="bg.gif">
<table width="779" height="206" border="0" align="center" cellpadding="１" cellspacing="１" bordercolor="#FFFFFF" bgcolor="#FFFFFF">
  <form name="form1" action="gaojicl.asp" method="post">
    <tr bordercolor="#bdcff7" bgcolor="#bdcff7">
      <td height="71" colspan="3" bordercolor="#FFFFFF" bgcolor="#FFFFFF"> 　　　　　　　
          <div align="center" class="style3">输入要搜索的内容：
            <input name="content" type="text" id="content" style="height:23px " size="22">
          </div>
      <td width="136" height="71" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><p align="left">　　
        <input name="imageField2" type="image" src="images/souan2.gif" width="77" height="20" border="0" onClick="return Mycheck();">
      </p>
      <td width="137" align="center" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><a href="index.asp" class="style3">返回首页 </a>
    <tr bordercolor="#bdcff7" bgcolor="#bdcff7"> </tr>
    <tr bordercolor="#bdcff7" bgcolor="#bdcff7">
      <td width="494" height="60" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><div align="center" class="style3">选择搜索结果显示的条数：
        <select name="select1" id="select1">
                <option value="10" selected>每页显示10条</option>
                <option value="15">每页显示15条</option>
                <option value="20">每页显示20条</option>
                <option value="25">每页显示25条</option>
                <option value="30">每页显示30条</option>
              </select>
      </div></td>
    </tr>
    <tr bordercolor="#bdcff7" bgcolor="#bdcff7">
      <td align="center" valign="middle" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><div align="center" class="style3"> 选择要搜索的类型：　
        <select name="type1" id="type1">
                <option>PHP</option>
                <option>JSP</option>
                <option>ASP Javascript</option>
                <option selected>ASP VBscript</option>
                <option>ASP.NET C#</option>
                <option>ASP.NET VB</option>
                <option>ColdFusion</option>
                <option>ColdFusion组件</option>
              </select>
      </div></td>
    </tr>
    <tr bordercolor="#bdcff7" bgcolor="#bdcff7"> </tr>
  </form>
</table>
</body>
</html>
