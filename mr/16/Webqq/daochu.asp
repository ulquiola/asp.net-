<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>导出所有聊天记录信息</title>
<style type="text/css">
<!--
.STYLE1 {
	font-size: 9pt;
	color: #FF0000;
}
.STYLE2 {font-size: 9pt}
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
</head>
<script language="javascript">
function mycheck()
{//创建自定义函数
if(form1.aa.value=="")//判断文件路径是否为空
{alert("请输入所要导出的文件路径");form1.aa.focus();return false}
//弹出提示对话框
form1.submit();
}
</script>
<body>
<table width="303" height="152" border="0" cellpadding="0" cellspacing="0" background="Images/jl.gif">
  <tr>
    <td>
	<form name="form1" action="jjjjjj.asp" method="post">
  <br>
  &nbsp;&nbsp;&nbsp;&nbsp;
  <p>&nbsp;&nbsp;&nbsp;
    <input name="aa"  type="text" id="aa">
    <input type="button" name="Submit" value="提交" onClick="mycheck();" >
  </p>
  <p class="STYLE1">&nbsp;&nbsp;&nbsp;&nbsp;注意：路径的格式为“c:/”或“c:/sml/”  </p>
</form>
	</td>
  </tr>
</table>
</body>
</html>
