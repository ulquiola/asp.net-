<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<style type="text/css">
<!--
.STYLE1 {
	font-size: 9pt;
	color: #FF0000;
}
-->
</style>
</head>
<script language="javascript">
function mycheck()
{
if(form1.aa.value=="")
{alert("请输入所要导出的文件路径");form1.aa.focus();return false}
form1.submit();
}
</script>
<body>
<form name="form1" action="jjjjjj.asp" method="post">
  <p>
  <input name="aa"  type="text" id="aa">
  <input type="submit" name="Submit" value="提交">
</p>
  <p class="STYLE1">注意：路径的格式为“c:/”或“c:/sml/”  </p>
</form>
</body>
</html>
