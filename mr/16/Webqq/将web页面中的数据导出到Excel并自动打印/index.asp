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
{alert("��������Ҫ�������ļ�·��");form1.aa.focus();return false}
form1.submit();
}
</script>
<body>
<form name="form1" action="jjjjjj.asp" method="post">
  <p>
  <input name="aa"  type="text" id="aa">
  <input type="submit" name="Submit" value="�ύ">
</p>
  <p class="STYLE1">ע�⣺·���ĸ�ʽΪ��c:/����c:/sml/��  </p>
</form>
</body>
</html>
