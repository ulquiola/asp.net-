<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�������������¼��Ϣ</title>
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
{//�����Զ��庯��
if(form1.aa.value=="")//�ж��ļ�·���Ƿ�Ϊ��
{alert("��������Ҫ�������ļ�·��");form1.aa.focus();return false}
//������ʾ�Ի���
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
    <input type="button" name="Submit" value="�ύ" onClick="mycheck();" >
  </p>
  <p class="STYLE1">&nbsp;&nbsp;&nbsp;&nbsp;ע�⣺·���ĸ�ʽΪ��c:/����c:/sml/��  </p>
</form>
	</td>
  </tr>
</table>
</body>
</html>
