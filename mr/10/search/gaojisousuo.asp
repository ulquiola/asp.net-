<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="shangbiao.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<script language="javascript">
function Mycheck()
{
if (form1.content.value=="")
{ alert("������Ҫ�����Ĳ������ݣ�");form1.content.focus();return false;}
form1.submit();
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��������</title>
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
<table width="779" height="206" border="0" align="center" cellpadding="��" cellspacing="��" bordercolor="#FFFFFF" bgcolor="#FFFFFF">
  <form name="form1" action="gaojicl.asp" method="post">
    <tr bordercolor="#bdcff7" bgcolor="#bdcff7">
      <td height="71" colspan="3" bordercolor="#FFFFFF" bgcolor="#FFFFFF"> ��������������
          <div align="center" class="style3">����Ҫ���������ݣ�
            <input name="content" type="text" id="content" style="height:23px " size="22">
          </div>
      <td width="136" height="71" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><p align="left">����
        <input name="imageField2" type="image" src="images/souan2.gif" width="77" height="20" border="0" onClick="return Mycheck();">
      </p>
      <td width="137" align="center" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><a href="index.asp" class="style3">������ҳ </a>
    <tr bordercolor="#bdcff7" bgcolor="#bdcff7"> </tr>
    <tr bordercolor="#bdcff7" bgcolor="#bdcff7">
      <td width="494" height="60" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><div align="center" class="style3">ѡ�����������ʾ��������
        <select name="select1" id="select1">
                <option value="10" selected>ÿҳ��ʾ10��</option>
                <option value="15">ÿҳ��ʾ15��</option>
                <option value="20">ÿҳ��ʾ20��</option>
                <option value="25">ÿҳ��ʾ25��</option>
                <option value="30">ÿҳ��ʾ30��</option>
              </select>
      </div></td>
    </tr>
    <tr bordercolor="#bdcff7" bgcolor="#bdcff7">
      <td align="center" valign="middle" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><div align="center" class="style3"> ѡ��Ҫ���������ͣ���
        <select name="type1" id="type1">
                <option>PHP</option>
                <option>JSP</option>
                <option>ASP Javascript</option>
                <option selected>ASP VBscript</option>
                <option>ASP.NET C#</option>
                <option>ASP.NET VB</option>
                <option>ColdFusion</option>
                <option>ColdFusion���</option>
              </select>
      </div></td>
    </tr>
    <tr bordercolor="#bdcff7" bgcolor="#bdcff7"> </tr>
  </form>
</table>
</body>
</html>
