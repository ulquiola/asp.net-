<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" --><!--�������ݿ������ļ�-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>����Ա��¼</title>
<link rel="stylesheet" type="text/css" href="../css/style.css">
<style type="text/css">
<!--
.STYLE1 {
	font-size: 9pt;
	color: #000000;
}
.STYLE15 {color: #000066; font-size: 12px; }
-->
</style>
</head>
<script language="javascript">
function mycheck(myform,tool)//�����Զ��庯��
{
if (myform.UID.value=="")//�ж��û����ı����Ƿ���ֵ
{alert("������"+tool+"��");myform.UID.focus();return;}//������ʾ�Ի���
if(myform.PWD.value=="")//�ж������ı����Ƿ���ֵ
{alert("���������룡");myform.PWD.focus();return;}		//������ʾ�Ի���
if(myform.yanzheng.value=="")//�ж���֤���ı����Ƿ���ֵ
{alert("��������֤��!");myform.yanzheng.focus();return;}//������ʾ�Ի���
if(myform.yanzheng.value!=myform.verifycode2.value)//�ж������������֤���Ƿ���ͬ
{alert("��������ȷ����֤��!!");myform.yanzheng.focus();return;}//������ʾ�Ի���
myform.submit();//�ύ��
}
</script>
<script src="../js/fun.js"></script>
<body bgcolor="#B9DFFF">
<%   
    '�Զ�����4λ�������֤��
	randomize
	i=0					'Ϊ����i��ֵ
	do while i<4		
	num1=int(9*rnd)		'�����ɵ�����ȡ��
	numimage="<img src=../images/num/"&num1&".gif>" '��ȡͼƬ
	numi=numi&numimage	'Ϊ����numi��ֵ
	num=num&num1		'Ϊ����num��ֵ
	i=i+1				'iѭ����1
	loop
	session("numi")=numi	'Ϊsession("numi")������ֵ
	session("num")=num		'Ϊsession("num")������ֵ
%>
<table width="500" height="400" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td background="../images/login.gif"><table width="350" height="200" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="54">&nbsp;</td>
        <td width="296">
		<table width="259" height="139"  border="0" align="center" cellpadding="-2" cellspacing="-2">
  <tr>
    <td height="139" align="center"><form name="form_M" method="post" action="Login_M.asp">
        <table width="262" height="120" border="0" align="center" cellpadding="-2" cellspacing="-2" class="tabBorder">
          <tr align="center" valign="middle">
            <td height="24" colspan="2" valign="bottom" background="Images/bg_Login.GIF"><font color="#505875">=== ����Ա��¼ === </font> </td>
          </tr>
          <tr>
            <td width="95" height="27" align="right" class="STYLE1">&nbsp;&nbsp;&nbsp;����Ա���ƣ�</td>
            <td width="167"><div align="left">
              <input name="UID" type="text" class="inputcss1" maxlength="20">
            </div></td>
          </tr>
          <tr>
            <td height="16" align="right" valign="middle" class="STYLE1">��&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;�룺</td>
            <td valign="middle"><input name="PWD" type="password" class="inputcss1" onKeyDown="if(event.keyCode==13) mycheck(form_M,'��������')"  maxlength="20"></td>
          </tr>
          <tr>
            <td height="13" align="right" valign="middle" class="STYLE1"><span class="STYLE1">��&nbsp;&nbsp;֤&nbsp;&nbsp;�룺</span></td>
            <td valign="middle"><span class="STYLE16">
              <input name="yanzheng" type="text" class="inputcss1" id="yanzheng" size="10">
			  <input type="hidden" name="verifycode2" value="<%=session("num")%>"><%=session("numi")%>
            </span></td>
          </tr>
          <tr align="center" valign="middle">
            <td height="27" colspan="2"><input name="Submit" type="button" class="btn_grey" value="��¼" onClick="mycheck(form_M,'��������')">
&nbsp;
            <input name="Submit2" type="reset" class="btn_grey" value="����"></td>
          </tr>
        </table>
    </form></td>
  </tr>
</table>
		
		</td>
      </tr>
    </table></td>
  </tr>
</table>
<table width="500" height="30" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><div align="center" class="STYLE1" style=" line-height:2">��Ȩ���У�����ʡ���տƼ����޹�˾<br>Copyright&nbsp;&copy;&nbsp;www.mingrisoft.com All Rights Reserved!
</div></td>
  </tr>
</table>
</body>
</html>
