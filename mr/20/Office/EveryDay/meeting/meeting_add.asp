<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../../Conn/conn.asp" -->
<%
if request.Form("subject")<>"" then
	mtime=request.Form("mtime")
	CPerson=request.Form("CPerson")
	ZPerson=request.Form("ZPerson")
	address=request.Form("address")
	subject=request.Form("subject")
	content=request.Form("content")
	Ins="Insert into Tab_meeting (mtime,CPerson,ZPerson,address,subject,content) values('"&_
	mtime&"','"&CPerson&"','"&ZPerson&"','"&address&"','"&subject&"','"&content&"')"
	conn.execute(Ins)%>
	<script language="javascript">
	alert("���ݱ���ɹ���");
	opener.location.reload();  
	//ˢ�¸������е���ҳ
	window.close();
	//�رյ�ǰ������
	</script>
<% end if%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>¼������¼</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<script language="javascript">
function Mycheck()
{
if (form1.mtime.value=="")
{ alert("���������ʱ�䣡");form1.mtime.focus();return;}
  if(form1.mtime.value.length!=16)
    {alert("����������ڲ��ԣ��������׼���ڣ��磺2004/07/27 07:30��2004-07-27 07:30��");form1.mtime.focus();return;}
  if(form1.mtime.value.indexOf('/',0)==-1&form1.mtime.value.indexOf('-',0)==-1)
    {alert("����������ڲ��ԣ��������׼���ڣ��磺2004/07/27 07:30��2004-07-27 07:30��");form1.mtime.focus();return;}
	  if(form1.mtime.value.substr(4,1)!="-"&form1.mtime.value.substr(4,1)!="/")
    {alert("����������ڲ��ԣ��������׼���ڣ��磺2004/07/27 07:30��2004-07-27 07:30��");form1.mtime.focus();return;}
  if(form1.mtime.value.substr(7,1)!="-"&form1.mtime.value.substr(7,1)!="/")
    {alert("����������ڲ��ԣ��������׼���ڣ��磺2004/07/27 07:30��2004-07-27 07:30��");form1.mtime.focus();return;}
  if(form1.mtime.value.substr(10,1)!=" ")
    {alert("����������ڲ��ԣ��������׼���ڣ��磺2004/07/27 07:30��2004-07-27 07:30��");form1.mtime.focus();return;}
  if(form1.mtime.value.substr(13,1)!=":")
    {alert("����������ڲ��ԣ��������׼���ڣ��磺2004/07/27 07:30��2004-07-27 07:30��");form1.mtime.focus();return;}
  if(form1.mtime.value.substr(5,2)>"12")
    {alert("����������ڲ��ԣ�");form1.mtime.focus();return;}
  if(form1.mtime.value.substr(8,2)>"31")
    {alert("����������ڲ��ԣ�");form1.mtime.focus();return;}
  if(form1.mtime.value.substr(11,2)>"23"||form1.mtime.value.substr(11,2)<"0")
    {alert("����������ڲ��ԣ�");form1.mtime.focus();return;}	
  if(form1.mtime.value.substr(14,2)>"59"||form1.mtime.value.substr(11,2)<"0")
    {alert("����������ڲ��ԣ�");form1.mtime.focus();return;}		
if (form1.ZPerson.value=="")
{ alert("��������������ˣ�");form1.ZPerson.focus();return;}
if (form1.CPerson.value=="")
{ alert("����������ϯ�ˣ�");form1.CPerson.focus();return;}
if (form1.subject.value=="")
{ alert("�����빫�����⣡");form1.subject.focus();return;}
if (form1.content.value=="")
{ alert("�����빫�����ݣ�");form1.content.focus();return;}
form1.submit();
}
</script>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
.style1 {color: #C60001}
.style2 {color: #669999}
.STYLE3 {
	font-size: 9pt;
	color: #FF0000;
}
.STYLE5 {
	font-size: 9pt;
	color: #FFFFFF;
}
.STYLE6 {font-size: 9pt; color: #000000; }
-->
</style></head>

<body>
<table width="557" height="401" border="0" cellpadding="-2" cellspacing="-2" background="../../Images/detail0.gif">
  <tr>
    <td width="817" height="349" valign="top"><table width="100%" height="93"  border="0" cellpadding="-2" cellspacing="-2">
      <tr>
        <td height="43" valign="bottom">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="STYLE5">¼������¼</span></td>
      </tr>
      <tr>
        <td height="50">&nbsp;</td>
      </tr>
    </table>      
      <form ACTION="meeting_add.asp" METHOD="POST" name="form1">
        <table width="93%" height="232"  border="0" align="center" cellpadding="-2" cellspacing="-2">
          <tr>
            <td width="11%" height="27"><div align="center" class="STYLE6">ʱ&nbsp;&nbsp;�䣺</div></td>
            <td width="34%"><input name="mtime" type="text" class="Sytle_text" id="mtime"></td>
            <td width="14%" class="STYLE6">��ϯ�ˣ�</td>
            <td width="41%"><input name="CPerson" type="text" class="Sytle_text" id="CPerson" size="30"></td>
          </tr>
          <tr>
            <td height="25"><span class="STYLE6">�����ˣ�</span></td>
            <td><input name="ZPerson" type="text" class="Sytle_text" id="ZPerson"></td>
            <td class="STYLE6">����ص㣺</td>
            <td><input name="address" type="text" class="Sytle_text" id="address" size="30"></td>
          </tr>
          <tr>
            <td height="27"><div align="center"><span class="STYLE6">��&nbsp;&nbsp;�⣺</span></div></td>
            <td colspan="3"><input name="subject" type="text" class="Style_subject" id="subject"></td>
          </tr>
          <tr>
            <td height="133" valign="top"><div align="center" class="STYLE6">��&nbsp;&nbsp;�ݣ�</div></td>
            <td colspan="3"><textarea name="content" cols="53" rows="6" class="Style_content" id="content"></textarea></td>
          </tr>
          <tr>
            <td colspan="4"><div align="center">
                <input name="Button" type="button" class="Style_button" value="�ύ" onClick="Mycheck()">
                <input name="Submit2" type="reset" class="Style_button" value="����">                
                <input name="myclose" type="button" class="Style_button" id="myclose" value="�ر�"
				 onClick="javascrip:window.close()">
                </div></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
</body>
</html>
