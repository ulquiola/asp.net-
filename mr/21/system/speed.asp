<%@ Language="VBScript" %>
<!--#include file="nums.asp" -->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>���̼��</TITLE>
<style>
<!--
BODY
{
	FONT-FAMILY: ����;
	FONT-SIZE: 9pt
}
TD
{
	FONT-SIZE: 9pt
}
A
{
	COLOR: #000000;
	TEXT-DECORATION: none
}
A:hover
{
	COLOR: #3F8805;
	TEXT-DECORATION: underline
}
.input
{
	BORDER: #111111 1px solid;
	FONT-SIZE: 9pt;
	BACKGROUND-color: #F8FFF0
}
.backs
{
	BACKGROUND-COLOR: #3F8805;
	COLOR: #ffffff;

}
.backq
{
	BACKGROUND-COLOR: #EEFEE0
}
.backc
{
	BACKGROUND-COLOR: #3F8805;
	BORDER: medium none;
	COLOR: #ffffff;
	HEIGHT: 18px;
	font-size: 9pt
}
.fonts
{
	COLOR: #3F8805
}
-->
</STYLE>
</HEAD>
<BODY>
<%
Response.Flush
if ObjTest("Scripting.FileSystemObject") then
	set fsoobj=server.CreateObject("Scripting.FileSystemObject")
%>
<table BORDER=1 align="center" CELLPADDING=4 CELLSPACING=0 bordercolorlight=#c0c0c0 bordercolordark=#FFFFFF width="500">
  <tr height=18  align=center>
    <td colspan="6">�����ļ������ٶȲ���</td>
  </tr>
  <tr height=18>
	<td colspan="2" align=left><%
Response.Flush
	' �����ļ���д
	Response.Write "�����ظ�������д���ɾ���ı��ļ�50��..."
	dim thetime3,tempfile,iserr
iserr=false
	t1=timer
	tempfile=server.MapPath("./") & "\aspchecktest.txt"
	for i=1 to 50
		Err.Clear

		set tempfileOBJ = FsoObj.CreateTextFile(tempfile,true)
		if Err <> 0 then
			Response.Write "�����ļ�����<br><br>"
			iserr=true
			Err.Clear
			exit for
		end if
		tempfileOBJ.WriteLine "Only for test. Ajiang ASPcheck"
		if Err <> 0 then
			Response.Write "д���ļ�����<br><br>"
			iserr=true
			Err.Clear
			exit for
		end if
		tempfileOBJ.close
		Set tempfileOBJ = FsoObj.GetFile(tempfile)
		tempfileOBJ.Delete 
		if Err <> 0 then
			Response.Write "ɾ���ļ�����<br><br>"
			iserr=true
			Err.Clear
			exit for
		end if
		set tempfileOBJ=nothing
	next
	t2=timer
if iserr <> true then
	thetime3=cstr(int(((t2-t1)*10000 )+0.5)/10)
	Response.Write "...����ɣ�ʱ��Ϊ<font color=red>"&thetime3 & "����</font>��<br>"
	Response.Flush
%></td>
  </tr>
  <tr height=18 align=center>
	<td width=320>�� ǰ �� �� ��</td>
	<td width=130>���ʱ��(����)</td>
  </tr>
  <tr height=18>
	<td align=center>&nbsp;<font color=red>��̨������: <%=Request.ServerVariables("SERVER_NAME")%></font>&nbsp;</td><td align="center">&nbsp;<font color=red><%=thetime3%></font></td>
  </tr>
</table>
<%
end if
Response.Flush
	set fsoobj=nothing
end if	
%>
</BODY>
</HTML>
