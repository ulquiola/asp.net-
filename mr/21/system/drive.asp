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

<br>
<table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="18" align="center"><font class=fonts>������ؼ��</font></td>
  </tr>
</table>
<table BORDER=1 align="center" CELLPADDING=4 CELLSPACING=0 bordercolorlight=#c0c0c0 bordercolordark=#FFFFFF width="500">
  <tr height=18  align=center>
    <td colspan="6">������������Ϣ</td>
  </tr>
  <tr height="18" align=center>
	<td width="100">�̷��ʹ�������</td>
	<td width="50">����</td>
	<td width="80">���</td>
	<td width="60">�ļ�ϵͳ</td>
	<td width="80">���ÿռ�</td>
	<td width="80">�ܿռ�</td>
  </tr>
<%

	' ���Դ�����Ϣ
	
	set drvObj=fsoobj.Drives
	for each d in drvObj
%>
  <tr height="18" align=center>
	<td align="right"><%=cdrivetype(d.DriveType) & " " & d.DriveLetter%>:</td>
<%
	if d.DriveLetter = "A" then	'Ϊ��ֹӰ������������������
		Response.Write "<td></td><td></td><td></td><td></td><td></td>"
	else
%>
	<td><%=cIsReady(d.isReady)%></td>
	<td>&nbsp;<%=d.VolumeName%></td>
	<td><%=d.FileSystem%></td>
	<td align="right"><%=cSize(d.FreeSpace)%></td>
	<td align="right"><%=cSize(d.TotalSize)%></td>
<%
	end if
%>
  </tr>
<%
	next
%>
</table>

<br>
<table BORDER=1 align="center" CELLPADDING=4 CELLSPACING=0 bordercolorlight=#c0c0c0 bordercolordark=#FFFFFF width="500">
  <tr height=18  align=center>
    <td colspan="5">��ǰ�ļ�����Ϣ</td>
  </tr>
  <td colspan="5" align="center"><%
	Response.Flush
	dPath = server.MapPath("./")
	set dDir = fsoObj.GetFolder(dPath)
	set dDrive = fsoObj.GetDrive(dDir.Drive)
%>
�ļ���: <%=dPath%></td>
  </tr>
  <tr height="18" align="center">
	<td width="75">���ÿռ�</td>
	<td width="75">���ÿռ�</td>
	<td width="75">�ļ�����</td>
	<td width="75">�ļ���</td>
	<td width="150">����ʱ��</td>
  </tr>
  <tr height="18" align="center">
	<td><%=cSize(dDir.Size)%></td>
	<td><%=cSize(dDrive.AvailableSpace)%></td>
	<td><%=dDir.SubFolders.Count%></td>
	<td><%=dDir.Files.Count%></td>
	<td><%=dDir.DateCreated%></td>
  </tr>
</table>

<%
end if%>
</BODY>
</HTML>
<%
function cdrivetype(tnum)
    Select Case tnum
        Case 0: cdrivetype = "δ֪"
        Case 1: cdrivetype = "���ƶ�����"
        Case 2: cdrivetype = "����Ӳ��"
        Case 3: cdrivetype = "�������"
        Case 4: cdrivetype = "CD-ROM"
        Case 5: cdrivetype = "RAM ����"
    End Select
end function

function cIsReady(trd)
    Select Case trd
		case true: cIsReady="<font class=fonts><b>��</b></font>"
		case false: cIsReady="<font color='red'><b>��</b></font>"
	End Select
end function

function cSize(tSize)
    if tSize>=1073741824 then
		cSize=int((tSize/1073741824)*1000)/1000 & " GB"
    elseif tSize>=1048576 then
    	cSize=int((tSize/1048576)*1000)/1000 & " MB"
    elseif tSize>=1024 then
		cSize=int((tSize/1024)*1000)/1000 & " KB"
	else
		cSize=tSize & "B"
	end if
end function
%>