<%@ Language="VBScript" %>
<% Option Explicit %>
<!--#include file="nums.asp" -->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>ϵͳ���</TITLE>
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
<table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center"><font class=fonts>���֧��������</font></td>
  </tr>
</table>
<br>
<table width="500" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#3F8805" style="border-collapse: collapse">
  <tr height=18  align=center>
    <td colspan="2">��������������������Ҫ���������ProgId��ClassId��</td>
  </tr>
  <FORM action=<%=Request.ServerVariables("SCRIPT_NAME")%> method=post id=form1 name=form1>
    <tr height="18">
      <td align=center height=30><input class=input type=text value="" name="classname" size=40>
        <INPUT type=submit value=" ȷ �� " class=backc id=submit1 name=submit1>
        <INPUT type=reset value=" �� �� " class=backc id=reset1 name=reset1>      </td>
    </tr>
  </FORM>
</table>
<br>
<table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center"><font class=fonts>���֧�����</font>
      <%
Dim strClass
	strClass = Trim(Request.Form("classname"))
	If "" <> strClass then
	Response.Write "<br>��ָ��������ļ������"
	Dim Verobj1
	ObjTest(strClass)
	  If Not IsObj then 
		Response.Write "<br><font color=red>���ź����÷�������֧�� " & strclass & " �����</font>"
	  Else
		if VerObj="" or isnull(VerObj) then 
			Verobj1="�޷�ȡ�ø�����汾"
		Else
			Verobj1="������汾�ǣ�" & VerObj
		End If
		Response.Write "<br><font class=fonts>��ϲ���÷�����֧�� " & strclass & " �����" & verobj1 & "</font>"
	  End If
	  Response.Write "<br>"
	end if
	%></td>
  </tr>
</table>
<br>
<table width="500" border=1 align="center" cellpadding=1 cellspacing=0 bordercolorlight=#c0c0c0 bordercolordark=#FFFFFF>
  <tr height=18  align=center>
    <td colspan="2">IIS�Դ���ASP���</td>
  </tr>
  <tr height=18  align=center>
    <td width=320>�� �� �� ��</td>
    <td width=130>֧�ּ��汾</td>
  </tr>
  <%For i=0 to 10%>
  <tr height="18">
    <td align=left>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
    <td align=left>&nbsp;
      <%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>��</b></font>"
		Else
			Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%></td>
  </tr>
  <%next%>
</table>
<br>
<table width="500" border=1 align="center" cellpadding=1 cellspacing=0 bordercolorlight=#c0c0c0 bordercolordark=#FFFFFF>
  <tr height=18  align=center>
    <td colspan="2">�������ļ��ϴ��͹������</td>
  </tr>
  <tr height=18 align=center>
    <td width=320>�� �� �� ��</td>
    <td width=130>֧�ּ��汾</td>
  </tr>
  <%For i=11 to 15%>
  <tr height="18">
    <td align=left>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
    <td align=left>&nbsp;
      <%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>��</b></font>"
		Else
			Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%></td>
  </tr>
  <%next%>
</table>
<br>
<table width="500" border=1 align="center" cellpadding=1 cellspacing=0 bordercolorlight=#c0c0c0 bordercolordark=#FFFFFF>
  <tr height=18  align=center>
    <td colspan="2">�������շ��ʼ����</td>
  </tr>
  <tr height=18 align=center>
    <td width=320>�� �� �� ��</td>
    <td width=130>֧�ּ��汾</td>
  </tr>
  <%For i=16 to 23%>
  <tr height="18" >
    <td align=left>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
    <td align=left>&nbsp;
      <%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>��</b></font>"
		Else
			Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%></td>
  </tr>
  <%next%>
</table>
<br>
<table width="500" border=1 align="center" cellpadding=1 cellspacing=0 bordercolorlight=#c0c0c0 bordercolordark=#FFFFFF>
  <tr height=18  align=center>
    <td colspan="2">ͼ�������</td>
  </tr>
  <tr height=18 align=center>
    <td width=320>�� �� �� ��</td>
    <td width=130>֧�ּ��汾</td>
  </tr>
  <%For i=24 to 25%>
  <tr height="18">
    <td align=left>&nbsp;<%=ObjTotest(i,0) & "<font color=#888888>&nbsp;" & ObjTotest(i,1)%></font></td>
    <td align=left>&nbsp;
      <%
		If Not ObjTotest(i,2) Then 
			Response.Write "<font color=red><b>��</b></font>"
		Else
			Response.Write "<font class=fonts><b>��</b></font> <a title='" & ObjTotest(i,3) & "'>" & left(ObjTotest(i,3),11) & "</a>"
		End If%></td>
  </tr>
  <%next%>
</table>
</BODY>
</HTML>
