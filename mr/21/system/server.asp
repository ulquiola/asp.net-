<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ޱ����ĵ�</title>
<%
function object_install(strclassstring)
  on error resume next
  object_install=false
  dim xtestobj
  err=0
  set xtestobj=server.createobject(strclassstring)
  if err=0 then object_install=true
  set xtestobj=nothing
  err=0
end function
%> 
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

</head>

<body>
<table width='100%' height="384" border=1 cellpadding=1 cellspacing=0 bordercolorlight=#c0c0c0 bordercolordark=#FFFFFF>
  <tr>
    <td height="27" colspan=2 align=center bgcolor=#ffffff class=red_3>���������йز���</td>
  </tr>
  <tr>
    <td width="26%" height="23">&nbsp;����������</td>
    <td width="74%" height="23">&nbsp;
        <%response.write Request.ServerVariables("SERVER_NAME")%></td>
  </tr>
  <tr>
    <td height="23">&nbsp;������IP��</td>
    <td height="23">&nbsp;
        <%response.write Request.ServerVariables("LOCAL_ADDR")%></td>
  </tr>
  <tr>
    <td height="23">&nbsp;�������˿ڣ�</td>
    <td height="23">&nbsp;
        <%response.write Request.ServerVariables("SERVER_PORT")%></td>
  </tr>
  <tr>
    <td height="23">&nbsp;������ʱ�䣺</td>
    <td height="23">&nbsp;
        <%response.write now%></td>
  </tr>
  <tr>
    <td height="23">&nbsp;IIS�汾��</td>
    <td height="23">&nbsp;
        <%response.write Request.ServerVariables("SERVER_SOFTWARE")%></td>
  </tr>
  <tr>
    <td height="23">&nbsp;����������ϵͳ��</td>
    <td height="23">&nbsp;
        <%response.write Request.ServerVariables("OS")%></td>
  </tr>
  <tr>
    <td height="23">&nbsp;�ű���ʱʱ�䣺</td>
    <td height="23">&nbsp;
        <%response.write Server.ScriptTimeout%>
      ��</td>
  </tr>
  <tr>
    <td height="23">&nbsp;վ������·����</td>
    <td height="23">&nbsp;
        <%response.write request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
  </tr>
  <tr>
    <td height="23">&nbsp;������CPU������</td>
    <td height="23">&nbsp;
        <%response.write Request.ServerVariables("NUMBER_OF_PROCESSORS")%>
      ��</td>
  </tr>
  <tr>
    <td height="23">&nbsp;�������������棺</td>
    <td height="23">&nbsp;
        <%response.write ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
  </tr>
  <tr>
    <td height="23" colspan=2 align=center bgcolor=#ffffff class=red_3>���֧���йز���</td>
  </tr>
  <tr>
    <td height="23">&nbsp;���ݿ�(ADO)֧�֣�</td>
    <td height="23">&nbsp;
        <%if object_install("adodb.connection")=false then%>
      <font class=red><b>��</b></font> ����֧�֣�
      <% else %>
      <b>��</b> ��֧�֣�
      <% end if %></td>
  </tr>
  <tr>
    <td height="23">&nbsp;FSO�ı���д��</td>
    <td height="23">&nbsp;
        <%if object_install("scripting.filesystemobject")=false then%>
      <font class=red><b>��</b></font> ����֧�֣�
      <% else %>
      <b>��</b> ��֧�֣�
      <% end if %></td>
  </tr>
  <tr>
    <td height="23">&nbsp;Stream�ļ�����</td>
    <td height="23">&nbsp;
        <%if object_install("Adodb.Stream")=false then%>
      <font class=red><b>��</b></font> ����֧�֣�
      <% else %>
      <b>��</b> ��֧�֣�
      <% end if %></td>
  </tr>
  <tr>
    <td height="23">&nbsp;Jmail���֧�֣�</td>
    <td height="23">&nbsp;
        <%If object_install("JMail.SMTPMail")=false Then%>
      <font class=red><b>��</b></font> ����֧�֣�
      <% else %>
      <b>��</b> ��֧�֣�
      <% end if %></td>
  </tr>
  <tr>
    <td height="23">&nbsp;CDONTS���֧�֣�</td>
    <td height="23">&nbsp;
        <%If object_install("CDONTS.NewMail")=false Then%>
      <font class=red><b>��</b></font> ����֧�֣�
      <% else %>
      <b>��</b> ��֧�֣�
      <% end if %></td>
  </tr>
</table>
</body>
</html>
