<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
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
	FONT-FAMILY: 宋体;
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
    <td height="27" colspan=2 align=center bgcolor=#ffffff class=red_3>服务器的有关参数</td>
  </tr>
  <tr>
    <td width="26%" height="23">&nbsp;服务器名：</td>
    <td width="74%" height="23">&nbsp;
        <%response.write Request.ServerVariables("SERVER_NAME")%></td>
  </tr>
  <tr>
    <td height="23">&nbsp;服务器IP：</td>
    <td height="23">&nbsp;
        <%response.write Request.ServerVariables("LOCAL_ADDR")%></td>
  </tr>
  <tr>
    <td height="23">&nbsp;服务器端口：</td>
    <td height="23">&nbsp;
        <%response.write Request.ServerVariables("SERVER_PORT")%></td>
  </tr>
  <tr>
    <td height="23">&nbsp;服务器时间：</td>
    <td height="23">&nbsp;
        <%response.write now%></td>
  </tr>
  <tr>
    <td height="23">&nbsp;IIS版本：</td>
    <td height="23">&nbsp;
        <%response.write Request.ServerVariables("SERVER_SOFTWARE")%></td>
  </tr>
  <tr>
    <td height="23">&nbsp;服务器操作系统：</td>
    <td height="23">&nbsp;
        <%response.write Request.ServerVariables("OS")%></td>
  </tr>
  <tr>
    <td height="23">&nbsp;脚本超时时间：</td>
    <td height="23">&nbsp;
        <%response.write Server.ScriptTimeout%>
      秒</td>
  </tr>
  <tr>
    <td height="23">&nbsp;站点物理路径：</td>
    <td height="23">&nbsp;
        <%response.write request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
  </tr>
  <tr>
    <td height="23">&nbsp;服务器CPU数量：</td>
    <td height="23">&nbsp;
        <%response.write Request.ServerVariables("NUMBER_OF_PROCESSORS")%>
      个</td>
  </tr>
  <tr>
    <td height="23">&nbsp;服务器解译引擎：</td>
    <td height="23">&nbsp;
        <%response.write ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
  </tr>
  <tr>
    <td height="23" colspan=2 align=center bgcolor=#ffffff class=red_3>组件支持有关参数</td>
  </tr>
  <tr>
    <td height="23">&nbsp;数据库(ADO)支持：</td>
    <td height="23">&nbsp;
        <%if object_install("adodb.connection")=false then%>
      <font class=red><b>×</b></font> （不支持）
      <% else %>
      <b>√</b> （支持）
      <% end if %></td>
  </tr>
  <tr>
    <td height="23">&nbsp;FSO文本读写：</td>
    <td height="23">&nbsp;
        <%if object_install("scripting.filesystemobject")=false then%>
      <font class=red><b>×</b></font> （不支持）
      <% else %>
      <b>√</b> （支持）
      <% end if %></td>
  </tr>
  <tr>
    <td height="23">&nbsp;Stream文件流：</td>
    <td height="23">&nbsp;
        <%if object_install("Adodb.Stream")=false then%>
      <font class=red><b>×</b></font> （不支持）
      <% else %>
      <b>√</b> （支持）
      <% end if %></td>
  </tr>
  <tr>
    <td height="23">&nbsp;Jmail组件支持：</td>
    <td height="23">&nbsp;
        <%If object_install("JMail.SMTPMail")=false Then%>
      <font class=red><b>×</b></font> （不支持）
      <% else %>
      <b>√</b> （支持）
      <% end if %></td>
  </tr>
  <tr>
    <td height="23">&nbsp;CDONTS组件支持：</td>
    <td height="23">&nbsp;
        <%If object_install("CDONTS.NewMail")=false Then%>
      <font class=red><b>×</b></font> （不支持）
      <% else %>
      <b>√</b> （支持）
      <% end if %></td>
  </tr>
</table>
</body>
</html>
