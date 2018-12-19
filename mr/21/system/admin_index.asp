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
<link href="Img/admin.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="100%">
<TR> 
<TD ALIGN=right height="5"></TD>
</TR>
<TR> 
<TD ALIGN=center HEIGHT=22><b>欢迎管理员：<font color="#FF0000"><%=Session(EwManageUser)%></font> 使用本系统</b></TD>
</TR>
<TR> 
<TD ALIGN=center HEIGHT=18><b> 
</b></TD>
</TR>
<TR> 
<TD> 
<table width=570 height='100%' border=0 align=center cellpadding="0" cellspacing="0">
<tr height='100%' align=center><td width='100%'>
  <TABLE width="100%" height="52" BORDER=1 align="center" CELLPADDING=4 CELLSPACING=0 bordercolorlight=#c0c0c0 bordercolordark=#FFFFFF>
    <TR>
      <TD height="27" colspan="2" align="center">版权信息</TD>
    </TR>
    <TR>
      <TD width="26%" height="23"><div align="left">版权所有：</div></TD>
      <TD width="74%" height="23"><%=Web_Name%> @2008-2010</TD>
    </TR>
  </TABLE>
  <br>
  <table width='100%' height="384" border=1 cellpadding=1 cellspacing=0 bordercolorlight=#c0c0c0 bordercolordark=#FFFFFF>
  <tr><td height="27" colspan=2 align=center bgcolor=#ffffff class=red_3>服务器的有关参数</td>
  </tr>
  <tr><td width="26%" height="23">&nbsp;服务器名：</td>
  <td width="74%" height="23">&nbsp;
    <%response.write Request.ServerVariables("SERVER_NAME")%></td></tr>
  <tr><td height="23">&nbsp;服务器IP：</td>
  <td height="23">&nbsp;
    <%response.write Request.ServerVariables("LOCAL_ADDR")%></td></tr>
  <tr><td height="23">&nbsp;服务器端口：</td>
  <td height="23">&nbsp;
    <%response.write Request.ServerVariables("SERVER_PORT")%></td></tr>
  <tr><td height="23">&nbsp;服务器时间：</td>
  <td height="23">&nbsp;
    <%response.write now%></td></tr>
  <tr><td height="23">&nbsp;IIS版本：</td>
  <td height="23">&nbsp;
    <%response.write Request.ServerVariables("SERVER_SOFTWARE")%></td></tr>
  <tr><td height="23">&nbsp;服务器操作系统：</td>
  <td height="23">&nbsp;
    <%response.write Request.ServerVariables("OS")%></td></tr>
  <tr><td height="23">&nbsp;脚本超时时间：</td>
  <td height="23">&nbsp;
    <%response.write Server.ScriptTimeout%> 秒</td></tr>
  <tr><td height="23">&nbsp;站点物理路径：</td>
  <td height="23">&nbsp;
    <%response.write request.ServerVariables("APPL_PHYSICAL_PATH")%></td></tr>
  <tr><td height="23">&nbsp;服务器CPU数量：</td>
  <td height="23">&nbsp;
    <%response.write Request.ServerVariables("NUMBER_OF_PROCESSORS")%> 个</td></tr>
  <tr><td height="23">&nbsp;服务器解译引擎：</td>
  <td height="23">&nbsp;
    <%response.write ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td></tr>
  <tr><td height="23" colspan=2 align=center bgcolor=#ffffff class=red_3>组件支持有关参数</td>
  </tr>
  <tr><td height="23">&nbsp;数据库(ADO)支持：</td>
  <td height="23">&nbsp;
    <%if object_install("adodb.connection")=false then%><font class=red><b>×</b></font> （不支持）<% else %><b>√</b> （支持）<% end if %></td></tr>
  <tr><td height="23">&nbsp;FSO文本读写：</td>
  <td height="23">&nbsp;
    <%if object_install("scripting.filesystemobject")=false then%><font class=red><b>×</b></font> （不支持）<% else %><b>√</b> （支持）<% end if %></td></tr>
  <tr><td height="23">&nbsp;Stream文件流：</td>
  <td height="23">&nbsp;
    <%if object_install("Adodb.Stream")=false then%><font class=red><b>×</b></font> （不支持）<% else %><b>√</b> （支持）<% end if %></td></tr>
  <tr><td height="23">&nbsp;Jmail组件支持：</td>
  <td height="23">&nbsp;
    <%If object_install("JMail.SMTPMail")=false Then%><font class=red><b>×</b></font> （不支持）<% else %><b>√</b> （支持）<% end if %></td></tr>
  <tr><td height="23">&nbsp;CDONTS组件支持：</td>
  <td height="23">&nbsp;
    <%If object_install("CDONTS.NewMail")=false Then%><font class=red><b>×</b></font> （不支持）<% else %><b>√</b> （支持）<% end if %></td></tr>
  </table>
</td></tr>
</table>
<br>
</TD>
</TR>
</TABLE>
