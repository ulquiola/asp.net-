<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn/conn.asp" --><!--包含数据库连接文件-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../style.css">
<html>
<style type="text/css">
<!--
.STYLE3 {font-size: 10pt}
.STYLE5 {font-size: 10pt; color: #000000; }
-->
</style>
<head>
<title>无标题文档</title>
</head>
<BODY topMargin=0 leftmargin="0" marginheight="0"  bgcolor="B9DFFF">
 <p>&nbsp;</p>
  <table width="97%" border="1" align="center" cellpadding="3" cellspacing="0"  bordercolorlight="#4EAEE1" bordercolordark="#ffffff" bgcolor="#FFFFFF">
    <tr > 
      <td height=25 colspan="2" bgcolor="#4EAEE1"><b><font color="#FFFFFF">&nbsp;&nbsp;<span class="STYLE3"><font color="#FFFFFF"><strong>当前位置：</strong></font>服务器信息</span></font></b></td>
    </tr>
    <tr> 
      <td width="30%" valign=middle> <span class="STYLE5">显示客户发出的所有HTTP标题 </span></td>
      <td width="70%" class="STYLE5"><%=request.ServerVariables("All_Http")%></td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> 检取ISAPIDLL的metabase路径 </td>
      <td width="70%" class="STYLE5"><%=request.ServerVariables("APPL_MD_PATH")%></td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> 显示站点物理路径 </td>
      <td width="70%" class="STYLE5"><%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> 路径信息 </td>
      <td width="70%" class="STYLE5"><%=request.ServerVariables("PATH_INFO")%></td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> 显示请求机器IP地址 </td>
      <td width="70%" class="STYLE5"><%=request.ServerVariables("REMOTE_ADDR")%></td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> 服务器IP地址 </td>
      <td width="70%" class="STYLE5"><%=Request.ServerVariables("LOCAL_ADDR")%></td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> 显示执行SCRIPT的虚拟路径 </td>
      <td width="70%" class="STYLE5"><%=request.ServerVariables("SCRIPT_NAME")%></td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> 返回服务器的主机名，DNS别名，或IP地址 </td>
      <td width="70%" class="STYLE5"><%=request.ServerVariables("SERVER_NAME")%></td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> 返回服务器处理请求的端口 </td>
      <td width="70%" class="STYLE5"><%=request.ServerVariables("SERVER_PORT")%></td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> 协议的名称和版本 </td>
      <td width="70%" class="STYLE5"><%=request.ServerVariables("SERVER_PROTOCOL")%></td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> 服务器的名称和版本 </td>
      <td width="70%" class="STYLE5"><%=request.ServerVariables("SERVER_SOFTWARE")%>&nbsp;</td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> 脚本超时时间 </td>
      <td width="70%" class="STYLE5"><%=Server.ScriptTimeout%>秒</td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> 服务器解译引擎 </td>
      <td width="70%" class="STYLE5"><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
    </tr>
</table>
</body>
</html>
