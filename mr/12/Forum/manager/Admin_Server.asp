<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn/conn.asp" --><!--�������ݿ������ļ�-->
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
<title>�ޱ����ĵ�</title>
</head>
<BODY topMargin=0 leftmargin="0" marginheight="0"  bgcolor="B9DFFF">
 <p>&nbsp;</p>
  <table width="97%" border="1" align="center" cellpadding="3" cellspacing="0"  bordercolorlight="#4EAEE1" bordercolordark="#ffffff" bgcolor="#FFFFFF">
    <tr > 
      <td height=25 colspan="2" bgcolor="#4EAEE1"><b><font color="#FFFFFF">&nbsp;&nbsp;<span class="STYLE3"><font color="#FFFFFF"><strong>��ǰλ�ã�</strong></font>��������Ϣ</span></font></b></td>
    </tr>
    <tr> 
      <td width="30%" valign=middle> <span class="STYLE5">��ʾ�ͻ�����������HTTP���� </span></td>
      <td width="70%" class="STYLE5"><%=request.ServerVariables("All_Http")%></td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> ��ȡISAPIDLL��metabase·�� </td>
      <td width="70%" class="STYLE5"><%=request.ServerVariables("APPL_MD_PATH")%></td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> ��ʾվ������·�� </td>
      <td width="70%" class="STYLE5"><%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> ·����Ϣ </td>
      <td width="70%" class="STYLE5"><%=request.ServerVariables("PATH_INFO")%></td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> ��ʾ�������IP��ַ </td>
      <td width="70%" class="STYLE5"><%=request.ServerVariables("REMOTE_ADDR")%></td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> ������IP��ַ </td>
      <td width="70%" class="STYLE5"><%=Request.ServerVariables("LOCAL_ADDR")%></td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> ��ʾִ��SCRIPT������·�� </td>
      <td width="70%" class="STYLE5"><%=request.ServerVariables("SCRIPT_NAME")%></td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> ���ط���������������DNS��������IP��ַ </td>
      <td width="70%" class="STYLE5"><%=request.ServerVariables("SERVER_NAME")%></td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> ���ط�������������Ķ˿� </td>
      <td width="70%" class="STYLE5"><%=request.ServerVariables("SERVER_PORT")%></td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> Э������ƺͰ汾 </td>
      <td width="70%" class="STYLE5"><%=request.ServerVariables("SERVER_PROTOCOL")%></td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> �����������ƺͰ汾 </td>
      <td width="70%" class="STYLE5"><%=request.ServerVariables("SERVER_SOFTWARE")%>&nbsp;</td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> �ű���ʱʱ�� </td>
      <td width="70%" class="STYLE5"><%=Server.ScriptTimeout%>��</td>
    </tr>
    <tr> 
      <td width="30%" valign=top class="STYLE5"> �������������� </td>
      <td width="70%" class="STYLE5"><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
    </tr>
</table>
</body>
</html>
