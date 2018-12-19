<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../../Conn/conn.asp" -->
<%
Set rs = Server.CreateObject("ADODB.Recordset")
sql="SELECT mTime,ZPerson,CPerson,subject,address,content FROM Tab_meeting WHERE ID="&_
Request.QueryString("ID")&""
rs.open sql,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>会议记录详细信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
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
	color: #000000;
}
.STYLE5 {
	font-size: 9pt;
	color: #FFFFFF;
}
-->
</style></head>

<body>
<table width="557" height="401" border="0" cellpadding="-2" cellspacing="-2" background="../../Images/detail0.gif">
  <tr>
    <td width="817" height="349" valign="top"><table width="100%" height="90"  border="0" cellpadding="-2" cellspacing="-2">
      <tr>
        <td height="43" valign="bottom">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="STYLE5">会议记录详细信息</span></td>
      </tr>
      <tr>
        <td height="47">&nbsp;</td>
      </tr>
    </table>      
      <form name="form1">
        <table width="93%" height="232"  border="0" align="center" cellpadding="-2" cellspacing="-2">
          <tr>
            <td width="12%" height="27"><div class="style1 STYLE3">时&nbsp;&nbsp;间：</div></td>
            <td width="34%"><input name="mtime" type="text" class="Sytle_text" id="mtime"
			 readonly value="<%=rs("mTime")%>"></td>
            <td width="13%" class="STYLE3">出席人：</td>
            <td width="41%"><input name="CPerson" type="text" class="Sytle_text" id="CPerson"
			 readonly value="<%=rs("CPerson")%>" size="30"></td>
          </tr>
          <tr>
            <td height="25"><span class="STYLE3">主持人：</span></td>
            <td><input name="ZPerson" type="text" class="Sytle_text" id="ZPerson" readonly
			 value="<%=rs("ZPerson")%>"></td>
            <td class="STYLE3">会议地点：</td>
            <td><input name="address" type="text" class="Sytle_text" id="address" 
			readonly value="<%=rs("address")%>" size="30"></td>
          </tr>
          <tr>
            <td height="27"><div class="STYLE3">主&nbsp;&nbsp;题：
			</div></td>
            <td colspan="3"><input name="subject" type="text" class="Style_subject"
			 id="subject" readonly value="<%=rs("subject")%>"></td>
          </tr>
          <tr>
            <td height="133" valign="top"><div class="STYLE3"><br>
              内&nbsp;&nbsp;容
			：</div></td>
            <td colspan="3"><textarea name="content" cols="53" rows="6" class="Style_content"
			 id="content"><%=rs("content")%></textarea></td>
          </tr>
          <tr>
            <td colspan="4"><div align="center">
                <input name="myclose" type="button" class="Style_button_del" id="myclose"
				 value="关闭窗口" onClick="javascrip:window.close()">
                </div></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
</body>
</html>
