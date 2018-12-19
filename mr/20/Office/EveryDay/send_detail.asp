<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Conn/conn.asp" -->
<%
Set rs_send = Server.CreateObject("ADODB.Recordset")
sql_SP= "SELECT * FROM Tab_Send WHERE ID = "&Request.QueryString("ID")&""
rs_send.open sql_SP,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>发文详细信息页</title>
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
.STYLE4 {
	font-size: 9pt;
	color: #FFFFFF;
}
-->
</style></head>

<body>
<table width="557" height="401" border="0" cellpadding="-2" cellspacing="-2"
 background="../Images/detail0.gif">
  <tr>
    <td width="817" height="349" valign="top"><table width="100%" height="90"
	  border="0" cellpadding="-2" cellspacing="-2">
      <tr>
        <td height="43" valign="bottom">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="STYLE4">发文详细信息页</span></td>
      </tr>
      <tr>
        <td height="47">&nbsp;</td>
      </tr>
    </table>      
      <form name="form1">
        <table width="92%" height="207"  border="0" align="center" cellpadding="-2" cellspacing="-2">
          <tr>
            <td width="20%" height="27"><div align="left" class="style1">收文人：</div></td>
            <td width="30%"><input name="Lperson" type="text" class="Sytle_text" id="Lperson"
			 value="<%=rs_send("LPerson")%>" readonly></td>
            <td width="20%"><div align="right" class="style1">发文日期：</div></td>
            <td width="30%"><input name="textfield" type="text" class="Sytle_text"
			 value="<%=rs_send("sTime")%>" readonly></td>
          </tr>
          <tr>
            <td height="27"><span class="style1">主题：</span></td>
            <td colspan="3"><input name="subject" type="text" class="Style_subject" 
			id="subject" value="<%=(rs_send.Fields.Item("subject").Value)%>" readonly></td>
          </tr>
          <tr>
            <td height="10" valign="top"><div align="center" class="style1">
              <div align="left">内容：</div>
            </div></td>

            <td colspan="3"><textarea name="content" cols="53" rows="6"
			 id="content"><%=rs_send("content")%></textarea></td>
          </tr>
          <tr>
            <td colspan="4"><div align="center">
               <input name="myclose" type="button" class="Style_button" id="myclose"
				 value="关闭" onClick="javascrip:opener.location.reload();window.close()">
                </div></td>
          </tr>
        </table>
      
        
      </form></td>
  </tr>
</table>
</body>
</html>
