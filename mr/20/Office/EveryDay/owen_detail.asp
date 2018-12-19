<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Conn/conn.asp" -->
<%
Set rs_lect = Server.CreateObject("ADODB.Recordset")
sql_lect= "SELECT * FROM Tab_Send WHERE ID = "&Request.QueryString("ID")&""
rs_lect.open sql_lect,conn,1,3
UP = "UPDATE  Tab_Send  SET flag ='是'  WHERE ID= '"&Request.QueryString("ID")&"' "
conn.Execute(UP)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>收文详细信息面页</title>
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
.STYLE5 {
	color: #000000;
	font-size: 9pt;
}
-->
</style></head>

<body>
<table width="557" height="401" border="0" cellpadding="-2" cellspacing="-2"
 background="../Images/detail0.gif">
  <tr>
    <td width="817" height="349" valign="top"><table width="100%" height="97"
	 border="0" cellpadding="-2" cellspacing="-2">
      <tr>
        <td height="44" valign="bottom">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="STYLE4">收文详细信息面页</span></td>
      </tr>
      <tr>
        <td height="53">&nbsp;</td>
      </tr>
    </table>      
      <form name="form1">
        <table width="94%" height="207"  border="0" align="center" cellpadding="-2" cellspacing="-2">
          <tr>
            <td width="20%" height="27"><div align="left" class="STYLE5">发文人：</div></td>
            <td width="30%"><input name="Lperson" type="text" class="Sytle_text"
			 id="Lperson" value="<%=rs_lect("SPerson")%>" readonly></td>
            <td width="20%"><div align="right" class="STYLE5">发文日期：</div></td>
            <td width="30%"><input name="textfield" readonly type="text"
			 class="Sytle_text" value="<%=rs_lect("sTime")%>"></td>
          </tr>
          <tr>
            <td height="27"><span class="STYLE5">主题：</span></td>
            <td colspan="3"><input name="subject" type="text" class="Style_subject"
			 id="subject" value="<%=rs_lect("subject")%>" readonly></td>
          </tr>
          <tr>
            <td height="10" valign="top"><div align="center" class="style1">
              <div align="left" class="STYLE5">内容：</div>
            </div></td>

            <td colspan="3"><textarea name="content" cols="53" rows="6"
			 id="content"><%=rs_lect("content")%></textarea></td>
          </tr>
          <tr>
            <td colspan="4"><div align="center">
                <input name="myclose" type="button" class="Style_button" id="myclose"
				 value="关闭" onClick="javascrip:window.close()">
                </div></td>
          </tr>
        </table>
      
        
      </form></td>
  </tr>
</table>
</body>
</html>
<%
rs_lect.Close()
Set rs_lect = Nothing
%>
