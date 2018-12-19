<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Conn/conn.asp" -->
<%
if request.QueryString("ID")<>"" then
	session("ID")=request.QueryString("ID")
end if
Set rs_lect = Server.CreateObject("ADODB.Recordset")
sql_lect= "SELECT * FROM Tab_Send WHERE ID = "&session("ID")&""
rs_lect.open sql_lect,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>删除收文信息！</title>
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
<table width="557" height="401" border="0" cellpadding="-2" cellspacing="-2" background="../Images/detail0.gif">
  <tr>
    <td width="817" height="349" valign="top"><table width="100%" height="78"  border="0" cellpadding="-2" cellspacing="-2">
      <tr>
        <td height="43" valign="bottom">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="STYLE4">删除收文信息</span></td>
      </tr>
      <tr>
        <td height="35">&nbsp;</td>
      </tr>
    </table>      
      <form ACTION="send_del_ok.asp" METHOD="POST" name="form1">
        <table width="89%" height="207"  border="0" align="center" cellpadding="-2" cellspacing="-2">
          <tr>
            <td width="20%" height="27"><div align="left" class="style1">收文人：</div></td>
            <td width="30%"><input name="Lperson" type="text" class="Sytle_text" id="Lperson" value="<%=rs_lect("LPerson")%>" readonly></td>
            <td width="20%"><div align="right" class="style1">发文日期：</div></td>
            <td width="30%"><input name="textfield" readonly type="text" class="Sytle_text" value="<%=rs_lect("sTime")%>"></td>
          </tr>
          <tr>
            <td height="27"><span class="style1">主题：</span></td>
            <td colspan="3"><input name="subject" type="text" class="Style_subject" id="subject" value="<%=rs_lect("subject")%>" readonly></td>
          </tr>
          <tr>
            <td height="10" valign="top"><div align="center" class="style1">
              <div align="left">内容：</div>
            </div></td>

            <td colspan="3"><textarea name="content" cols="53" rows="6" id="content"><%=rs_lect("content")%></textarea></td>
          </tr>
          <tr>
            <td colspan="4"><div align="center">
                <input name="Button" type="submit" class="Style_button_del" value="确定删除">
                <input name="myclose" type="button" class="Style_button_del" id="myclose" value="关闭窗口" onClick="javascrip:window.close()">
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
