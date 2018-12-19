<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<%
If request.QueryString("id")<>"" then id=request.QueryString("id")
set rs=server.CreateObject("adodb.recordset")
sqlstr="select Mtitle,Mname,Mword from tab_music where id="&id&""
rs.open sqlstr,conn,1,1
Mtitle=rs("Mtitle")
Mword=rs("Mword")
rs.close
Set rs=Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css/css.css" />
<title>mp3下载歌词详细信息显示页</title>
<style type="text/css">
<!--
.style1 {color: #CDEAF6}
.style2 {	font-size: 9pt;
	font-weight: bold;
}
.style3 {	font-size: 9pt;
	color: #000000;
	font-weight: bold;
}
.style5 {font-size: 9pt}
.STYLE6 {color: #FFFFFF}
body {
	background-color: #805729;
}
-->
</style>
</head>

<body>
<p>&nbsp;</p>
<p>&nbsp;</p>
<table width="650" height="324" border="1" align="center" cellpadding="1" cellspacing="0" bordercolor="#805729">
  <tr>
    <td height="322" valign="baseline" bordercolor="#FFFFFF" bgcolor="#FFFFFF"><table width="647" border="0" cellpadding="0" cellspacing="0" bgcolor="#805729">
        <tr>
          <td height="28"><div align="center" class="style2">
            <div align="left"> &nbsp;&nbsp;<span class="STYLE6">歌曲名：</span><font color=#ffAA66><%=Mtitle%></font></div>
          </div></td>
        </tr>
      </table>
        <table width="647" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td><br>
              &nbsp;<%=replace(Mword,chr(13),"<br>")%></td>
          </tr>
        </table><br>                
        <table width="647" height="23" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td bgcolor="#F0F0F0">&nbsp;</td>
          </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
