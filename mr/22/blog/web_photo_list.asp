<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="css/css.css" />
<title>博客</title>
</head>
<body>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="2"><!--#include file="top.asp"--></td>
  </tr>
  <tr>
    <td width="210" valign="top" bgcolor="#FFFFFF"><!--#include file="left.asp"--></td>
    <td width="568" valign="top" bgcolor="#FFFFFF"><table width="98%" height="400"  border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td valign="top" bgcolor="#FFFFFF">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="22">我的相册--<%=Request.QueryString("classname")%></td>
          </tr>
          <tr>
		  <br />
<%
classid=Request.QueryString("classid")
n=1
Set rs=Server.CreateObject("ADODB.Recordset")
sqlstr="select id,Pname,Ppic from tab_photo where Pclass="&classid&""
rs.open sqlstr,conn,1,1
while not rs.eof
%>
            <td><table width="200" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td align="center"><img src="upfile/<%=rs("Ppic")%>" height="100" width="120" border="0" /></td>
                </tr>
                <tr>
                  <td height="22" align="center"><%=rs("Pname")%></td>
                </tr>
            </table></td>
            <%If n mod 3=0 Then%>
          </tr>
          <tr>
            <%End If%>
            <%
n=n+1
rs.movenext
wend
rs.close
Set rs=Nothing
%>
          </tr>
      </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td colspan="2"><!--#include file="bottom.asp"--></td>
  </tr>
</table>
</body>
</html>
