<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>博客</title>
<link rel="stylesheet" href="css/css.css" />
</head>
<body>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="2"><!--#include file="top.asp"--></td>
  </tr>
  <tr>
    <td width="227" valign="top" bgcolor="#FFFFFF"><!--#include file="left.asp"--></td>
    <td width="673" valign="top" bgcolor="#FFFFFF"><table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td bgcolor="#FFFFFF">
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="22">-&gt; 我的相册</td>
              </tr>
              <tr>
                <%
n=1
Set rs=Server.CreateObject("ADODB.Recordset")
sqlstr="select id,Pcname from tab_photo_class"
rs.open sqlstr,conn,1,1
while not rs.eof
Set rsc=conn.Execute("select top 1 Ppic from tab_photo where Pclass="&rs("id")&"")
%>
                <td><table width="200" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td align="center"><a href="web_photo_list.asp?classid=<%=rs("id")%>&amp;classname=<%=rs("Pcname")%>">
                        <%If Not rsc.eof Then%>
                        <img src="upfile/<%=rsc("Ppic")%>" height="100" width="120" border="0" />
                        <%Else%>
                        <img src="upfile/instead.jpg" height="100" width="120" border="0" />
                        <%End If%>
                      </a></td>
                    </tr>
                    <tr>
                      <td height="22" align="center"><a href="web_photo_list.asp?classid=<%=rs("id")%>&amp;classname=<%=rs("Pcname")%>"><%=rs("Pcname")%></a></td>
                    </tr>
                </table></td>
                <%If n mod 3=0 Then%>
              </tr>
              <tr>
                <%End If%>
                <%
Set rsc=Nothing
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
