<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="include/conn.asp"-->
<!--#include file="include/function.asp"-->
<%
'����Time.aspҳ��Ĵ�ֵ
If Request.Form("add")<>"" Then	Session("add")=Request.Form("add")
If Request.Form("minus")<>"" Then Session("minus")=Request.Form("minus")
'��ȡ���ݵĲ���,���ݲ���ֵȷ��SQL��ѯ���
classid=Request.QueryString("classid")
classname=Request.QueryString("classname")
If classname<>"" Then megstr="<font color=#FF0000>"&classname&"</font>"&" ֮"
btype=Request.Form("btype")
keyword=Request.Form("keyword")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>����</title>
<link rel="stylesheet" href="css/css.css" />
</head>
<body>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="2"><!--#include file="top.asp"--></td>
  </tr>
  <tr>
    <td width="230" valign="top" bgcolor="#FFFFFF"><!--#include file="left.asp"--></td>
    <td width="670" valign="top" bgcolor="#FFFFFF"><table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td bgcolor="#FFFFFF"><!--#include file="Item.asp"-->
            <table width="100%" border="0" cellspacing="2" cellpadding="2">
              <tr>
                <td height="20" colspan="2">&nbsp;&nbsp;<%=megstr%>�����б�</td>
              </tr>
              <%
Set rs=Server.CreateObject("ADODB.Recordset")
sqlstr="select top 25 id,Atitle,Adate,Aclass from tab_article where 1=1"
If IsNumeric(classid) and classid<>"" Then sqlstr=sqlstr&" and Aclass="&classid&""
If keyword<>"" Then
Select case btype
case "1"
  sqlstr=sqlstr&" and Atitle like '%"&keyword&"%'"
case "2"
  sqlstr=sqlstr&" and Acontent like '%"&keyword&"%'"
case "3"
  sqlstr=sqlstr&" and Aauthor like '%"&keyword&"%'"
End Select
End If
If Session("times")<>"" Then 
sqlstr=sqlstr&" and Adate like '"&Session("times")&"%'"
Session("times")=""
End If
sqlstr=sqlstr&" order by id desc"
rs.open sqlstr,conn,1,1
If rs.eof Then
%>
              <tr>
                <td height="20" colspan="2" align="center">����ѯ�ļ�¼�����ղأ�</td>
              </tr>
              <%End IF%>
              <%
while not rs.eof
Set rsc=conn.Execute("select Acname from tab_article_class where id="&rs("Aclass")&"")
%>
              <tr>
                <td width="23%" height="20">&nbsp;&nbsp;[<%=formatDateTime(rs("Adate"),2)%>]</td>
                <td width="77%" height="20"><a href="web_blog_view.asp?id=<%=rs("id")%>&amp;classname=<%=rsc("Acname")%>"><%=rs("Atitle")%></a></td>
              </tr>
              <%
Set rsc=Nothing
rs.movenext
wend
rs.close
Set rs=Nothing
%>
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
