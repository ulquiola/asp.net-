<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<table width="200" border="0" cellspacing="0" cellpadding="0">
<%
Set rs=Server.CreateObject("ADODB.Recordset")
sqlstr="select top 8 id,Atitle,Aclass from tab_article order by id desc"
rs.open sqlstr,conn,1,1
while not rs.eof
Set rsc=conn.Execute("select Acname from tab_article_class where id="&rs("Aclass")&"")
%>
  <tr>
    <td background="images/line.jpg">&nbsp;<img src="images/point.gif" width="13" height="11" />&nbsp;<a href="web_blog_view.asp?id=<%=rs("id")%>&classname=<%=rsc("Acname")%>" title="<%=rs("Atitle")%>"><%=Left(rs("Atitle"),16)%></a></td>
  </tr>
<%
Set rsc=Nothing
rs.movenext
wend
rs.close
Set rs=Nothing
%>
  <tr>
    <td height="23" align="right" background="images/line.jpg"><a href="web_blog_list.asp"><img src="images/more.gif" width="23" height="5" border="0"></a>&nbsp;</td>
  </tr>
</table>