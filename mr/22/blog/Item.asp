<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="40"><table width="594" height="40" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td colspan="2"><table width="100%" height="40" border="0" cellpadding="0" cellspacing="1">
            <tr>
              <%
n=1
Set rs=Server.CreateObject("ADODB.Recordset")
sqlstr="select id,Acname from tab_article_class"
rs.open sqlstr,conn,1,1
while not rs.eof
%>
              <td align="center">[&nbsp;<a href="web_blog_list.asp?classid=<%=rs("id")%>&classname=<%=rs("Acname")%>"><%=rs("Acname")%></a>&nbsp;]</td>
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
    <td background="images/xuxian.gif" height="10"></td>
  </tr>
</table>
