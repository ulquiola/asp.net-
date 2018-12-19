<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<table width="200" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
<%
Set rs=Server.CreateObject("ADODB.Recordset")
sqlstr="select top 8 id,Mtitle,Mname from tab_music order by id desc"
rs.open sqlstr,conn,1,1
while not rs.eof
str=rs("Mtitle")&"("&rs("Mname")&")"
Response.Write("<a href='web_music.asp?s_type=0&keyword="&rs("Mtitle")&"'>"&Left(str,16)&"</a>"&"<BR>")
Set rsc=Nothing
rs.movenext
wend
rs.close
Set rs=Nothing
%>	
	</td>
  </tr>
  <tr>
    <td height="23" align="right"><a href="web_music.asp"><img src="images/more.gif" width="23" height="5" border="0"></a>&nbsp;</td>
  </tr>
</table>
