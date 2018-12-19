<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="Css/style.css">
<style type="text/css">
<!--
.yellow02{color:#E56B04}
a.text{color:#000000}
-->
</style>
<ul>
<%
    Set rs=Server.CreateObject("ADODB.Recordset")
	sqlstr="select top 10 id,Aftitle,Afcontent,Afdate from tb_affiche order by Afdate desc"
	rs.open sqlstr,conn,1,1
	while not rs.eof
  %>
      <li><a href="#" onclick="javascript:window.open('Affiche_show.asp?id=<%=rs("id")%>','new','height=390,width=670,scrollbars=yes')" title="<%=formatdatetime(rs("Afdate"),1)%>发布" class="text"><%=rs("Aftitle")%></a><br /><hr style=" border: 1px #999999 thin;" /></li>
      <%
    rs.movenext
	wend
	rs.close
	Set rs=Nothing
  %>
</ul>