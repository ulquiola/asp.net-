<%@language="VBScript" %>
<!--#include file="Conn/conn.asp" -->
<%
Set rs = Server.CreateObject("ADODB.Recordset")
sql_rs= "SELECT bookname, poll FROM DB_Poll ORDER BY bookname DESC"
rs.open sql_rs,conn,3
%>
<%
Set rs_sum = Server.CreateObject("ADODB.Recordset")
sql_sum= "SELECT sum(poll) AS sumpoll  FROM DB_Poll"
rs_sum.open sql_sum,conn,3
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>投票成功！</title>
<link href="style.css" rel="stylesheet">
<style>
<!--
.style1 {
	color: #990000;
	font-size: 10pt;
	font-weight: bold;
}
.style4 {color: #990000}
.style5 {color: #336699}
.STYLE3 {	font-size: 9pt;
	color: #000000;
}
-->
</style>
</head>
<body leftmargin="0" topmargin="0">
<div align="center">
  <table width="804" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td height="157" background="images/top.jpg">&nbsp;</td>
    </tr>
    <tr>
      <td height="309" align="center" background="images/middle.jpg">
        <table width="82%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="302"><img src="images/count.jpg" width="183" height="30"></td>
            <td width="174">&nbsp;</td>
          </tr>
        </table>
        <br>
        <table width="80%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td  valign="bottom"><table width="100%" height="22" border="1" align="center"
			   cellpadding="1" cellspacing="1" bordercolor="#874809" >
                <%i=1
				rs.movefirst %>
                <% While NOT rs.EOF %>
                <% imgname="images/image_poll/img_"&i&".gif"%>
                <tr>
                  <td width="25" align="center" valign="middle" nowrap="nowrap"><img src="<%= imgname %>" width="10" height="17" /></td>
                  <td height="20" align="left" valign="middle" nowrap="nowrap">&nbsp;<span class="style5"><%=rs("bookname")%> </span></td>
                  <td width="120" align="center" valign="middle" nowrap="nowrap">共有[<font color="#990000"> <%=rs("poll")%></font>]人投票</td>
                  <td width="84" align="center" valign="middle" nowrap="nowrap">占 <font color="#990000"> <%=round(rs("poll")/rs_sum("sumPoll")*100,2)%></font>%</td>
                </tr>
                <% rs.MoveNext()
				  i=i+1
				Wend %>
              </table>
                <table width="100%"  border="0" cellspacing="-2" cellpadding="-2">
                  <tr>
                    <td height="18"><div align="right"><br>
                      目前共有[<span class="style4"> <%=rs_sum("sumpoll")%></span>]人投票！</div></td>
                  </tr>
            </table></td>
          </tr>
        </table>
      <p><a href="index.asp"><img src="images/reset.jpg" width="117" height="46" border="0"></a></p></td>
    </tr>
    <tr>
      <td height="73" background="images/bottom.jpg">&nbsp;</td>
    </tr>
  </table>
</div>
</body>
</html>