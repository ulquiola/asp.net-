<%Sub Html(title)%>
<html>
<head>
<title><%=title%></title>
<meta http-equiv=content-type content="text/html; charset=gb2312">
<LINK href="images/Style.css" type=text/css rel=stylesheet>
</head>
<body bottommargin=0 leftmargin=0 topmargin=0 rightmargin=0>
<%End Sub%>
<%
Sub Top(Title,Url)
Call Html(Title)
%>
<table class="toptable" height=27 cellspacing=0 cellpadding=0 width=778 align=center border=0>
  <tbody>
  <tr>
    <td height=107 align=middle><img src="images/top.jpg" width="777" height="127"></td>
  </tr></tbody></table>
	<table class="table bgcolor_1" cellspacing=1 cellpadding=1 width=778 align=center>
  <tbody>
  <tr>
    <td width="73%" rowspan=3 align=left style="height: 20px">&nbsp;当前位置：<a href="index.asp">首页</a> >> <a 
            href="<%=url%>"><font 
          color=red><%=title%></font></a></td>
    <td width="2%" rowspan=3 align=right style="height: 20px">&nbsp;</td>
    <td style="height: 20px" align=left rowspan=3 width="25%"><div id=jnkc></div>
            <SCRIPT>setInterval("jnkc.innerHTML=new Date().toLocaleString()+' 星期'+'日一二三四五六'.charAt (new Date().getDay());",1000);</SCRIPT></TD>
  </tr>
  <tr></tr></tbody></table>
<%End Sub%>
<%Sub Bottom()
Call CloseDatabase()
%>
<table width=778 align=center cellspacing=1 cellpadding=1 class="bottomtable bgcolor_1">
  <tbody>
  <tr>
    <td align=middle colspan=3 rowspan=3 class="copyright_td"><br>
	  页面执行时间：
	            <%
	            Dim Endtime
				Endtime = Timer()
				response.write"0"&FormatNumber((Endtime-Startime),5)&" 秒"
				%>
	  </td></tr>
  <tr></tr>
  <tr></tr></tbody></table>
</body></html>
<%End Sub%>