<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<%
set rs=server.CreateObject("adodb.recordset")
sql="select * from DB_Poll order by bookname ASC"
rs.open sql,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>读者信息反馈！</title>
<style type="text/css">
<!--
.style1 {
	font-size: 12pt;
	font-weight: bold;
}
body {
	margin-top: 0px;
}
.STYLE3 {
	font-size: 9pt;
	color: #000000;
}
-->
</style>
</head>

<body>
<form action="Poll_OK.asp" method="post"  name="form1">
  <table width="804" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td height="157" background="images/top.jpg">&nbsp;</td>
    </tr>
    <tr>
      <td height="309" align="center" background="images/middle.jpg"><table width="476" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="302"><img src="images/banner.jpg" width="188" height="37"></td>
          <td width="174">&nbsp;</td>
        </tr>
      </table>
        <table width="476" height="36" border="0" align="center"
		 cellpadding="0" cellspacing="0" bordercolor="#FFFFFF"  bordercolorlight="#FFFFFF" bordercolordark="#0000FF">
        
        <tr>
          <td width="472" height="9" align="right" valign="middle" nowrap></td>
          </tr>
        <tr>
          <td height="27" align="right" valign="middle" nowrap><table width="86%"  border="1" cellpadding="0" cellspacing="0"
			   bordercolordark="#874809">
            <% 
				While NOT rs.EOF
				%>
            <tr>
              <td width="7%"><div align="center">
                <input name="myif" type="radio" value="<%=rs("ID")%>"
					  checked="checked">
                </div></td>
                  <td width="93%" align="left"><span class="STYLE3">&nbsp;<%=(rs("bookname"))%></span></td>
                </tr>
            <% 
				rs.MoveNext()
				Wend
				%>
            </table></td>
          </tr>
        </table>
          <br>
      <input type="image" name="imageField" src="images/btn.jpg"></td></tr>
    <tr>
      <td height="73" background="images/bottom.jpg">&nbsp;</td>
    </tr>
  </table>
  </form>
</body>
</html>
