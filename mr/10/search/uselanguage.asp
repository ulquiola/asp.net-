<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp"-->
<!--#include file="shangbiao.asp"-->
<%
set u=server.CreateObject("adodb.recordset")
u_1="select * from tb_language"
u.open u_1,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>明日在线搜索</title>
</head>
<body onLoad="clockon()"background="images/bg.gif">
<table width="780" height="287" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr bgcolor="#FFFFFF">
    <td width="43" valign="top"></td>
    <td width="736" valign="top"><table width="780" height="51" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="20" background="images/sousou2.gif">&nbsp;</td>
          <td width="76" height="47" background="images/tubiao.gif">&nbsp;</td>
          <td width="4" height="47" valign="bottom" background="images/sousou2.gif" bgcolor="#FFFFFF">&nbsp;</td>
          <td width="80" valign="bottom" background="images/changyongshuyu.gif" bgcolor="#FFFFFF">&nbsp;</td>
          <td width="381" align="right" valign="bottom" background="images/sou2.gif" bgcolor="#FFFFFF"><a href="index.asp"style="color:#3661a6"></a></td>
          <td width="73" align="right" valign="bottom" background="images/sou2.gif" bgcolor="#FFFFFF"><div align="center"><a href="index.asp"style="color:#3661a6; font-size: 9pt">返回首页</a></div></td>
          <td width="91" valign="bottom" background="images/sousouhou.gif" bgcolor="#FFFFFF"><a href="index.asp"style="color:#3661a6"></a></td>
        </tr>
        <tr>
          <td height="3" colspan="7" background="images/erjiye1.gif"></td>
        </tr>
        <tr>
          <td height="30" colspan="7">&nbsp;</td>
        </tr>
      </table>
    </p>
    <table width="703" height="51" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
	  <%for i=1 to u.recordcount %>
	  <%if i mod 4=0 then %>
        <td height="25" width="25%" style="padding-left:10pt;">
		<a href="#"onClick="javascript:window.open('introduce8.asp?id=<%=u("ID")%>')"><%=u("name1")%></a>
		</td>
        </tr>
		<%else%>
        <td height="25" width="25%" style="padding-left:10pt;">
		<a href="#"onClick="javascript:window.open('introduce8.asp?id=<%=u("ID")%>')"><%=u("name1")%></a>
		</td>
    <%end if %>
    <%
	  u.movenext
	  next
	  %>
 </table>  
    <br>
    <br>
    <br>
    <table width="757" height="13" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="734" align="right">[共<font color="red"><%=i-1%></font>条记录]</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
