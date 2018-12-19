<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="shangbiao.asp"-->
<!--#include file="Conn/conn.asp"-->
<%
set rs=server.CreateObject("adodb.recordset")
rs1="select * from tb_introduce"
rs.open rs1,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>明日在线搜索</title>
</head>
<body onLoad="clockon()" background="images/bg.gif">
<table width="779" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="165" valign="top" bgcolor="#FFFFFF"><table width="780" height="51" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="76" background="images/sousou2.gif">&nbsp;</td>
          <td width="76" height="47" background="images/tubiao.gif">&nbsp;</td>
          <td width="83" height="47" valign="bottom" background="images/guanyuwomen.gif" bgcolor="#FFFFFF">&nbsp;</td>
          <td width="381" align="right" valign="bottom" background="images/sou2.gif" bgcolor="#FFFFFF"><a href="index.asp"style="color:#3661a6"></a></td>
          <td width="73" align="right" valign="bottom" background="images/sou2.gif" bgcolor="#FFFFFF"><div align="center"><a href="index.asp"style="color:#3661a6; font-size: 9pt">返回首页</a></div></td>
          <td width="91" valign="bottom" background="images/sousouhou.gif" bgcolor="#FFFFFF"><a href="index.asp"style="color:#3661a6"></a></td>
        </tr>
        <tr>
          <td height="3" colspan="6" background="images/erjiye1.gif"></td>
        </tr>
        <tr>
          <td height="30" colspan="6">&nbsp;</td>
        </tr>
      </table>      
      <table width="630" height="143" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td width="778" valign="top"><span style="font-size: 9pt"><%=rs("content")%> </span></td>
        </tr>
      </table>
      <p>&nbsp;</p>      </td>
  </tr>
  <tr>
    <td height="102" valign="top" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
</table>
<!--#include file="copyright.asp"-->
</body>
</html>
