<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="shangbiao.asp"-->
<!--#include file="Conn/conn.asp"-->
<%
set cc=server.CreateObject("adodb.recordset")
kk="select * from tb_books where type1='出版社'"
cc.open kk,conn,1,3
set cc1=server.CreateObject("adodb.recordset")
kk1="select * from tb_books where type1='图书网站'"
cc1.open kk1,conn,1,3
set cc2=server.CreateObject("adodb.recordset")
kk2="select * from tb_books where type1='电子图书馆'"
cc2.open kk2,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>搜索引擎</title>
</head>
<body onLoad="clockon()" background="images/bg.gif">
<table width="780" height="375" border="0" align="center" cellpadding="０" cellspacing="０" bgcolor="#FFFFFF">
  <tr>
    <td width="10" height="375" valign="top"></td>
    <td width="780" align="right" valign="top">      
      <table width="780" height="51" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="20" background="images/sousou2.gif">&nbsp;</td>
          <td width="76" height="47" background="images/tubiao.gif">&nbsp;</td>
          <td width="4" height="47" valign="bottom" background="images/sousou2.gif" bgcolor="#FFFFFF">&nbsp;</td>
          <td width="80" valign="bottom" background="images/tushuziyuan.gif" bgcolor="#FFFFFF">&nbsp;</td>
          <td width="381" align="right" valign="bottom" background="images/sou2.gif" bgcolor="#FFFFFF"><a href="index.asp"style="color:#3661a6"></a></td>
          <td width="73" align="right" valign="bottom" background="images/sou2.gif" bgcolor="#FFFFFF"><div align="center"><a href="index.asp"style="color:#3661a6; font-size: 9pt">返回首页</a></div></td>
          <td width="91" valign="bottom" background="images/sousouhou.gif" bgcolor="#FFFFFF"><a href="index.asp"style="color:#3661a6"></a></td>
        </tr>
        <tr>
          <td height="3" colspan="7" background="images/erjiye1.gif"></td>
        </tr>
        <tr>
          <td height="30" colspan="7"></td>
        </tr>
      </table>      
      <table width="699" height="235" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="100%" height="13" valign="middle"  colspan="5"><div align="left"><img src="images/pub.gif" width="522" height="29"></div></td>
      </tr>
        <td valign="top">
		<tr>
		<%for i=1 to cc.recordcount %>	    
          <%if i mod 5=0 then %>
		  <td height="25">		  
		  <a href="<%=cc("bookaddress")%>" target="_blank"><%=cc("name1")%></a>	 </td>
      </tr>
	  <%else%> 
	  <td height="25">
		<a href="<%=cc("bookaddress")%>" target="_blank" style="font-size: 9pt"><%=cc("name1")%> </a></td> 
	    <%end if %>
	  <%
	  cc.movenext
	  next
	  %>
	  <tr>
        <td height="43" valign="middle" colspan="5"><div align="center" style="font-size: 9pt">
          <p>&nbsp;</p>
          <p align="left"><img src="images/booknet.gif" width="522" height="29"></p>
        </div></td>
      </tr>
	  <tr>
	  <%for i=1 to cc1.recordcount%>
	  <%if i mod 5=0 then %>
	  <td height="25">
	  <a href="<%=cc1("bookaddress")%>" target="_blank" style="font-size: 9pt"><%=cc1("name1")%> </a></td>
	 </tr>
	 <%else%>
	 <td>
	 <a href="<%=cc1("bookaddress")%>" target="_blank" style="font-size: 9pt"><%=cc1("name1")%> </a></td>
	   <%end if %>
	 <%
	 cc1.movenext
	 next
	 %>
      <tr>
        <td height="21" colspan="5" valign="middle"><div align="center" style="font-size: 9pt">
          <p align="left"><img src="images/diangzhi.gif" width="522" height="29">　　　　　　　　　</p>
        </div></td>
      </tr>
	  <tr>
	  <%for i=1 to cc2.recordcount%>
	  <%if i mod 5=0 then %>
	  <td height="25">
	  <a href="<%=cc2("bookaddress")%>" target="_blank" style="font-size: 9pt"><%=cc2("name1")%> </a></td>
	 </tr>
	 <%else%>
	 <td>
	 <a href="<%=cc2("bookaddress")%>" target="_blank" style="font-size: 9pt"><%=cc2("name1")%> </a></td>
	   <%end if 
	 cc2.movenext
	 next
	 %>
    </table></td>
  </tr>
</table>
</body>
</html>
