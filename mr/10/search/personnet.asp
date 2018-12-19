<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="shangbiao.asp"-->
<!--#include file="Conn/conn.asp"-->
<%
set rs=server.createobject("adodb.recordset")
rs_1="select * from tb_personnetaddress"
rs.open rs_1,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>明日在线搜索</title>
</head>
<body onLoad="clockon()" background="images/bg.gif">
<table width="780" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="780" height="218" bgcolor="#FFFFFF"><table width="780" height="278" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="730">
     </td>
                </tr>
                                <td width="730" valign="top"><table width="780" height="51" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="20" background="images/sousou2.gif">&nbsp;</td>
          <td width="76" height="47" background="images/tubiao.gif">&nbsp;</td>
          <td width="4" height="47" valign="bottom" background="images/sousou2.gif" bgcolor="#FFFFFF">&nbsp;</td>
          <td width="80" valign="bottom" background="images/rencaiwangzhan.gif" bgcolor="#FFFFFF">&nbsp;</td>
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
                    <table width="702" height="25" border="0" align="center" cellpadding="０" cellspacing="０" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#ccccff">
	          <tr>
	    <%for i=1 to rs.recordcount %>
	    <%if i mod 5=0 then %>
          <td height="24">
		  <a href="<%=rs("netaddress")%>" target="_blank"><%=rs("name")%></a>
		  </td>
          </tr>
		  <%else%>
          <td height="24">
		  <a href="<%=rs("netaddress")%>" target="_blank"><%=rs("name")%></a>
		  </td>
	        <%end if %>
	  <%rs.movenext
	next 
	%>
                  </table>
                    <p>&nbsp;</p>
                    <table width="757" height="13" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="734" align="right">[共<font color="red"><%=i-1%></font>条记录]</td>
                      </tr>
                    </table>
                                </table>
            
   </td>
  </tr>
</table>
<table width="681" height="69" border="0" align="center" cellpadding="０" cellspacing="０">
</table>
</body>
</html>
