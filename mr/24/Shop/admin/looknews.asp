<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../include/conn.asp" -->
<!--#include file="chk_admin.asp" -->
<!--#include file="../include/include.asp" -->
<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#799AE1">
 <tr> 
    <td align="center"><font color="#FFFFFF"> �鿴վ������</font></td>
  </tr>
  <tr> 
    <td valign="top" bgcolor="#FFFFFF"><br> 
      <table width="90%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="799AE1">
<%
sql="select * from [news] where id="&request("id")&";"
set rs=Server.CreateObject("ADODB.Recordset")
rs.open sql,conn,3,3
%>
        <tr height="20" bgcolor="#FFFFFF" align="center"> 
          <td width="25%">�����ˣ�</td>
          <td width="75%"><div align="left"><%=rs("user")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>ʱ&nbsp;&nbsp;�䣺</td>
          <td><div align="left"><%=rs("shijian")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td>��������</td>
          <td><div align="left"><%=rs("biaoti")%></div></td>
        </tr>
        <tr height="20" bgcolor="#FFFFFF" align="center">
          <td height="141">��&nbsp;&nbsp;�ݣ�</td>
          <td><div align="left"><%=HTMLEncode(rs("neirong"))%></div></td>
        </tr>
<%
rs.close
set rs=nothing
%>
      </table>
      <br></td>
  </tr>
</table>