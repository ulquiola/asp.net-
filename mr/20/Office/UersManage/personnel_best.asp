<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Conn/conn.asp" -->
<%
Set rs = Server.CreateObject("ADODB.Recordset")
sql= "SELECT top 3  UserName,name, AccessTime  FROM Tab_User  ORDER BY AccessTime DESC"
rs.open sql,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>����ְԱ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
.style1 {color: #C60001}
.style2 {color: #669999}
.style3 {color: #FF0000}
.STYLE6 {color: #669999; font-size: 9pt; }
.STYLE7 {
	font-size: 9pt;
	color: #000000;
}
-->
</style></head>

<body>
<table width="421" height="260" border="0" cellpadding="-2" cellspacing="-2">
  <tr>
    <td width="421" valign="top"><table width="100%" height="88"  border="0"
	 cellpadding="-2" cellspacing="-2">
      <tr>
        <td height="88">&nbsp;</td>
      </tr>
    </table>      
      
        <table width="83%"  border="1" align="center" cellpadding="-2" cellspacing="-2"
		 bordercolor="#FFCCCC" bordercolordark="#FFFFFF">
          <tr>
            <td width="21%" height="23"><div align="center" class="STYLE7">����</div></td>
            <td width="30%" height="23"><div align="center" class="STYLE7">�û���</div></td>
            <td width="28%"><div align="center" class="STYLE7">Ա������</div></td>
            <td width="21%" height="23"><div align="center" class="STYLE7" >��վ����</div></td>
          </tr>
		 <% 
		 Mycount=1
		 While NOT rs.EOF 
		 %>
          <tr>
            <td height="23" align="center"><span class="STYLE6">��[</span> <span class="STYLE6">
			<%=Mycount%></span> <span class="STYLE6">]��</span></td>
            <td height="23"><div align="center" class="style2"><%=(rs("UserName"))%></div></td>
            <td height="23"><div align="center" class="style2"><%=(rs("Name"))%></div></td>
            <td height="23" align="center"><span class="style2"><%=(rs("AccessTime"))%></span></td>
          </tr>
          <% 
		  rs.MoveNext()
		    Mycount=Mycount+1
	     Wend 
		 %>
      </table>
        <table width="83%" height="40"  border="0" align="center" cellpadding="-2" cellspacing="-2">
          <tr>
            <td height="10"></td>
          </tr>
		  <tr>
            <td valign="bottom"> <div align="center"> 
              <form name="form1" method="post">
                <input name="myclose" type="button" class="Style_button_del" id="myclose"
				 value="�رմ���" onClick="javascrip:window.close()">
              </form></div></td>
          </tr>
    </table></td>
  </tr>
</table>
</body>
</html>>
