<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->
<%
'��ѯȫ����Ʒ��Ϣ
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM DB_SPInfo"
rs.open sql,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>��Ʒ��Ϣ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	background-image: url(../Images/bg.gif);
}
.style1 {color: #FFFFFF}
.style2 {color: #a2bcc5}
-->
</style></head>

<body>
<table width="595" height="242" border="0" cellpadding="-2" cellspacing="-2" background="../Images/manage.gif">
  <tr>
    <td valign="top"><table width="595" height="400" cellpadding="-2" cellspacing="-2">
      <tr>
        <td width="10" height="94">&nbsp;</td>
        <td width="577">&nbsp;</td>
        <td width="10"><span class="style2"></span></td>
      </tr>
      <tr>
        <td height="258">&nbsp;</td>
        <td valign="top">
  
          <div align="center">
            <table width="97%" height="31" cellpadding="-2"  cellspacing="-2">
              <tr>
                <td width="3%"><div align="right"><img src="../Images/add.gif" width="18" height="18"></div></td>
                <td width="51%"><div align="left"><a href="#" onClick="JScript:window.open('add_spinfo.asp','','width=595,height=400')">&nbsp;�������Ʒ</a></div></td>
              <td width="40%"><div align="right"><img src="../Images/back.GIF" width="14" height="14"></div></td>
                <td width="6%"><a href="default.asp">����</a></td>
			  </tr>
            </table>
              <table width="100%" height="47" border="1" cellpadding="-2"  cellspacing="-2" 
			  bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr bgcolor="#809EA4">
                <td width="14%"><div align="center" class="style1">��Ʒ����</div></td>
                <td width="10%"><div align="center" class="style1">��װ</div></td>
                <td width="28%"><div align="center" class="style1">��������</div></td>
                <td width="11%"><div align="center" class="style1">����</div></td>
                <td width="14%"><div align="center" class="style1">����</div>
                </td>
                <td width="12%"><div align="center" class="style1">��ζ����</div></td>
                <td width="5%"><div align="center" class="style1">�޸�</div></td>
                <td width="6%"><div align="center" class="style1">ɾ��</div></td>
              </tr>
            <%'��ҳ��ʾ��Ʒ��Ϣ
			rs.pagesize=7
			page=CLng(Request("page"))
			if page<1 then page=1
			rs.absolutepage=page
			for i=1 to rs.pagesize %>
              <tr>
                  <td><span class="style1">&nbsp;</span><A HREF="#"
				  onClick="JScript:window.open('detail_spinfo.asp?spnumber=<%= rs("spnumber")
				   %>','','width=589,height=400')"><%=rs("spname")%></A></td>
                  <td><div align="center"><%=rs("pack")%></div></td>
                  <td><%=rs("sccs")%></td>
                  <td><%=rs("type")%></td>
                  <td><%=rs("price")%>(Ԫ)</td>
                  <td><%=rs("kwtype")%></td>
                  <td><div align="center" class="style1"><A HREF="#"
				   onClick="JScript:window.open('modify_spinfo.asp?spnumber=<%= rs("spnumber")
				    %>','','width=589,height=400')"> <!--ʵ�ִ��ݲ���������-->
					<img src="../Images/modify.gif" width="20" height="18" border="0"></A></div></td>
                  <td><div align="center" class="style1"><A HREF="#"
				   onClick="JScript:window.open('del_spinfo.asp?spnumber=<%= rs("spnumber")
				    %>','','width=589,height=400')"><img src="../Images/del.gif" width="22"
					 height="22" border="0"></A></div></td>
			  </tr>
            <% rs.movenext
			if rs.eof then exit for 
			next %>
              </table>
            <table width="97%" height="31" cellpadding="-2"  cellspacing="-2">
              <tr>
                <td><div align="right">
            <% if page<>1 then %><a href=<%=path%>?page=1>��һҳ</a>
			<a href=<%=path%>?page=<%=(page-1)%>>��һҳ</a> 
			<%end if 
			if page<>rs.pagecount then %>
   			<a href=<%=path%>?page=<%=(page+1)%>>��һҳ</a> 
			<a href=<%=path%>?page=<%=rs.pagecount%>>���һҳ</a> 
			<%end if %>
			</div></td>
              </tr>
            </table>
          </div></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>

