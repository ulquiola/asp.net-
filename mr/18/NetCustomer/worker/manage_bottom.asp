<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/conn.asp" -->
<!--#include file="../../NetCustomer/F_Display_z.asp" -->
<%
'��ѯҵ��Ա��Ϣ
Set rs_user = Server.CreateObject("ADODB.Recordset")
sql_user="SELECT ywynumber, name, workqy FROM DB_WorkerInfo WHERE name ='"&_
Session("UserName")&"'"
rs_user.open sql_user,conn,1,3
'����Ʒ���ƻ���������Ϣ
Set rs_sp = Server.CreateObject("ADODB.Recordset")
sql_sp="SELECT spname,pack,sum(sellsum) as H_sum,sum(zmoney) as H_money FROM Q_sell"&_
" WHERE name ='"&Session("UserName")&"' group by spname,pack"
rs_sp.open sql_sp,conn,1,3
'���ͻ����ƻ���������Ϣ
Set rs_kh = Server.CreateObject("ADODB.Recordset")
sql_kh="SELECT khname,sum(sellsum) as H_sellsum,sum(zmoney) as H_money FROM Q_sell"&_
" WHERE name ='"&Session("UserName")&"' group by khname"
rs_kh.open sql_kh,conn,1,3
'�����ڻ���������Ϣ
Set rs_date = Server.CreateObject("ADODB.Recordset")
if request.Form("sDate")<>"" and request.Form("eDate")<>"" then
sql_date="SELECT sum(sellsum) as H_sellsum,sum(zmoney) as H_money  FROM Q_sell"&_
" WHERE D_date between #"&request.Form("sDate")&"# and #"&request.Form("eDate")&_
"# and name ='"&Session("UserName")&"'"
else
sql_date="SELECT sum(sellsum) as H_sellsum,sum(zmoney) as H_money  FROM Q_sell"&_
" WHERE name ='"&Session("UserName")&"'"
end if
rs_date.open sql_date,conn,1,3
Set rs = Server.CreateObject("ADODB.Recordset")
sql_sell = "SELECT *  FROM Q_sell"
rs.open sql_sell,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>������Ϣ��ѯ��</title>
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
.style3 {color: #990000}
-->
</style></head>

<body>
<table width="595" height="142" border="0" cellpadding="-2" cellspacing="-2" bgcolor="#FFFFFF">
  <tr>
    <td valign="top"><table width="595" cellpadding="-2" cellspacing="-2">
      <tr>
        <td width="10" height="40">&nbsp;</td>
        <td width="577" valign="top">          <div align="center">
<%if not rs.eof then%>
<% if request.Form("Butt_quest")<>"����" then
flag=0
else
flag=1
end if
%>
<!-----------����Ʒ���ƻ���--------------->	
<%if request.Form("TJ")="2" then%>	
            <table width="96%" border="1" cellpadding="-2"  cellspacing="-2" 
			bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr bgcolor="#809EA4">
                <td width="32%"><div align="center" class="style1">��Ʒ����</div></td>
                <td width="20%"><div align="center" class="style1">��װ</div></td>
                <td width="20%"><div align="center" class="style1">��������</div></td>
                <td width="28%"><div align="center" class="style1">���(Ԫ)</div></td>
                </tr>
<% while not rs_sp.eof %>				
				<tr>
                  <td><div align="center"><%=rs_sp("spname")%></div></td>
                  <td><div align="center"><%=rs_sp("pack")%></div></td>
                  <td><%=rs_sp("H_sum")%></td>
                  <td><div align="center"><%=rs_sp("H_money")%></div></td>
                </tr>
<%rs_sp.movenext
wend%>
            </table>
<%end if%>	
<!-----------���ͻ����ƻ���--------------->	
<%if request.Form("TJ")="1" or flag=0 then%>
            <table width="96%" border="1" cellpadding="-2"  cellspacing="-2"
			 bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr bgcolor="#809EA4">
                <td width="43%"><div align="center" class="style1">�ͻ�����</div></td>
                <td width="22%"><div align="center" class="style1">��������</div></td>
                <td width="35%"><div align="center" class="style1">���(Ԫ)</div></td>
                </tr>
<% while not rs_kh.eof %>				
				<tr>
                  <td><div align="center"><%=rs_kh("khname")%></div></td>
                  <td><%=rs_kh("H_sellsum")%></td>
                  <td><div align="center"><%=rs_kh("H_money")%></div></td>
                </tr>
<%rs_kh.movenext
wend%>
            </table>
<%end if%>	
<!-----------�����ڻ���--------------->	
<%if request.Form("TJ")="3" then%>
	<% if request.Form("sDate")="" then%>
		<script language="javascript">
			alert("�������ѯ��ʼ���ڣ�");
		</script>
	<%end if
	if request.Form("eDate")="" then%>
		<script language="javascript">
			alert("�������ѯ�������ڣ�");
		</script>
	<%end if%>
	<%if not request.Form("sDate")="" and not request.Form("eDate")="" then %>
			<BR>
            <table width="96%" border="1" cellpadding="-2"  cellspacing="-2"
			 bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr bgcolor="#809EA4">
                <td width="12%" height="40"><div align="center" class="style1"></div>                  
                <div align="center" class="style1">������������<%=rs_date("H_sellsum")%>(��)</div>                
                </td>
                <td width="15%"><div align="center" class="style1">�ܽ�
				<%=rs_date("H_money")%>(Ԫ)</div></td>
                </tr>
            </table>
	<%end if
end if%>	
			  <% end if 
			  if rs.eof and rs.bof then
			  %>
<table width="426" height="45" cellpadding="-2" cellspacing="-2">
              <tr>
                <td><div align="center">�����ۼ�¼��</div></td>
              </tr>
            </table>
			<% end if %>
          </div></td>
        <td width="10">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
<table width="595" height="60" cellpadding="-2" cellspacing="-2"
 background="../Images/manage_03.gif">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
