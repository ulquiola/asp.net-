<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/conn.asp" -->
<!--#include file="../../NetCustomer/F_Display_z.asp" -->
<%
'��ѯҵ��Ա����
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM Q_Sell WHERE name = '"&session("username")&"'"
rs.open sql,conn,1,3
'��ѯҵ��Ա��Ϣ
Set rs_user = Server.CreateObject("ADODB.Recordset")
sql_user = "SELECT ywynumber,name,workqy FROM DB_WorkerInfo WHERE name='"&session("username")&"'"
rs_user.open sql_user,conn,1,3
'�������۽��
Set rs_money = Server.CreateObject("ADODB.Recordset")
sql_M = "SELECT sum(zmoney) as sMoney FROM Q_sell where name='"&session("username")&"'"
rs_money.open sql_M,conn,1,3
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
	background-image: url(../Images/bg_top.gif);
}
.style1 {color: #FFFFFF}
.style2 {color: #a2bcc5}
.style3 {color: #990000}
-->
</style></head>
<script language="javascript">
function mycheck(){
if (form1.sDate.value=="")
{alert("�������ѯ��ʼ���ڣ�");form1.sDate.focus();return;}
if (form1.eDate.value=="")
{alert("�������ѯ�������ڣ�");form1.eDate.focus();return;}
form1.submit();
}
</script>
<body><form method="post" action="quest_bottom.asp" target="info" name="form1">
<table width="595" height="100" border="0" cellpadding="-2" cellspacing="-2" background="../Images/manage_01_q.gif">
  <tr>
    <td valign="top">&nbsp;</td>
  </tr>
</table>
<table width="595" height="100" cellpadding="-2" cellspacing="-2" bgcolor="#FFFFFF">
  <tr>
    <td valign="top"><table width="97%" height="24" cellpadding="-2"  cellspacing="-2">
      <tr>
        <td width="4%"><div align="right"><img src="../Images/add.gif" width="18" height="18"></div></td>
        <td width="71%"><div align="left">&nbsp;ҵ��Ա��<%=session("UserName")%>������������
                <%qyNO=rs_user("workqy")
				  call display(qyNO)%>
        </div></td>
        <td width="25%"><div align="right">�����۶�Ϊ��<span class="style3">
		<%if Isnull(rs_money("sMoney")) then 
		sMoney=0 
		else 
		sMoney=rs_money("sMoney")
		end if%>
		<%=sMoney%>(Ԫ)</span></div></td>
      </tr>
    </table>
	
      <table width="595" height="27" cellpadding="-2" cellspacing="-2">
        <tr>
          <td width="43">
            <div align="right">
              <input name="F" type="checkbox" id="F" value="1" checked>
            </div></td>
          <td width="114"><div align="center">��ѡ���ѯ������</div></td>
          <td width="83"><select name="field" id="field">
            <option value="khname" selected>�ͻ�����</option>
            <option value="spname">��Ʒ����</option>
          </select></td>
          <td width="61"><select name="TJ" id="TJ">
            <option value="=" selected>����</option>
            <option value="like">Like</option>
          </select></td>
          <td width="147"><input name="INvalue" type="text" class="Sytle_auto" id="INvalue"></td>
          <td width="145"><input name="Button" type="button" class="Style_button" value="��ѯ" onClick="mycheck()">
		  <input name="Button" type="button" class="Style_button" value="����" onClick="JScript:window.parent.location.href='default.asp'"></td>
        </tr>
      </table>
      <table width="595" height="27" cellpadding="-2" cellspacing="-2">
        <tr>
          <td width="42"><div align="right">
            <input name="D" type="checkbox" id="D" value="1">
          </div></td>
          <td width="115"><div align="right">��������ʼ���ڣ�</div></td>
          <td width="119"><input name="sDate" type="text" class="Sytle_text" id="sDate" value="2004-5-1"></td>
          <td width="65">�������ڣ�</td>
          <td width="128"><input name="eDate" type="text" class="Sytle_text" id="eDate" value="<%=date()%>"></td>
          <td width="124">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
</form></body>
</html>
