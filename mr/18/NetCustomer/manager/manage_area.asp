<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->
<%
Set rs_province = Server.CreateObject("ADODB.Recordset")
sql_P = "SELECT *  FROM Q_area  WHERE TypeID = 1 or TypeID=4 or TypeID=5"
rs_province.open sql_P,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
 "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>网上客户管理系统--管理员登录！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	background-image: url(../Images/bg.gif);
}
.style2 {color: #a2bcc5}
.style3 {color: #669999}
-->
</style></head>

<body>
<table width="595" height="109" border="0" cellpadding="-2" cellspacing="-2"
 background="../Images/manage.gif">
  <tr>
    <td valign="top">&nbsp;</td>
  </tr>
</table>
<table width="595" height="171" cellpadding="-2" cellspacing="-2">
  <tr>
    <td valign="top"><table width="595" cellpadding="-2" cellspacing="-2" bgcolor="#FFFFFF">
      <tr>
        <td width="33">&nbsp;</td>
        <td width="343" valign="top">地域信息：
            <table width="99%" height="204" border="1" cellpadding="-2"  cellspacing="-2"
			 bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr>
                <td height="202" valign="top">

<script language="JavaScript">
function show_div(menu)
	{if (menu.style.display == "none")
		{menu.style.display = "";}
	else
		{menu.style.display = "none";}
	}
      </script>
<% flag=1
rs_province.movefirst
while not rs_province.eof
	menuName="menu"&flag %>
	<P>　<A Href="#" onClick="show_div(<%=menuName
	 %>)"><img src="../Images/area/jcxx_1.gif" BORDER="0" ALIGN="ABSMIDDLE" width="39" height="16">
	 <%=rs_province("Name")%>
	(<%=rs_province("TypeName")%>)<br>
 </a>
<DIV ID=<%=menuName%> style="display:none">
<%
Set rs_city = Server.CreateObject("ADODB.Recordset")
myID=rs_province("ID")
sql="select * from Q_Area where father="&myID
rs_city.open sql,conn,1,3
flag_city=1
if not rs_city.eof then
rs_city.movefirst
while not rs_city.eof
 if flag_city<>rs_city.recordcount then
 %>
&nbsp;&nbsp;<img src="../Images/area/open_1.gif" width="39" height="16">
              <%else%>
&nbsp;&nbsp;<img src="../Images/area/open_2.gif" width="39" height="16">
              <%end if%>
              <%
  menuName_A=menuName&"_"& flag_city
  %>
              <A Href="#" onClick="show_div(<%=menuName_A%>)"><%=rs_city("name")%>
			  (<%=rs_city("TypeName")%>)</a><br>
              <DIV ID=<%=menuName_A%> STYLE="display:None">
                <%
Set rs_town = Server.CreateObject("ADODB.Recordset")
myID=rs_city("ID")
sql_T="select * from Q_Area where father="&myID
rs_town.open sql_T,conn,1,3
flag_town=1
if not rs_town.eof then
rs_town.movefirst
while not rs_town.eof
if flag_city<>rs_city.recordcount then
 if flag_town<>rs_town.recordcount then
 %>
&nbsp;&nbsp;<img src="../Images/area/open_1_A.gif" height="16">
<%else%>
&nbsp;&nbsp;<img src="../Images/area/open_2_A.gif" height="16">
<%end if
else
 if flag_town<>rs_town.recordcount then %>
&nbsp;&nbsp;<img src="../Images/area/open_1_B.gif" height="16">
<%else%>
&nbsp;&nbsp;<img src="../Images/area/open_2_B.gif" height="16">
<%end if
end if
%>
<a href="#"><font size="2"><%=rs_town("name")%>
(<%=rs_town("TypeName")%>)</font></a><br>
<% flag_town=flag_town+1
rs_town.movenext
wend
end if
 %>
</div>
<% flag_city=flag_city+1
rs_city.movenext
wend
end if
 %>
</DIV>
<% rs_province.movenext
flag=flag+1
wend %>
</td>
              </tr>
          </table></td>
        <td width="200" valign="top">&nbsp;
            <table width="85%" height="210" align="center" cellpadding="-2"  cellspacing="-2">
              <tr>
                <td><div align="center">
               <input name="Button" type="button" class="Style_button_del" value="添加省级名称" 
				onClick="JScript:window.open('add_area_P.asp','','width=398,height=270')">
                </div></td>
              </tr>
              <tr>
                <td><div align="center">
                <input name="Submit2" type="submit" class="Style_button_del" value="添加市/县名称" 
				onClick="JScript:window.open('add_area_C.asp','','width=398,height=270')">
                </div></td>
              </tr>
              <tr>
                <td><div align="center">
               <input name="Submit3" type="submit" class="Style_button_del" value="添加区/镇/乡名称"
				onClick="JScript:window.open('add_area_T.asp','','width=398,height=270')">
                </div></td>
              </tr>
              <tr>
                <td><div align="center">
                    <input name="Submit4" type="submit" class="Style_button_del" value="删除省份名称"
					 onClick="JScript:window.open('del_area_P.asp','','width=398,height=270')">
                </div></td>
              </tr>
              <tr>
                <td><div align="center">
                    <input name="Submit5" type="submit" class="Style_button_del" value="删除市/县名称"
					 onClick="JScript:window.open('del_area_C.asp','','width=398,height=270')">
                </div></td>
              </tr>
              <tr>
                <td><div align="center">
                    <input name="Submit6" type="submit" class="Style_button_del" value="删除区/镇/乡名称"
					 onClick="JScript:window.open('del_area_T.asp','','width=398,height=270,status=yes')">
                </div></td>
              </tr>
              <tr>
                <td><span class="style3">注：添加自治区、直辖市请点击[添加省级名称]按钮。</span></td>
              </tr>
            </table>            </td>
        <td width="17">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
<table width="595" height="60" cellpadding="-2" cellspacing="-2" background="../Images/manage_03.gif">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
