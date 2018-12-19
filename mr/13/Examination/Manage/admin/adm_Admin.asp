<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../../Conn/conn.asp"-->
<%
getcondition = replace(trim(request("condition")),"'","''")
getkey = replace(trim(request("key")),"'","''")
if(getcondition = "" or getkey = "")then
rssql = "select * from tb_Administrator order by JoinTime Desc"
else
rssql = "select * from tb_Administrator where "&getcondition&" like '%"&getkey&"%' order by JoinTime desc"
end if
set rs = server.createobject("adodb.recordset") 
rs.open rssql,conn,1,3
rs.pagesize = 5
'实现分页
if rs.eof then
rs_total = 0
else
rs_total = rs.recordcount
end if

dim pageno
pageno = replace(trim(request("pageno")),"'","")
if(pageno = "")then
pageno = 1
End if
'if(not rs.eof)then
'rs.absolutepage = pageno
'end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>管理员信息</title>
<link href="../../Css/style.css" rel="stylesheet">
<script language="javascript" src="js/js.js"></script>
<style type="text/css">
<!--
.style1 {color: #FFFFFF}
-->
</style>
</head>
<body  topmargin="0" leftmargin="0">
<table width="100%"  border="0" cellspacing="0" cellpadding="0" align=center >
	<form name="queryadmin" method="post" action="adm_Admin.asp">
	<tr>
		<td height="27" colspan=6>&nbsp;</td>
	</tr>
  	<tr>
   	 	<td height="25px" align="center">查询条件：
			<select name="condition" class="wenbenkuang" style="width:150;border:1px solid;">
			<option value="Name" <%If(InStr(getcondition,"Name") > 0)Then%>selected<%End If%>>管理员名称</option>
			<option value="JoinTime" <%If(InStr(getcondition,"JoinTime") > 0)Then%>selected<%End If%>>加入时间</option>
			</select>
		关键字：
			<input name="key" type="text" class="txt_grey" value="<%=Server.HTMLEncode(Replace(getkey,"''","'"))%>">
		<input name="submit" type="submit" class="btn_grey" style="border:1px solid;" value="查询"></td>
    </tr>
	</form>
</table>

<form name="deladmin" method="post" action="#">
<table width="98%"  border="1" cellspacing="0" cellpadding="0" align="center"
 bordercolor="#FFFFFF" bordercolordark="#1985DB" bordercolorlight="#FFFFFF"> 
	<tr>
  		<td width="5%" height="27"  align="center" nowrap style="color:black;">选 项</td>
    	<td width="28%" height="27"  align="center" nowrap style="color:black;">管 理 员 名 称</td>
		<td width="25%" height="27"  align="center" nowrap style="color:black;">管 理 员 密 码</td>
		<td width="32%" height="27"  align="center" nowrap style="color:black;">加 入 时 间</td>
		<td width="10%" height="27" align="center" nowrap style="color:black;">修 改</td>
	</tr>
<%
  if(rs.eof)then
%>
  <tr><td height="27" colspan=7 align=center nowrap><font color="red">没有符合条件的记录！</font></td></tr>	
<%
  else
  repeat_rows = 0
  while(repeat_rows < rs.pagesize) and (not rs.eof)
%>
	<tr>
		<td height="27"  align=center nowrap>
    	<input type="checkbox" name="ChkBox" value="<%=server.htmlencode(rs("ID"))%>"  style="border:0;"></td>
		<td height="27"  align=center nowrap style="color:black;"><%=server.htmlencode(rs("Name"))%></td>
		<td height="27"  align=center nowrap style="color:black;"><%=server.htmlencode(rs("PWD"))%></td>
		<td height="27"  align=center nowrap style="color:black;"><%=Mid(rs("JoinTime"),1,InStr(rs("JoinTime")," "))%></td>
		<td height="27"  align=center nowrap style="color:black;">[<a href="adm_UpdateAdmin.asp?id=<%=rs("ID")%>&condition=<%=getcondition%>&key=<%=getkey%>&pageno=<%=pageno%>"  target="mainr">修 改</a>]</td>	
  </tr>
<%
  	repeat_rows = repeat_rows + 1
  rs.movenext
  wend
%>
</table>

<table  width="100%"  border="0" cellspacing="0" cellpadding="0" align=center>
<tr>
<td colspan=5>&nbsp;</td>
</tr>
<tr>
<td width="3%" height="26" align=right nowrap><input name="Chkall" type="checkbox" style="border:0;" onClick="CheckAll(this.form.ChkBox,this.form.Chkall)">
</td>
<td width="7%" align=center nowrap style="color:black;">[全选/反选]</td>
<td width="30%" nowrap>[<a style="cursor:hand;color:red;" onClick="javascript:Check(deladmin,<%=session("Adminid")%>)">删除</a>]
[<a href="adm_AddAdmin.asp" target="mainr"><font color="red">添加管理员</font></a>]
<input type="hidden" name="condition" value="<%=getcondition%>">
<input type="hidden" name="key" value="<%=getkey%>">
<input type="hidden" name="pageno" value="<%=getpageno%>"></td>
<td width="2%"></td>
<td width="58%">
<table  width="100%"  border="0" cellspacing="0" cellpadding="0" align=center>
<tr>
<td align=right nowrap style="color:black;">
[<%=pageno%>/<%=rs.pagecount%>]&nbsp;[共<font color="red"><%=rs_total%></font>条记录]&nbsp;
<%
if(pageno <> 1)then
%>
<a href="?condition=<%=getcondition%>&key=<%=getkey%>">第一页</a>
<%
End if
%>
<%
if(pageno <> 1)then
%>
&nbsp;<a href="?condition=<%=getcondition%>&key=<%=getkey%>&pageno=<%=(pageno-1)%>">上一页</a>&nbsp;
<%
end if
%>
<%
if(instr(pageno,cstr(rs.pagecount)) = 0)then
%>
<a href="?condition=<%=getcondition%>&key=<%=getkey%>&pageno=<%=(pageno+1)%>">下一页</a>
<%
end if
%>
<%
if(instr(pageno,cstr(rs.pagecount)) = 0)then
%>
&nbsp;<a href="?condition=<%=getcondition%>&key=<%=getkey%>&pageno=<%=rs.pagecount%>">最后一页</a>
<%
end if
rs.close
set rs = nothing
%></td>
</tr>
</table>
</td>
</tr>

</table>
</form>
<%
end if
%>
</body>
</html>
