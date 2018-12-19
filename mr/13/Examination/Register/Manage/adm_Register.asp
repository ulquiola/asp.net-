<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../../Conn/conn.asp"-->
<%
if Session("AdminName")="" then
%>
<script language="javascript">alert("请先登录，再访问后台管理页");window.location.href="../../Manage/login.asp";</script>
<%
end if
%>
<%
set rs = server.createobject("adodb.recordset")
getcondition = replace(trim(request("condition")),"'","''")
getkey = replace(trim(request("key")),"'","''")
if(getkey = "")then
	rssql = "select * from tb_Student order by UserID asc"
elseif(InStr(getcondition,"JoinTime") > 0)Then
	If(InStr(getkey,"-") > 0)Then
		darray = Split(getkey,"-")
		getkey = ""
	For a=0 To Ubound(darray)
		If(len(darray(a)) = 1 )Then
				darray(a) = "0"&darray(a)
		End If
			getkey = getkey&darray(a)&"-"
	Next
		getkey = Mid(getkey,1,len(getkey)-1)
	End If
	rssql = "select * from tb_Student Where format("&getcondition&",'yyyy-mm-dd') like '%"&getkey&"%' order by Jointime desc"
else
	rssql = "select * from tb_Student Where UserID like '%"&getkey&"%' order by UserID desc"
End If
rs.open rssql,conn,1,3
rs.pagesize = 5 
'实现分页
if rs.eof then
	rs_total = 0
else
	rs_total = rs.recordcount
end if

dim pageno
	getpageno = replace(trim(request("pageno")),"'","")
if(getpageno = "")then
	pageno = 1
else
	pageno = getpageno
End if
if(not rs.eof)then
	rs.absolutepage = pageno
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>管理注册学生</title>
<link href="../Css/style.css" rel="stylesheet">
<script language="javascript" src="../Js/js.js"></script>
<style type="text/css">
<!--
.style1 {
	color: #FFFFFF;
	font-weight: bold;
}
-->
</style>
</head>
<body  topmargin="0" leftmargin="0">

  <form name="querystu" method="post" action="adm_Register.asp">
  <table width="100%"  border="0" cellspacing="0" cellpadding="0" align=center>
  <tr align="left">
  <td height="80" colspan=5 background="../../Manage/images/studentinfo.gif">&nbsp;</td>
  </tr>
  <tr>
    <td width="29%" height="26" align=right nowrap style="color:black;">查询条件：</td>
    <td width="17%">
	<select name="condition" class="wenbenkuang" style="width:150;border:1px solid;">
	<option value="UserID" <%If(InStr(getcondition,"UserID") > 0)Then%>selected<%End If%>>考号</option>
	<option value="JoinTime" <%If(InStr(getcondition,"JoinTime") > 0)Then%>selected<%End If%>>注册时间</option>
	</select>
	</td>
    <td width="9%" align = right nowrap style="color:black;">关键字：</td>
    <td width="24%">
		<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0" align="center">
			<tr>
				<td>
				<input name="key" type="text" class="txt_grey" value="<%=Server.HTMLEncode(Replace(getkey,"''","'"))%>">
				</td>
			</tr>
		</table>	
	</td>
    <td width="21%"><input name="submit" type="submit" class="btn_grey" style="border:1px solid;" value="查询"></td>
    </tr>
	</table>
</form>
<form name="delstu" method="post">
<table width="89%"  border="1" cellspacing="0" cellpadding="0" align="center"
 bordercolor="#FFFFFF" bordercolordark="#1985DB" bordercolorlight="#FFFFFF"> 
  <tr>
  	<td width="7%" height="27"  align="center" nowrap style="color:black;">选 项</td>
    <td width="21%" height="27"  align="center" nowrap style="color:black;">考号</td>
	<td width="22%" height="27"  align="center" nowrap style="color:black;">密码</td>
	<td width="30%" height="27"  align="center" nowrap style="color:black;">注册时间</td>
	</tr>
  <%
  if rs.eof and rs.bof  then
  %>
  <tr><td height="27" colspan=6 align=center nowrap><font color="red">没有符合条件的记录！</font></td></tr>	
  <%
  else
  repeat_rows = 0
  while(repeat_rows < rs.pagesize) and (not rs.eof)
  %>
  <tr>
    <td height="27"  align=center nowrap>
    <input type="checkbox" name="ChkBox" value="<%=server.htmlencode(rs("userID"))%>"  style="border:0;">	</td>
	<td height="27"  align=center nowrap style="color:black;"><%=rs("UserID")%></td>
	<td height="27"  align=center nowrap style="color:black;"><%=server.htmlencode(rs("PWD"))%></td>
	<td height="27"  align=center nowrap style="color:black;"><%=Mid(rs("JoinTime"),1,InStr(rs("JoinTime")," "))%></td>
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
<td width="30%" nowrap>[<a style="cursor:hand;color:red;" onClick="javascript:Check(delstu)">删除</a>]
<input type="hidden" name="condition" value="<%=getcondition%>">
<input type="hidden" name="key" value="<%=getkey%>">
<input type="hidden" name="pageno" value="<%=getpageno%>">
</td>
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
