<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Conn/conn.asp"-->
<%
getcondition = replace(trim(request("condition")),"'","''")
getkey = replace(trim(request("key")),"'","''")
if(getkey = "")then
rssql = "select * from tb_StuResult order by ���� asc"
else
	rssql = "select * from tb_StuResult where ���� like '%"&getkey&"%' order by ���� asc"
end if
set rs = server.createobject("adodb.recordset") 
rs.open rssql,conn,1,3
rs.pagesize = 5
'ʵ�ַ�ҳ
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
<title>��ѯ�����ɼ�</title>
<link href="../Css/style.css" rel="stylesheet">
<script language="javascript" src="js/js.js"></script>
<style type="text/css">
<!--
.style1 {color: #FFFFFF}
body {
	margin-left: 0px;
	margin-top: 0px;
}
-->
</style>
</head>
<body>
  <form name="queryuser" method="post" action="adm_Mainright.asp">
  <table width="100%"  border="0" cellspacing="0" cellpadding="2" align=center>
  <tr>
  <td height="27" colspan=4>&nbsp;</td>
  </tr>
  <tr>
    <td width="28%" height="26" align=right nowrap style="color:black;">&nbsp;</td>
    <td width="13%" align = right nowrap style="color:black;">���뿼�ţ�</td>
    <td width="24%"><input name="key" type="text" class="txt_grey" value="<%=Server.HTMLEncode(Replace(getkey,"''","'"))%>"></td>
    <td width="35%"><input name="submit" type="submit" class="btn_grey" value="��ѯ"></td>
    </tr>
	</table>
</form>
<form name="delresult" method="post">
<table width="98%"  border="1" cellspacing="0" cellpadding="0" align="center"
 bordercolor="#FFFFFF" bordercolordark="#1985DB" bordercolorlight="#FFFFFF"> 
  <tr>
  	<td width="9%" height="27"  align="center" nowrap style="color:black;">ѡ��</td>
	<td width="20%" height="27"  align="center" nowrap style="color:black;">����ID��</td>
    <td width="18%" height="27"  align="center" nowrap style="color:black;">����</td>
    <td width="12%" height="27"  align="center" nowrap style="color:black;">����</td>
	<td width="19%" height="27"  align="center" nowrap style="color:black;">����</td>
	<td width="10%" height="27"  align="center" nowrap style="color:black;">�Ա�</td>
	<td width="12%" height="27"  align="center" nowrap style="color:black;">�������</td>
	</tr>
  <%
  if(rs.eof)then
  %>
  <tr><td height="27" colspan=8 align=center nowrap><font color="red">û�з��������ļ�¼��</font></td></tr>	
  <%
  else
  	repeat_rows = 0
 	while(repeat_rows < rs.pagesize) and (not rs.eof)
  %>
  <tr>
    <td height="27"  align=center nowrap>
    <input type="checkbox" name="ChkBox" value="<%=server.htmlencode(rs("ID"))%>"  style="border:0;"></td>
	<td height="27"  align=center nowrap style="color:black;"><%=(rs("����ID��"))%>&nbsp;</td>
	<td  align=center nowrap><%=server.htmlencode(rs("����"))%>&nbsp;</td>
	<td height="27"  align=center style="color:black;"><%=server.htmlencode(rs("����"))%>&nbsp;</td>
	<td height="27"  align=center nowrap style="color:black;"><%=rs("����")%>&nbsp;</td>
	<td height="27"  align=center nowrap style="color:black;"><%=(rs("�Ա�"))%>&nbsp;</td>
	<td height="27"  align=center nowrap style="color:black;"><%=(rs("�������"))%><a href="tre_/tre_reg.asp?id=<%=rs("id")%>"></a>&nbsp;</td>
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
<td width="7%" align=center nowrap style="color:black;">[ȫѡ/��ѡ]</td>
<td width="30%" nowrap>[<a style="cursor:hand;color:red;" onClick="javascript:Check(delresult)">ɾ��</a>]</td>
<td width="2%"></td>
<td width="58%">
<table  width="100%"  border="0" cellspacing="0" cellpadding="0" align=center>
<tr>
<td align=right nowrap style="color:black;">
[<%=pageno%>/<%=rs.pagecount%>]&nbsp;[��<font color="red"><%=rs_total%></font>����¼]&nbsp;
<%
if(pageno <> 1)then
%>
[<a href="adm_Mainright.asp?condition=<%=getcondition%>&key=<%=getkey%>">��һҳ</a>
<%
End if
%>
<%
if(pageno <> 1)then
%>
&nbsp;<a href="adm_Mainright.asp?condition=<%=getcondition%>&key=<%=getkey%>&pageno=<%=(pageno-1)%>">��һҳ</a>&nbsp;
<%
end if
%>
<%
if(instr(pageno,cstr(rs.pagecount)) = 0)then
%>
<a href="adm_Mainright.asp?condition=<%=getcondition%>&key=<%=getkey%>&pageno=<%=(pageno+1)%>">��һҳ</a>
<%
end if
%>
<%
if(instr(pageno,cstr(rs.pagecount)) = 0)then
%>
&nbsp;<a href="adm_Mainright.asp?condition=<%=getcondition%>&key=<%=getkey%>&pageno=<%=rs.pagecount%>">���һҳ</a>]
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
