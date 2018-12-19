<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<!--#include file="../conn/conn.asp"-->
<%
set rs_user=server.CreateObject("ADODB.Recordset")
sql="select * from Tab_User where username='"&Session("UserName")&"'"
rs_user.open sql,conn,1,3
%>
<% if trim(rs_user("purview"))<>"系统" then
%>
 <script language="javascript">
	alert("对不起您没有添加详细信息的权限!!");
	window.location.href='../main.asp';
	</script>
<% end if%>
<%
if request.Form("name11")<>"" then
'检索用户名是否存在
Set rs= Server.CreateObject("ADODB.Recordset")
	sql="select name11 from Tab_tongxunadd where name11='"&Request.Form("name11")&"'"
	rs.open sql,conn,1,3
	if rs.eof then 
	name11=request.Form("name11")
	name1=request.Form("name1")
	birthday=request.Form("birthday")
	sex=request.Form("sex")
	hy=request.Form("hy")
	dw=request.Form("dw")
	department=request.Form("department")
	zw=request.Form("zw")
	sf=request.Form("sf")
	cs=request.Form("cs")
	phone=request.Form("phone")
	phone1=request.Form("phone1")
	email=request.form("email")
	postcode=request.Form("postcode")
	OICQ=request.Form("OICQ")
	family=request.Form("family")
	address=request.Form("address")
	remark=request.Form("remark")
	sql="insert into Tab_tongxunadd(name11,name1,birthday,sex,hy,dw,department,zw,sf,cs,phone,phone1,email,postcode,OICQ,family,address,remark) values('"&name11&"','"&name1&"','"&birthday&"','"&sex&"','"&hy&"','"&dw&"','"&department&"','"&zw&"','"&sf&"','"&cs&"','"&phone&"','"&phone1&"','"&email&"','"&postcode&"','"&OICQ&"','"&family&"','"&address&"','"&remark&"')"
	conn.execute(sql)
%>
<script language="javascript">
alert("通讯详细信息添加成功!!");
window.location.href='tongxun_index.asp';
</script>
<%else%>
<script language="javascript">
alert("该用户信息已经存在！");
window.location.href='tongxun_index.asp';
</script>
<%
end if
end if
%>