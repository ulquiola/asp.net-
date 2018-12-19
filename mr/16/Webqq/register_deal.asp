<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
'获取用户信息
UserName=request.Form("UserName")				'为UserName变量赋值
Password1=request.Form("Password1")				'为Password1变量赋值
sex=request.form("sex")							'为sex变量赋值
birthday=request.form("birthday")
touxiang=request.form("touxiang")
if request.form("touxiang")="" then
	touxiang=0
end if
Email=request.form("Email")
'判断用户名称是否为空
if UserName<>"" then
    '创建记录集
	set rs=server.createobject("adodb.recordset")
	rs1="select * from tb_user where UserName='"&UserName&"'"
	rs.open rs1,conn,1,3
	if rs.eof and rs.bof then 
	'应用insert into 语句实现用户信息的添加
	rs1="insert into tb_user (UserName,Password1,sex,birthday,touxiang,Email) values('"&UserName&"','"&Password1&"','"&sex&"','"&birthday&"','"&touxiang&"','"&Email&"')"
		conn.execute(rs1)
		%>
			<script language="javascript">
				alert("用户注册成功！！"); //弹出提示对话框
				window.location.href="index1.asp";//跳转到指定的页面
				window.close();
			</script>
<%		end if 
end if
%>