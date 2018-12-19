<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
response.ContentType="text/html; charset=gb2312"			'编码转换
Dim conn													'声明变量
set conn=Server.CreateObject("ADODB.Connection")			'创建数据源
sql="provider=sqloledb;data source=(local);initial catalog=db_QQ;user id=sa;password=;"		'建立数据库连接
conn.open(sql)												'执行SQL语句
if request("UserName")<>"" then								'判断用户名是否为空
	session("UserName")=request("UserName")					'为session("UserName")变量赋值
end if
Set rs=Server.CreateObject("ADODB.RecordSet")				'创建记录集
sql1="select * from tb_user where UserName='"&session("UserName")&"'"	'查询数据
conn.execute(sql1)											'执行sql1语句
rs.open sql1,conn,1,3										'打开记录集
if rs.eof and rs.bof then							
	response.Write("祝贺您！用户名["&session("UserName")&"]没有被注册！")
	'弹出提示信息
else
	response.Write("很报歉！用户名["&session("UserName")&"]已经被注册！")
	'弹出提示信息
end if 
%>
