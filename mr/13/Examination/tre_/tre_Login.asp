<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Conn/conn.asp"-->
<% 
'首先判断输入的验证码是否正确
dim verifycode,verifycode2
verifycode=trim(Request.Form("验证码"))
verifycode2=trim(Request.Form("验证码2"))
if verifycode<>verifycode2 then
response.write"<SCRIPT language=JavaScript>alert('您输入的验证码不正确。');"
response.write"location.href='../index.asp'</SCRIPT>"
else
session("verifycode")=""
getstuid = Replace(Request.Form("用户名"),"'","")
set rs1=server.CreateObject("adodb.recordset")
sql1="select * from tb_count where UserID='"&getstuid&"'"
rs1.open sql1,conn,1,3
	if rs1.eof then
		getstuid = Replace(Request.Form("用户名"),"'","")
		getpass = Replace(Request.Form("密码"),"'","")
		Sql = "Select * from tb_Student Where UserID = '"&getstuid&"' and PWD = '"&getpass&"'"
		Set rs = conn.Execute(Sql)
		If(rs.Eof)Then
			Response.Write("<script>alert('用户名不正确或密码错误！');history.back();</script>")
			Response.End()
		Else
			rs1.addnew
			rs1("UserID")=getstuid
			rs1("counts")=1
			rs1("datatime")=now()
			rs1("stoptime")=dateadd("m",6,now())
			rs1.update
			Session("StuID") = getstuid
			Session("PassWord") = getpass
			Response.Redirect("../main.asp")
		End If
	else
		if not rs1.eof and datediff("m",rs1("datatime"),now())>6 then
		Sql2 = "delete * from tb_Student Where UserID = '"&getstuid&"'"
		Set rs2 = conn.Execute(Sql2)
		response.write("<script>alert('用户名过期！');history.back();</script>")
		response.End()
		else
		getstuid = Replace(Request.Form("用户名"),"'","")
		getpass = Replace(Request.Form("密码"),"'","")
		Sql = "Select * from tb_Student Where UserID = '"&getstuid&"' and PWD = '"&getpass&"'"
		Set rs = conn.Execute(Sql)
		If(rs.Eof)Then
			Response.Write("<script>alert('用户名不正确或密码错误！');history.back();</script>")
			Response.End()
		Else
			
			Session("StuID") = getstuid
			Session("PassWord") = getpass
			Response.Redirect("../main.asp")
		End If
		rs.Close
		Set rs = Nothing
		End If
	end if
end if
%>
