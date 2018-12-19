<!--#include file="Conn.asp"-->
<!--#include file="Function.asp"-->
<!--#include file="inc.asp"-->
<%If Request.Form("action")="login" then
If ChkPost=false Then
Response.Write("<script language=JavaScript>alert('请不要重外部提交信息！');document.location.href='index.asp';</script>")
Response.End
End If
Dim UserName,UserPwd,Sql,Rs,IP
UserName=Trim(Request.Form("UserName"))
UserPwd=Trim(Request.Form("UserPwd"))
If UserName="" or UserPwd="" Then
Response.Write("<script language=JavaScript>alert('请输入您的用户名和密码！');document.location.href='javascript:window.history.go(-1);';</script>")
Response.End
End If
UserName=Checkstr(UserName)

Sql="Select * From tb_Users Where UserName='"&UserName&"' and UserPwd='"&UserPwd&"'"
Set Rs=server.CreateObject("adodb.recordset")
If Not IsObject(Conn) Then ConnectionDatabase()
Rs.open Sql,Conn,1,3
If Rs.Bof Or Rs.Eof Then
Rs.close
Set Rs=Nothing
CloseDatabase()
Response.Write("<script language=JavaScript>alert('您的用户名和密码错误！');document.location.href='javascript:window.history.go(-1);';</script>")
Response.End
Else
If Request.ServerVariables("HTTP_X_FORWARDED_FOR")<>"" then 
IP=Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
Else
IP=Request.ServerVariables("REMOTE_ADDR")
End If
Rs("UserLogins")=Rs("UserLogins")+1
Rs("UserLastTime")=Now()
Rs("UserLastIp")=IP
Session("UserName")=Rs("UserName")
Session("UserLogins")=Rs("UserLogins")
Session("UserAdmin")=Rs("UserType")
Session("GroupID")=Rs("GroupID")
Rs.UpDate
Rs.close
Set Rs=Nothing
CloseDatabase()
Response.Write("<script language=JavaScript>document.location.href='Index.asp';</script>")
End If

ElseIf Request("action")="logout" Then
Session.Abandon()
Response.Redirect("index.asp")
Else
Response.Redirect("index.asp")
End if
%>
