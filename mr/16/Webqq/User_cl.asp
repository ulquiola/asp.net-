<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
</head>
<body>
<%
user_login= Request.ServerVariables("URL")
If Request.QueryString<>"" Then user_login = user_login+"?"+Request.QueryString
username_value=CStr(Request.Form("UserName"))
application("UserName")=""
a_username=split(application("UserName"),",")
for i=0 to ubound(a_username)
	if a_username(i)=username_value then%>
	<script language="javascript">
	alert("对不起，该用户已在别处登录！");
	window.location.href="index1.asp";
	</script>
<%
response.End()
end if
next

If username_value <> "" Then
  user_action=""
  user_loginSuccessed="mainmenu.asp"
  user_loginfailed="index1.asp"
  set rsuser=Server.CreateObject("ADODB.Recordset")						'创建记录集
  rsuser.ActiveConnection = conn										
  rsuser.Source = "SELECT Username, Password1,state,id,touxiang"		'查询数据
  If user_action <> "" Then rsuser.Source = rsuser.Source & "," & user_action
  rsuser.Source = rsuser.Source & " FROM tb_user where Username='" & Replace(username_value,"'","''") &"' and Password1='" & Replace(Request.Form("Password1"),"'","''")&"'"
  '判断输入的信息是否正确
  rsuser.CursorType = 0					'定义游标类型
  rsuser.CursorLocation = 2				'定义游标访问区域
  rsuser.LockType = 3					
  rsuser.Open							'打开记录集
  If Not rsuser.eof or not rsuser.bof Then 	'查询是否有用户信息
    Session("username1")=username_value		'为Session("username1")变量赋值
	application.Lock()						
	application("username")=application("username")+","+username_value	'为application("username")变量赋值
	application.UnLock()
	response.Cookies("user_name") = username_value						'获取当前用户的
	response.Cookies("state") = rsuser.fields.item("state").value		'获取指定的字段信息
	response.Cookies("UserID") = rsuser.fields.item("id").value			'获取指定的字段信息
	Session("touxiang")=rsuser.fields.item("touxiang").value			'获取指定的字段信息
	Session("user_id")=rsuser.fields.item("id").value					'获取指定的字段信息
	Session("state")=rsuser.fields.item("state").value					'获取指定的字段信息
	if trim(application("OnLineUserID")) = "" then 
		application("OnLineUserID") = "S" & rsuser.fields.item("id").value & "E"
	else
	    strUserID =  ",S" & rsuser.fields.item("id").value & "E"
	    application("OnLineUserID") = replace(application("OnLineUserID"),strUserID,"")
	    application("OnLineUserID") = application("OnLineUserID") & ",S" & rsuser.fields.item("id").value&"E"
	End IF
    If (user_action <> "") Then
      Session("user_action") = CStr(rsuser.Fields.Item(user_action).Value)
    Else
      Session("user_action")=""											'清空 Session("user_action")变量
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then	
      user_loginSuccessed = Request.QueryString("accessdenied")		'为user_loginSuccessed变量赋值
    End If
    rsuser.Close
    Response.Redirect(user_loginSuccessed)					'跳转到指定的页面
  End If
  rsuser.Close
  Response.Redirect(user_loginfailed)						'跳转到指定的页面
End If
%>
</body>
</html>
