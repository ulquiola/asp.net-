<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!-- #include file="Conn/Conn.asp" -->
<!--包含数据库连接文件-->
<% 
Session.Timeout=30
function filter_Str(InString)
'***********************************
'功能：过滤输入字符串中的危险符号
'调用方法：filter_Str("String")
'***********************************
	NewStr=Replace(InString,"'","''")
	NewStr=Replace(NewStr,"<","&lt;")
	NewStr=Replace(NewStr,">","&gt;")
	NewStr=Replace(NewStr,"chr(60)","&lt;")
	NewStr=Replace(NewStr,"chr(37)","&gt;")
	NewStr=Replace(NewStr,"""","&quot;")
	NewStr=Replace(NewStr,";",";;")
	NewStr=Replace(NewStr,"--","-")
	NewStr=Replace(NewStr,"/*"," ")
	NewStr=Replace(NewStr,"%"," ")
	filter_Str=NewStr
end function
if Request.Form("UID")<>"" and request.Form("PWD")<>"" then
	UID=filter_Str(request.Form("UID"))
	PWD=filter_Str(request.Form("PWD"))
	sql="select UserName,PWD from tb_User where UserName='"&UID&"'"
	set rs=conn.execute(sql)
	if rs.eof then
	%>
  		<script language="javascript">
  		alert("您输入的用户名错误，请重新输入！");
		window.location.href="Login.asp";
 		 </script> 
	<%
	Session("HS")="User"
	else 
		if rs("PWD")=PWD then
			Session("UserName")=UID
	lastlogintime=now()
	UP="Update tb_user set logintimes=logintimes+1,lastlogintime='"&lastlogintime&"' where UserName='"&UID&"'"
	conn.execute(UP)
	%>
	<script language=javascript>
 alert("登录成功！");
parent.leftwindow.location.reload();
window.location.href="main.asp";
</script>
	<%
 		 else
	%>
   			<script language="javascript">
			 alert("您输入的用户密码错误，请重新输入！");
			 window.location.href="Login.asp";
  			</script>  		
		<%
		Session("HS")="User"
		end if
	end if
end if
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

