<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!-- #include file="Conn/Conn.asp" --><!--包含数据库连接文件-->
<!--包含数据库连接文件-->
<% 
Session.Timeout=30
function filter_Str(InString)
'***********************************
'功能：过滤输入字符串中的危险符号
'调用方法：filter_Str("String")
'***********************************
	NewStr=Replace(InString,"'","''")					'将'替换成''
	NewStr=Replace(NewStr,"<","&lt;")					'将<替换成&lt;
	NewStr=Replace(NewStr,">","&gt;")					'将>替换成&gt;
	NewStr=Replace(NewStr,"chr(60)","&lt;")				'将chr(60)替换成&lt;
	NewStr=Replace(NewStr,"chr(37)","&gt;")				'将chr(37) 替换成&gt;
	NewStr=Replace(NewStr,"""","&quot;")				'将""替换成&quot;
	NewStr=Replace(NewStr,";",";;")						'将; 替换成;;
	NewStr=Replace(NewStr,"--","-")						'将--替换成-
	NewStr=Replace(NewStr,"/*"," ")						'将/*替换成" "
	NewStr=Replace(NewStr,"%"," ")						'将%替换成" "
	filter_Str=NewStr									'将NewStr字符串赋给filter_Str
end function
if Request.Form("UID")<>"" and request.Form("PWD")<>"" then								'判断输入的用户名和密码是否为空
	UID=filter_Str(request.Form("UID"))													'获取表单元素UID的值并赋给UID变量
	PWD=filter_Str(request.Form("PWD"))													'获取表单元素PWD的值并赋给PWD变量	
	sql="select UserName,PWD from tb_User where UserName='"&UID&"' And Status='管理员'"	'查询数据
	set rs=conn.execute(sql)							'执行sql语句
	if rs.eof then
	%>
  		<script language="javascript">
  		alert("您输入的管理员名称错误，请重新输入！");	//弹出提示对话框
	     window.location.href="index.asp";				//跳转到指定的ASP页面
 		 </script> 
	<%
	Session("HS")="Manager"						'为Session("HS")变量赋值
	else 
		if rs("PWD")=PWD then					'判断输入的用户密码是否正确
			Session("UserName")=UID				'为Session("UserName")变量赋值
			Session("Status")="管理员"			'为Session("Status")变量赋值
	lastlogintime=now()							'获取当前系统的日期和时间
	UP="Update tb_user set logintimes=logintimes+1,lastlogintime='"&lastlogintime&"' where UserName='"&UID&"' And Status='管理员'"
	'应用Update语句更新指定的记录
	conn.execute(UP)								'执行UP语句
	%>
<script language=javascript>
 alert("管理员登录成功！");							//弹出提示对话框
//parent.leftwindow.location.reload();				//刷新指定的页面
parent.location.href="manage_main.asp";				//跳转到指定的ASP页面
</script>

		<%else%>
   			<script language="javascript">
			 alert("您输入的管理员密码错误，请重新输入！");			//弹出提示对话框
  			 window.location.href="index.asp";						//跳转到指定的ASP页面
  			</script>  		
		<%
		Session("HS")="Manager"										'为Session("HS")变量赋值
		end if
	end if
end if
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">