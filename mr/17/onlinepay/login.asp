<!--#include file="conn/conn.asp"-->
<%
if session("UserName")="" then			'判断用户名是否为空
%>
<% 
		session.Timeout=120				'设置Session对象的使用时限
	if request.Form("UserName")<>"" and request.Form("PWD")<>"" then	'判断用户名和密码是否为空
		session("UserName")=request.Form("UserName")					'将获取的用户名称赋值给Session变量
		session("PWD")=request.Form("PWD")								'将获取的用户密码赋值给Session变量
		sql="select UserName,PWD from tb_user where UserName='"&session("UserName")&"'"	'查询数据
		set rs=conn.execute(sql)														'执行SQL语句
		if rs.eof then 
	%>
			<script language="javascript">
			alert("您输入的用户名错误，请重新输入！");								//弹出提示对话框
			 </script> 
			 <%
			 session.Abandon()  
			 '删除所有存在Session对象中的对象
		else 
			if rs("PWD")=session("PWD") then									'判断用户密码是否相等
				session("PWD")=request.Form("pwd") '将用户密码赋值给Session变量%>							
				<script language="javascript">
				window.location.href="index.asp"					//跳转到指定的页面
				</script>
			 <%else%>
				<script language="javascript">
				 alert("您输入的用户密码错误，请重新输入！");		//弹出提示对话框
				</script>  		
				<%
				session.Abandon()
			end if
		end if
end if
%>
	<script language="javascript">
	function Mycheck()							//创建自定义函数
	{
	if(form1.username.value=="")				//判断用户名称的文本框的值是否为空
	{alert("请输入用户名后再进行登录!");form1.username.focus();return;}
	if(form1.pwd.value=="")						//判断用户密码的文本框的值是否为空
	{alert("用户密码不能为空!");form1.pwd.focus();return;}
	form1.submit();								//进行表单提交
	}
	</script>
	<style type="text/css">
<!--
.STYLE4 {
font-size: 9pt; color: #000000;
border:1px solid #B3B3B3;
}
.STYLE5 {
	font-size: 9pt;
	font-weight: bold;
	color: #606060;
}
-->
    </style>
	<style type="text/css">
<!--
body {
	margin-top: 0px;
	margin-bottom: 0px;
}
.STYLE6 {
	font-size: 10pt;
	color: #FF0000;
	font-weight: bold;
}

-->
    </style>
	<table width="217" height="109" border="0" align="center" cellpadding="0" cellspacing="0">
	  <tr>
		<td width="169" height="109">
		<form action="index.asp" name="form1" method="post">
				<table width="204" height="67" border="0" align="left" cellpadding="0" cellspacing="0">
				  <tr>
					<td width="68" height="28" valign="bottom"><div align="center">&nbsp;&nbsp;&nbsp;用户名</div></td>
					<td width="105" valign="bottom"><input name="username" type="text" class="STYLE4" id="username" size="15"></td>
				  </tr>
				  <tr>
					<td height="18"><div align="center">密 码</div></td>
					<td><input name="pwd" type="password" class="STYLE4" id="pwd" size="16" /></td>
				  </tr>
				  <tr>
					<td height="21" colspan="2" valign="bottom"><div align="center"><br>
					  &nbsp;
					  <input type="button" name="Submit2" value="登录" onClick="Mycheck();">
					  &nbsp;
					  <input type="reset" name="Submit3" value="重置">
					</div></td>
				  </tr>
	      </table>
<br><br><br><br><br><br>	            <p align="right" class="STYLE5" onClick="JScript:window.location.href='reg.asp';">用户注册</p>
		</form></td>
    </tr>
	</table>
	<%else%><table width="171" height="42" border="0" align="center" cellpadding="0" cellspacing="0">
			<tr>
			  <td width="66" valign="bottom"><div align="right" class="STYLE6"><%=session("UserName")%></div></td>
			  <td width="128" valign="bottom"><span class="STYLE4">&nbsp;欢迎您光临本网站</span></td>
			</tr>
			<tr>
			  <td height="21" colspan="2"><div align="center">
				  <input name="button1" type="button" value="退出" onClick="if(confirm('是否真的退出?')){window.location.href='exit.asp'}">
			  </div></td>
			</tr>
		  </table>
	<%end if%>
