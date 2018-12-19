<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Conn/conn.asp" -->
<%
set rs_user=server.CreateObject("ADODB.Recordset")
sql="select * from Tab_User where username='"&Session("UserName")&"'"
rs_user.open sql,conn,1,3
%>
<% if trim(rs_user("purview"))<>"系统" then
%>
 <script language="javascript">
	alert("对不起您没有添加新员工的权限!!");
	window.location.href='../main.asp';
	</script>
<% end if%>
<%
if request.Form("username")<>""then
'检测用户名是否存在
	Set rs= Server.CreateObject("ADODB.Recordset")
	sql="SELECT UserName FROM Tab_User WHERE UserName='"&Request.Form("UserName")&"'"
	rs.open sql,conn,1,3
	if rs.eof then 
	    '保存员工信息
		username=request.Form("username")
		cname=request.Form("name")
		branch=request.Form("branch")
		PWD=request.Form("PWD")
		job=request.Form("job")
		tel=request.Form("tel")
		purview=request.Form("purview")
		sex=request.Form("sex")
		email=request.Form("email")
		address=request.Form("address")
		Ins="Insert into Tab_User (username,name,branch,PWD,job,tel,purview,sex,"&_
		"email,address) values('"&username&"','"&cname&"','"&branch&"','"&PWD&"','"&_
		job&"','"&tel&"','"&purview&"','"&sex&"','"&email&"','"&address&"')"
		conn.execute(Ins) 
		%>
		<script language="javascript">
		alert("员工信息添加成功！");
	    window.location.href='personnel_add.asp';
		</script>
	<%else%>
		<script language="javascript">
		alert("该员工信息已经存在！");
		</script>
	<% 
	    end if
		end if
	%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>添加新员工</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<script language="javascript">
function Mycheck()
{
if (form1.username.value=="")
{ alert("请输入员工姓名！");form1.username.focus();return;}
if(form1.name.value=="")
{alert("请输入姓名??");form1.name.focus();return;}
if (form1.PWD.value=="")
{ alert("请输入密码！");form1.PWD.focus();return;}
if (form1.tel.value=="")
{ alert("请输入员工的联系电话！");form1.tel.focus();return;}
if(form1.email.value=="")
{alert("请输入E-mail地址??");form1.email.focus();return;}
if(form1.email.value!="" && (form1.email.value.indexOf('@',0)==-1|| form1.email.value.indexOf('.',0)==-1))
{alert("您输入的E-mail地址不对！");form1.email.focus();return;}
if (form1.address.value=="")
{ alert("请输入员工的住址！");form1.address.focus();return;}
form1.submit();
}
</script>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
.style1 {color: #C60001}
.style2 {color: #669999}
.STYLE8 {
	font-size: 9pt;
	color: #000000;
}
-->
</style>
</head>
<body>
<table width="814" height="505" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="488" valign="top" background="../Images/main.gif"><table width="89%" height="84%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td>
		
		<table width="557" height="401" border="0" cellpadding="-2" cellspacing="-2" background="../images/bbs/bbs_add.gif">
  <tr>
    <td width="817" height="349" valign="top"><table width="100%" height="104"  border="0" cellpadding="-2" cellspacing="-2">
      <tr>
        <td width="6%" height="51" valign="bottom"><div align="right"><img src="../Images/add.gif" width="20" height="18"></div></td>
        <td width="94%" valign="bottom">&nbsp;<img src="../Images/yuangong.jpg" width="101" height="17"></td>
      </tr>
      <tr>
        <td height="53" colspan="2">&nbsp;</td>
      </tr>
    </table>      
      <form ACTION="personnel_add.asp" METHOD="POST" name="form1">
        <table width="100%" height="177"  border="0" align="center" cellpadding="-2" cellspacing="-2">
		  <tr>
            <td width="13%" height="27">
              <div align="center" class="STYLE8">用户名：</div></td><td width="28%"><input name="username" type="text" class="Sytle_text" id="username"></td>
            <td width="14%"><div align="center" class="STYLE8">部门：</div></td>
            <td><select name="branch" id="select4">
              <option value="开发部" selected>开发部</option>
              <option value="人事部">人事部</option>
              <option value="销售部">销售部</option>
              <option value="文档部">文档部</option>
            </select></td>
          </tr>
		　<tr>
            <td><div align="center" class="STYLE8">员工姓名：</div></td>
            <td><input name="name" type="text" class="Sytle_text" id="name"></td>
            <td><div align="center" class="STYLE8">性别：</div></td>
            <td><select name="sex" id="sex">
              <option value="男" selected>男</option>
              <option value="女">女</option>
            </select></td>
            </tr>
          <tr>
            <td height="27" valign="top">
              <div align="center" class="STYLE8">密码：</div></td>
            <td><input name="PWD" type="password" class="Sytle_text" id="PWD2"></td>
            <td><div align="center" class="STYLE8">职务：</div></td>
            <td><select name="job" id="job">
              <option value="总经理" selected>总经理</option>
              <option value="部门经理">部门经理</option>
              <option value="人力资源主管">人力资源主管</option>
              <option value="主任">主任</option>
              <option value="经理助理">经理助理</option>
              <option value="普通员工">普通员工</option>
              <option value="销售员">销售员</option>
            </select></td>
          </tr>
          <tr>
            <td height="27" valign="top">
              <div align="center" class="STYLE8">电话：</div></td>
            <td><input name="tel" type="text" class="Sytle_text" id="tel3"></td>
            <td><div align="center" class="STYLE8">权限：</div></td>
            <td width="45%"><select name="purview" id="select">
              <option value="系统" selected>系统</option>
              <option value="读写">读写</option>
              <option value="只读">只读</option>
            </select></td>
          </tr>
          <tr>
            <td height="27" valign="top"><div align="center" class="STYLE8">Email：</div></td>
            <td colspan="3"><input name="email" type="text" class="Style_subject" id="email"></td>
          </tr>
          <tr>
            <td height="27" valign="top"><div align="center" class="STYLE8">地址：</div></td>
            <td colspan="3"><input name="address" type="text" class="Style_subject" id="address"></td>
          </tr>
          <tr>
            <td colspan="4"><div align="center">
                <input name="Button" type="button" class="Style_button" value="提交" onClick="Mycheck()">
                <input name="Submit2" type="reset" class="Style_button" value="重置">                
                </div></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>

&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
