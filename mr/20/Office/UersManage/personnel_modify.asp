<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Conn/conn.asp" -->
<%
'查询员工信息
If Request.QueryString("ID")<>""then
session("ID")=Request.QueryString("ID")
end if
Set rs_personnel = Server.CreateObject("ADODB.Recordset")
sql_P="SELECT ID,UserName,name,PWD,purview,branch,job,sex,Email,Tel,Address"&_
" FROM Tab_User WHERE ID="&session("ID")&""
rs_personnel.open sql_p,conn,0,3
%>
<%
'修改员工信息
if request.Form("Name")<>"" then
	cname=request.Form("name")
	PWD=request.Form("PWD")
	purview=request.Form("purview")
	sex=request.Form("sex")
	tel=request.Form("tel")
	branch=request.Form("branch")
	job=request.Form("job")
	email=request.Form("email")
	address=request.Form("address")
	UP="Update Tab_User set name='"&cname&"',PWD='"&PWD&"',purview='"&_
	purview&"',sex='"&sex&"',tel='"&tel&"',branch='"&branch&"',job='"&job&_
	"',email='"&email&"',address='"&address&"' where ID='"&session("ID")&"'"
	conn.execute(UP)
	%>
	<script language="javascript">
	alert("数据修改成功！");
	opener.location.reload();
	window.close();
	</script>
<%end if%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>修改员工信息！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<script language="javascript">
function Mycheck()
{
if (form1.name.value=="")
{ alert("请输入员工姓名！");form1.name.focus();return;}
if (form1.PWD.value=="")
{ alert("请输入密码！");form1.PWD.focus();return;}
if (form1.tel.value=="")
{ alert("请输入员工的联系电话！");form1.tel.focus();return;}
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
.STYLE4 {
	font-size: 9pt;
	color: #FFFFFF;
}
-->
</style></head>
<body>
<table width="460" height="403" border="0" cellpadding="-2" cellspacing="-2" background="../Images/waichudeng.gif">
  <tr>
    <td width="817" height="349" valign="top"><table width="100%" height="102"  border="0" cellpadding="-2" cellspacing="-2">
      <tr>
        <td height="102"><table width="434" height="52" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="STYLE4">&nbsp;修改员工信息</span></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
          </tr>
        </table></td>
      </tr>
    </table>      
      <br>
      <form ACTION="personnel_modify.asp" METHOD="POST" name="form1">
        <table width="95%" height="222"  border="0" align="center" cellpadding="-2" cellspacing="-2">
          <tr>
            <td colspan="2"><span class="style1">&nbsp;用户名：<%=(rs_personnel("UserName"))%></span></td>
            <td class="style1" align="center">密码：</td>
            <td colspan="3"><span class="style1">
              <input name="PWD" type="password" class="Sytle_text" id="name" value="<%=rs_personnel("PWD")%>">
            </span></td>
          </tr>
		  <tr>
            <td width="11%" height="27"><div align="center" class="style1">姓名：</div></td>
            <td width="25%" height="27"><span class="style1">
              <input name="name" type="text" class="Sytle_text" id="username2" value="<%=(rs_personnel("Name"))%>" size="15">
            </span> </td>
            <td width="10%" height="27" class="style1"><div align="center">权限：</div></td>
            <td height="27"><select name="purview" id="purview">
              <option value="系统" <%If (Not isNull((rs_personnel("purview")))) Then If ("系统" = CStr((rs_personnel("purview")))) Then Response.Write("SELECTED") : Response.Write("")%>>系统</option>
              <option value="读写" <%If (Not isNull((rs_personnel("purview")))) Then If ("读写" = CStr((rs_personnel("purview")))) Then Response.Write("SELECTED") : Response.Write("")%>>读写</option>
              <option value="只读" <%If (Not isNull((rs_personnel("purview")))) Then If ("只读" = CStr((rs_personnel("purview")))) Then Response.Write("SELECTED") : Response.Write("")%>>只读</option>
            </select></td>
            <td width="9%" height="27" class="style1"><div align="center">性别：</div></td>
            <td width="29%" height="27"><span class="style1">
              <select name="sex" id="select">
                <option value="男" <%If (Not isNull((rs_personnel("sex")))) Then If ("男" = CStr((rs_personnel("sex")))) Then Response.Write("SELECTED") : Response.Write("")%>>男</option>
                <option value="女" <%If (Not isNull((rs_personnel("sex")))) Then If ("女" = CStr((rs_personnel("sex")))) Then Response.Write("SELECTED") : Response.Write("")%>>女</option>
              </select>
</span></td>
          </tr>
          <tr>
            <td height="27"><div align="center" class="style1">电话：</div></td>
            <td height="27"><input name="tel" type="text" class="Sytle_text" id="tel" value="<%=(rs_personnel("Tel"))%>" size="15"></td>
            <td width="10%" height="27" class="style1"><div align="center">部门：</div></td>
            <td width="16%" height="27"><select name="branch" id="branch">
              <option value="开发部" <%If (Not isNull((rs_personnel("branch")))) Then If ("开发部" = CStr((rs_personnel("branch")))) Then Response.Write("SELECTED") : Response.Write("")%>>开发部</option>
              <option value="人事部" <%If (Not isNull((rs_personnel("branch")))) Then If ("人事部" = CStr((rs_personnel("branch")))) Then Response.Write("SELECTED") : Response.Write("")%>>人事部</option>
              <option value="销售部" <%If (Not isNull((rs_personnel("branch")))) Then If ("销售部" = CStr((rs_personnel("branch")))) Then Response.Write("SELECTED") : Response.Write("")%>>销售部</option>
              <option value="文档部" <%if (Not isNull((rs_personnel("branch")))) then if ("文档部" = CStr((rs_personnel("branch")))) then response.Write("SELECTED") :response.Write("")%>>文档部</option>
            </select></td>
            <td width="9%" height="27"><div align="center"><span class="style1">职务：</span></div></td>
            <td height="27"><span class="style1">              <select name="job" id="job">
                <option value="总经理" <%If (Not isNull((rs_personnel("job")))) Then If ("总经理" = CStr((rs_personnel("job")))) Then Response.Write("SELECTED") : Response.Write("")%>>总经理</option>
                <option value="部门经理" <%If (Not isNull((rs_personnel("job")))) Then If ("部门经理" = CStr((rs_personnel("job")))) Then Response.Write("SELECTED") : Response.Write("")%>>部门经理</option>
                <option value="人力资源主管" <%If (Not isNull((rs_personnel("job")))) Then If ("人力资源主管" = CStr((rs_personnel("job")))) Then Response.Write("SELECTED") : Response.Write("")%>>人力资源主管</option>
                <option value="主任" <%If (Not isNull((rs_personnel("job")))) Then If ("主任" = CStr((rs_personnel("job")))) Then Response.Write("SELECTED") : Response.Write("")%>>主任</option>
                <option value="经理助理" <%If (Not isNull((rs_personnel("job")))) Then If ("经理助理" = CStr((rs_personnel("job")))) Then Response.Write("SELECTED") : Response.Write("")%>>经理助理</option>
                <option value="普通员工" <%If (Not isNull((rs_personnel("job")))) Then If ("普通员工" = CStr((rs_personnel("job")))) Then Response.Write("SELECTED") : Response.Write("")%>>普通员工</option>
                <option value="销售员" <%if (Not isNull((rs_personnel("job")))) then if ("销售员") = cstr((rs_personnel("job"))) then response.Write("SELECTED") : response.Write("")%>>销售员</option>
              </select>
</span></td>
          </tr>
          <tr>
            <td height="27" class="style1"><div align="center">Email：</div></td>
            <td height="27" colspan="5"><input name="email" type="text" class="Style_subject" id="email" value="<%=(rs_personnel("Email"))%>" size="15"></td>
          </tr>
          <tr>
            <td height="27" class="style1"><div align="center">地址：</div></td>
            <td height="27" colspan="5"><input name="address" type="text"
			 class="Style_subject" id="address" value="<%=(rs_personnel("Address"))%>"></td>
          </tr>
          <tr>
            <td colspan="6"><div align="center">
                <br>
                <input name="Button" type="button" class="Style_button" value="修改" onClick="Mycheck()">
                <input name="Submit2" type="reset" class="Style_button" value="重置">
                </div></td>
          </tr>
        </table>
      </form></td>
  </tr>
</table>
</body>
</html>
