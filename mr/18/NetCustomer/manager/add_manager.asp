<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->
<%
if request.Form("Name")<>"" then
	vName=request.Form("Name")
	pwd=request.Form("PWD")
	Purview=request.Form("purview")
	'检测用户输入的管理员是否存在
	set rs=Server.CreateObject("ADODB.RecordSet")
	sql="SELECT Name FROM DB_manager WHERE Name='" & vName & "'"
	rs.open sql,conn,1,3
	if rs.eof then
	'保存管理员信息
 		 Ins="Insert into DB_manager (name,pwd,purview) values('"&vname&"','"&pwd&"','"&purview&"')"
		 conn.execute(Ins)%>
		<script language="javascript">
	    alert("数据添加成功，继续添加！");
	    </script> 
	<%else%>
		<script language="javascript">
	    alert("该管理员信息已经存在！");
	    </script> 
	<%end if
end if %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>网上客户管理系统--管理员登录！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	background-image:  url(../Images/bg.gif);
}
.style1 {color: #FFFFFF}
.style2 {color: #a2bcc5}
-->
</style></head>

<body>
<script language="javascript">
function mycheck(){
if (form1.Name.value=="")
{alert("请输入管理员姓名！");form1.Name.focus();return;}
if(form1.PWD.value=="")
{alert("请输入管理员密码！");form1.PWD.focus();return;}
form1.submit();
}
</script>
<table width="400" height="242" border="0" cellpadding="-2" cellspacing="-2"
 background="../Images/add_manager.gif">
  <tr>
    <td valign="top"><table width="400" height="271" cellpadding="-2" cellspacing="-2">
      <tr>
        <td width="11" height="85">&nbsp;</td>
        <td width="373">&nbsp;</td>
        <td width="14"><span class="style2"></span></td>
      </tr>
      <tr>
        <td height="144">&nbsp;</td>
        <td valign="top">
  
          <div align="center">
            <form name="form1" method="POST" action="add_manager.asp">
            
            <table width="80%" height="136" border="1" cellpadding="-2" 
			 cellspacing="-2" bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr>
                <td width="28%" bgcolor="#809EA4"><div align="center" class="style1">
				管理员名称：</div></td>
                <td width="72%">
                  <div align="left">
                    &nbsp;
                    <input name="Name" type="text" class="Sytle_auto" id="Name">
                  </div></td>
                </tr>
              <tr>
                <td bgcolor="#809EA4"><div align="center"><span class="style1">&nbsp;管理员密码：
				</span></div></td>
                <td><div align="left">&nbsp;
                    <input name="PWD" type="password" class="Sytle_auto" id="PWD">
</div></td>
                </tr>
              <tr>
                <td bgcolor="#809EA4" class="style1"><div align="center">管理员权限：</div></td>
                <td><div align="left">&nbsp;
                  <select name="purview" id="purview">
                    <option value="系统">系统　　　</option>
                    <option value="一般" selected>一般　　　</option>
                  </select>
                </div></td>
              </tr>
              <tr>
                <td colspan="2"><div align="center">
                  <input name="Button" type="button" class="Style_button" value="保存" onClick="mycheck()">
                  <input name="Submit2" type="reset" class="Style_button" value="重置">
                  <input name="Button" type="button" class="Style_button" value="关闭"
				   onClick="javascrip:opener.location.reload();window.close()">
                </div></td>
                </tr>
              </table>
            </form>
            </div></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
