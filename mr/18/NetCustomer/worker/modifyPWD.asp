<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->
<%
if request.Form("PWD_N1")<>"" then
newPWD=request.Form("PWD_N1")
UP="Update DB_Purview set password='"&newPWD&"' where userno='"&session("UserNO")&"'"
conn.execute(UP)
 %>
	<script language="javascript">
	alert("密码修改成功！");
	window.close();
	</script>
<%
end if
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT ywynumber, name, password FROM Q_worker WHERE name = '"&session("UserName")&"'"
rs.open sql,conn,1,3
session("userNO")=rs("ywynumber")
Set rs_userNO = Server.CreateObject("ADODB.Recordset")
sql_UN = "SELECT *  from DB_Purview INNER JOIN DB_workerinfo ON DB_Purview.userno=DB_workerinfo.ywynumber where DB_workerinfo.name ='"&session("UserName")&"'"
rs_userNO.open sql_UN,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>修改密码！</title>
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
if (form1.PWD_S.value=="")
{alert("请输入原密码！");form1.PWD_S.focus();return;}
if (form1.PWD_S.value != form1.hPWD_S.value)
{alert("您输入原密码不正确！");form1.PWD_S.value="";form1.PWD_S.focus();return;}
if (form1.PWD_N1.value =="")
{alert("请输入新密码！");form1.PWD_N1.focus();return;}
if (form1.PWD_N2.value=="")
{alert("请确认新密码！");form1.PWD_N2.focus();return;}
if (form1.PWD_N1.value!=form1.PWD_N2.value)
{alert("您两次输入的新密码不一致！");form1.PWD_N1.value="";
form1.PWD_N2.value="";form1.PWD_N1.focus();return;}
form1.submit();
}
</script>
<table width="400" height="242" border="0" cellpadding="-2" cellspacing="-2" background="../Images/add_manager.gif">
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
            <form ACTION="modifyPWD.asp" METHOD="POST" name="form1">
            
            <table width="71%" border="1" cellpadding="-2"  cellspacing="-2" bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr>
                <td bgcolor="#809EA4"><div align="center" class="style1">业务员名称：</div></td>
                <td><div align="left">&nbsp;
                    <input name="UserName" type="text" class="Sytle_text" id="UserName" value="<%=(rs.Fields.Item("name").Value)%>" readonly="yes">
                </div></td>
              </tr>
			  <tr>
                <td width="33%" bgcolor="#809EA4"><div align="center" class="style1">输入原密码：</div></td>
                <td width="67%"><div align="left">&nbsp;
                  <input name="PWD_S" type="password" class="Sytle_text" id="PWD_S" size="10">
                  <input name="hPWD_S" type="hidden" id="hPWD_S" value="<%=(rs.Fields.Item("password").Value)%>">
                </div></td>
                </tr>
              <tr>
                <td bgcolor="#809EA4"><div align="center"><span class="style1">输入新密码：</span></div></td>
                <td>                  <div align="left">&nbsp;
                  <input name="PWD_N1" type="password" class="Sytle_text" id="PWD_N1" size="10">
                  </div></td>
                </tr>
              <tr>
                <td bgcolor="#809EA4"><div align="center" class="style1">确认新密码：</div></td>
                <td><div align="left">&nbsp;
                  <input name="PWD_N2" type="password" class="Sytle_text" id="PWD_N2" size="10">
                </div></td>
              </tr>
              <tr>
                <td height="35" colspan="2"><div align="center">
                  <input name="Button2" type="button" class="Style_button" value="保存" onClick="mycheck()">
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