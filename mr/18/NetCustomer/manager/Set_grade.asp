<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/conn.asp" -->
<%'查询客户等级信息
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "SELECT * FROM DB_khgrade"
rs.open sql,conn,1,3
%>
<%
'修改客户等级信息
if request.Form("A")<>"" then 
	A=request.Form("A")
	B=request.Form("B")
	C=request.Form("C")
	Up="Update DB_khgrade set AClass="&A&",BClass="&B&",CClass="&C
	conn.execute(Up)%>
	<script language="javascript">
	  alert("客户等级设置成功！");
	  </script>
<% end if %>
	  

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>管理员登录--客户等级设置！</title>
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
if (form1.A.value=="")
{alert("请输入A类的销售数量最大值！");form1.A.focus();return;}
if (form1.B.value=="")
{alert("请输入B类的销售数量最大值！");form1.B.focus();return;}
if (form1.C.value=="")
{alert("请输入C类的销售数量最大值！");form1.C.focus();return;}
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
            <form ACTION="set_grade.asp" METHOD="POST" name="form1">
            
            <table width="80%" border="1" cellpadding="-2"  cellspacing="-2" bordercolor="#669999" bordercolordark="#FFFFFF">
              <tr>
                <td width="17%" bgcolor="#809EA4"><div align="center" class="style1">A类：</div></td>
                <td width="83%"><div align="left">&nbsp;数量大于 0 小于等于
   				 <input name="A" type="text" class="Sytle_auto" id="A" value="<%=rs("Aclass")%>" size="10">
 				 的客户</div></td>
                </tr>
              <tr>
                <td bgcolor="#809EA4"><div align="center"><span class="style1">B类：</span></div></td>
                <td>
                  <div align="left">&nbsp;数量大于A类小于等于
                      <input name="B" type="text" class="Sytle_auto" id="B" value="<%=rs("Bclass")%>" size="10">
                  的客户</div></td>
                </tr>
              <tr>
                <td bgcolor="#809EA4"><div align="center" class="style1">C类：</div></td>
                <td><div align="left">&nbsp;数量大于B类小于等于
				<input name="C" type="text" class="Sytle_auto" id="C" value="<%=rs("Cclass")%>" size="10">
 				 的客户</div></td>
              </tr>
              <tr>
                <td bgcolor="#809EA4"><div align="center" class="style1">D类：</div></td>
                <td><div align="left">&nbsp;数量大于D类的的客户</div></td>
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

