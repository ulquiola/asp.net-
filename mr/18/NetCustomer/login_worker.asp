<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!-- #include file="Connections/conn_login.asp" -->  <!--包含数据库连接文件-->
<% session.Timeout=120 %>
<%
if request.Form("UserName")<>"" and request.Form("PWD")<>"" then
session("UserName")=request.Form("UserName")
session("PWD")=request.Form("PWD")
sql="select Name,PassWord from Q_worker where name='" & session("UserName")&"'"
set rs=conn.execute(sql)
if rs.eof then
  %>
  <script language="javascript">
  alert("您输入的操作员名称错误，请重新输入！");
  window.location.href="login_worker.asp"
  </script> 
  <%
else 
	if rs("password")=session("PWD") then
   %>
     <script language="javascript">
      window.location.href="worker/default.asp"
  	 </script>
  <%else%>
   <script language="javascript">
  alert("您输入的操作员密码错误，请重新输入！");
  window.location.href="login_worker.asp"
  </script>  		
	<%end if
end if
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
 "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>网上客户管理系统--业务员登录！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	background-image: url(Images/bg.gif);
}
.style1 {color: #FFFFFF}
.style2 {color: #a2bcc5}
.style4 {color: #527184}
-->
</style></head>

<body>
<script language="javascript">
function mycheck(){
if (form1.UserName.value=="")
{alert("请输入业务员姓名！");form1.UserName.focus();return;}
if(form1.PWD.value=="")
{alert("请输入密码！");form1.PWD.focus();return;}
form1.submit();
}
</script>
<table width="595" height="242" border="0" cellpadding="-2" cellspacing="-2"
 background="Images/login.gif">
  <tr>
    <td valign="top"><table width="595" height="399" cellpadding="-2" cellspacing="-2">
      <tr>
        <td width="112" height="140">&nbsp;</td>
        <td width="370">&nbsp;</td>
        <td width="111"><span class="style2"></span></td>
      </tr>
      <tr>
        <td height="211">&nbsp;</td>
        <td valign="top"><div align="center">
          <form name="form1" method="POST" action="">
            <table width="76%" height="154" border="1" cellpadding="-2"  cellspacing="-2"
			 bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#669999">
            <tr>
              <td height="19" bgcolor="#6B8E94"><div align="center" class="style1">
			  ==========&nbsp;[业务员登录]&nbsp;==========</div></td>
            </tr>
            <tr>
              <td height="93"><div align="center">
                <table width="63%" height="89" cellpadding="-2"  cellspacing="-2">
                  <tr>
				  <script language="javascript">
				  //响应回车事件
				  function EnterE(packName){
				  var keychar=event.keyCode;
				  if(keychar=="13")
				  {packName.focus();
				  }
				  }
				  </script>
                    <td width="48%"><div align="center"><span class="style1">业务员名称
					：</span></div></td>
                    <td width="52%"><div align="left">
                      <input name="UserName" type="text" class="Sytle_text" id="UserName" size="20"
					   onKeyPress="EnterE(document.form1.PWD)">
                    </div></td>
                  </tr>
                  <tr>
                    <td><div align="center"><span class="style1">业务员密码：</span></div></td>
                    <td><div align="left">
                      <input name="PWD" type="password" class="Sytle_text" id="PWD" size="20"
					   onKeyPress="EnterE(document.form1.Button)">
                    </div></td>
                  </tr>
                  <tr>
                    <td height="33" colspan="2"><div align="center">
                        <input name="Button" type="button" class="Style_button" value="登录" 
						onClick="mycheck()">
                        <input name="Submit2" type="reset" class="Style_button" value="重置">
                    </div></td>
                    </tr>
                </table>
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
