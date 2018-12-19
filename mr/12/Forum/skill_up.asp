<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>编程词典技术论坛</title>
<link rel="stylesheet" type="text/css" href="css/style.css">
</head>
<script src="JS/fun.js"></script>
<% 
	set rs_1=server.CreateObject("adodb.recordset")
	sql1="select * from tb_User where username='"&session("username")&"'"
  	rs_1.open sql1,conn,1,3
%>
<body topmargin="0" leftmargin="0" bottommargin="0" style="overflow:auto" class="scrollbar">
<table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="24" colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" height="24" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td background="images/back_1.gif">
		<div align="center">
		<%If Session("UserName")="" or Session("UserName")="Friend"  Then%>
		<font color="#FFFFFF">【<a href="login.asp" target="mainwindow" class="a3">用户登录</a>】【<a href="register.asp" target="mainwindow" class="a3">注册</a>】
		</font>
	 <%Else%>
		<table width="100%" height="22" border="0" align="center" cellpadding="0" cellspacing="0">
		<tr>
		<td width="19%" valign="top"><div align="center"><img src="images/head/<%=rs_1("ICO")%>" width="20" height="20"></div></td>
		<td width="81%"> <font color="#FFFFFF">用户：<%=Session("UserName")%>&nbsp;&nbsp;【<a href="logout.asp" class="a3">安全退出</a>】</font></td>
		</tr>
		</table>
		<%End If%>
		</font></div></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td width="15%" valign="top" bgcolor="#B9DFFF"><table width="96%" height="1" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td></td>
        </tr>
      </table>
      <table border="0" align="right" cellpadding="0" cellspacing="0">
  <tr>
    <td width="28" height="40" style="cursor:hand" title="论坛版块" background="Images/bk2.gif"><div align="center"><a href="left.asp"><img src="images/b1.gif" width="27" height="20" border="0"></div></td>
  </tr>
  <tr>
    <td width="28" height="1"></td>
  </tr>
  <tr>
    <td width="28" height="40" style="cursor:hand" title="用户信息" background="Images/bk2.gif"><div align="center"><a href="user_message.asp"><img src="images/b2.gif" width="24" height="23" border="0"></a></div></td>
  </tr>
  <tr>
    <td width="28" height="1"></td>
  </tr>
  <tr>
    <td width="28" height="40" style="cursor:hand" title="技术支持" background="Images/bk1.gif"><div align="center"><a href="skill_up.asp"><img src="images/b3.gif" width="27" height="27" border="0"></a></div></td>
  </tr>
  <tr>
    <td width="28" height="1"></td>
  </tr>
  <tr>
    <td width="28" height="40" style="cursor:hand" title="使用帮助" background="Images/bk2.gif"><div align="right"><a href="help.asp"><img src="images/b4.gif" width="23" height="24" border="0"></a></div></td>
  </tr>
</table>    </td>
    <td width="85%" valign="top" bgcolor="#FFFFFF"><table width="96%" height="2" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td></td>
      </tr>
    </table>
      <table width="88%" border="0" align="center" cellpadding="0" cellspacing="0" style="display:none" id="tb_2_2">
       
	    <tr>
          <td width="16" height="16">&nbsp;</td>
          <td width="16" height="16"><img src="images/tshaped.gif" width="16" height="16"></td>
          <td width="16" height="16"><img src="images/11.gif" width="16" height="16"></td>
          <td>&nbsp;<a href="" target="mainwindow" ></a></td>
        </tr>
		  
	     <tr>
          <td width="16" height="16">&nbsp;</td>
          <td width="16" height="16"><img src="images/upangle.gif" width="16" height="16"></td>
          <td width="16" height="16"><img src="images/11.gif" width="16" height="16"></td>
          <td>&nbsp;<a href="" ></a></td>
        </tr>
		 
      </table>
	
	

      <table width="96%" height="17" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td background="images/bbs_top.gif"><div align="center">编程词典技术论坛</div></td>
        </tr>
      </table>
      <br>
      <table width="85%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="13%" height="30" valign="top"><img src="images/mark.gif" width="10" height="18"></td>
    <td width="87%" style=" line-height:2">服务热线：<br>
      <font color="#000000"></font></td>
  </tr>
  <tr>
    <td height="30" valign="top"><img src="images/mark.gif" width="10" height="18"></td>
    <td height="30" style=" line-height:2">QQ在线咨询：<br><a href=""><font color="#000000"></font></a></td>
  </tr>
  <tr>
    <td height="30" valign="top"><img src="images/mark.gif" width="10" height="18"></td>
    <td height="30" style=" line-height:2">企业邮箱：<br><font color="#000000"></font></td>
  </tr>
  <tr>
    <td height="30" valign="top"><img src="images/mark.gif" width="10" height="18"></td>
    <td height="30" style=" line-height:2">公司地址：<br><font color="#000000"></font></td>
  </tr>
</table>

	

	<br></td>
  </tr>
</table>
</body>
</html>