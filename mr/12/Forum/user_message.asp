<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��̴ʵ似����̳</title>
<link rel="stylesheet" type="text/css" href="css/style.css">
<style type="text/css">
<!--
.STYLE6 {font-size: 12px}
.STYLE7 {font-size: 9px}
.STYLE10 {font-size: 9pt; color: #FF0000; }
-->
</style>
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
		<font color="#FFFFFF">��<a href="login.asp" target="mainwindow" class="a3">�û���¼</a>����<a href="register.asp" target="mainwindow" class="a3">ע��</a>��		</font>
	 <%Else%>
		<table width="100%" height="22" border="0" align="center" cellpadding="0" cellspacing="0">
		<tr>
		<td width="19%" valign="top"><div align="center"><img src="images/head/<%=rs_1("ICO")%>" width="20" height="20"></div></td>
		<td width="81%"> <font color="#FFFFFF">�û���<%=Session("UserName")%>&nbsp;&nbsp;��<a href="logout.asp" class="a3">��ȫ�˳�</a>��</font></td>
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
    <td width="28" height="40" style="cursor:hand" title="��̳���" background="Images/bk2.gif"><div align="center"><a href="left.asp"><img src="images/b1.gif" width="27" height="20" border="0"></a></div></td>
  </tr>
  <tr>
    <td width="28" height="1"></td>
  </tr>
  <tr>
    <td width="28" height="40" style="cursor:hand" title="�û���Ϣ" background="Images/bk1.gif"><div align="center"><a href="user_message.asp"><img src="images/b2.gif" width="24" height="23" border="0"></a></div></td>
  </tr>
  <tr>
    <td width="28" height="1"></td>
  </tr>
  <tr>
    <td width="28" height="40" style="cursor:hand" title="����֧��" background="Images/bk2.gif"><div align="center"><a href="skill_up.asp"><img src="images/b3.gif" width="27" height="27" border="0"></a></div></td>
  </tr>
  <tr>
    <td width="28" height="1"></td>
  </tr>
  <tr>
    <td width="28" height="40" style="cursor:hand" title="ʹ�ð���" background="Images/bk2.gif"><div align="right"><a href="help.asp"><img src="images/b4.gif" width="23" height="24" border="0"></a></div></td>
  </tr>
</table>    </td>
    <td width="85%" valign="top" bgcolor="#FFFFFF"><table width="96%" height="2" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td></td>
      </tr>
    </table>
      <table width="96%" height="17" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td background="images/bbs_top.gif"><div align="center">��̴ʵ似����̳</div></td>
        </tr>
      </table>
      <br>
     <%If Session("UserName")="" or Session("UserName")="Friend"  Then%>
	  <table width="85%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td style="line-height:2"><div align="center">����<a href="login.asp" class="a2" target="mainwindow"><font color="#FF0000">��¼</font></a>��Ȼ��鿴������Ϣ��</div></td>
        </tr>
      </table>
	<%Else%>
	 <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td><div align="center"><img src="images/head/<%=rs_1("ICO")%>"><br><br><%=Session("UserName")%><br><br><br></div></td>
            </tr>
      </table>
            <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td><div align="center"><img src="images/email.gif" width="45" height="16" title="<%=rs_1("Email")%>" style="cursor:hand"/><img src="images/qq.gif" width="16" height="16" title="<%=rs_1("QQ")%>" style="cursor:hand"/><img src="images/ip.gif" width="13" height="16" title="<%=rs_1("Tel")%>" style="cursor:hand"/><span class="STYLE6">�绰</span></div></td>
              </tr>
            </table>
            <br>
            <table width="85%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td style="padding:5; line-height:2"><p>ע��ʱ�䣺<br>
                    <span class="STYLE7"><span class="STYLE10">&nbsp;&nbsp;<%=rs_1("re_date")%></span></span><br>
                  ����¼ʱ�䣺<br>
                  <span class="STYLE10"><%=rs_1("lastlogintime")%></span> &nbsp;&nbsp;&nbsp;<br>
                  ��¼������<br>
                    &nbsp;<span class="STYLE10"><%=rs_1("logintimes")%></font>&nbsp;</span>��<br>
                  <br>
                ��<a href="Modify.asp" target="mainwindow" class="a1">�޸ĸ�����Ϣ</a>��<br>
                ��<a href="logout.asp" class="a1">�˳���¼</a>��</p></td>
              </tr>
            </table>
	<%End If%>
	  </td>
  </tr>
</table>
</body>
</html>