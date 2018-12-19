<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="js/Function.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>个人基本资料</title>
<style type="text/css">
<!--
body {
	margin-top: 0px;
	margin-bottom: 0px;
	margin-left: 0px;
	margin-right: 0px;
}
.STYLE7 {
	font-size: 9pt;
	color: #000000;
}
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
}
a:hover {
	text-decoration: none;
}
a:active {
	text-decoration: none;
}
.STYLE8 {font-size: 9pt}
-->
</style>
</head>
<body>
<table width="436" height="251" border=0 align=center cellpadding=0 cellspacing=0>
  <tr>
    <td width="436" height="251" valign=top background="Images/yonghu.gif"><p>&nbsp;</p>
      <table width=78% border=0 align=left cellpadding=0 cellspacing=0>
      <tr>
        <td colspan=3></td>
      </tr>
      <tr>
        <td colspan=3><!-- #include file="Conn/conn.asp" -->
 <% 
	if request.QueryString("id")<>"" then
	ID=request.QueryString("id") 
  	end if
	set rs_1=server.CreateObject("adodb.recordset")
	sql1="select * from tb_user where ID="&id
  	rs_1.open sql1,conn,1,2
%>
              <table width="311" height="160" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="107" height="25"><div align="center" class="STYLE7">用户ID：</div></td>
                  <td width="99" class="STYLE7"><%=rs_1("id")%>&nbsp;</td>
                  <td width="90" rowspan="8" valign="top"><div align="center"><img src="Images\head\<%=rs_1("touxiang")%>.gif" width="80" height="80" border="1">&nbsp;</div></td>
                </tr>
                <tr>
                  <td height="26"><div align="center" class="STYLE7">用户名称：</div></td>
                  <td height="26" class="STYLE7"><%=rs_1("username")%>&nbsp;</td>
                </tr>
                <tr>
                  <td height="23"><div align="center" class="STYLE7">性&nbsp;&nbsp;&nbsp;别：</div></td>
                  <td height="23" class="STYLE7"><%=rs_1("sex")%>&nbsp;</td>
                </tr>
                <tr>
                  <td height="20"><div align="center" class="STYLE7">用户昵称：</div></td>
                  <td height="20" class="STYLE7"><%=rs_1("state")%>&nbsp;</td>
                </tr>
                <tr>
                  <td height="32"><div align="center" class="STYLE7">生&nbsp;&nbsp;&nbsp;日：</div></td>
                  <td height="32" class="STYLE7"><%=rs_1("birthday")%>&nbsp;</td>
                </tr>
                <tr>
                  <td height="18"><div align="center" class="STYLE7">E-mail：</div></td>
                  <td height="18" class="STYLE7"><%=rs_1("email")%>&nbsp;</td>
                </tr>
            </table>
          <% 
	set rs_1=nothing
	conn.close
	set conn=nothing
%>        </td>
      </tr>
    <td width="62%" height="25" valign="bottom"><div align="center"><a href="#" class="STYLE8" onClick="javascript:window.close()">【关闭窗口】</a> </div>
    </table></td>
  </tr>
</table>
</body>
</html>
