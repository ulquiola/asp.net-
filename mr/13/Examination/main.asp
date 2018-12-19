<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp"--><!--包含数据库连接文件-->
<!--#include file="inc/Validate.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<link rel="stylesheet" type="text/css" href="Css/style.css">
<style type="text/css">
<!--
.STYLE1 {color: #000000}
-->
</style>
</head>
<body topmargin="0" leftmargin="0"> 
  <table width="770" height="573" border="0" align="center" cellpadding="0" cellspacing="0">
<form action="SelLesson.asp" method="post"> 
  <tr>
    <td height="21">&nbsp;</td>
  </tr>
  <tr>
    <td>
	<table cellpadding="0" cellspacing="0" border="0" width="100%" height="100%"> 
          <tr> 
            <td width="26"> </td> 
            <td  height="280"> <table cellpadding="0" cellspacing="0" border="0" width="100%" height="100%"> 
                <tr> 
                  <td width="70%" height="24" align="right"><a href="board.asp" target="mainFrame" class="STYLE1">预约专家</a>&nbsp;|<a href="QueryRel.asp" target="mainFrame" class="STYLE1">查询成绩</a>&nbsp;|&nbsp;<a href="Logout.asp" class="STYLE1">退出系统</a></td> 
				  	
					<%set rs=server.CreateObject("adodb.recordset")
	Sql = "select * from tb_count"
	rs.Open Sql,conn,1,3
					%>
                  <td width="30%" align="center">&nbsp;</td>
                </tr> 
                <tr> 
                  <td  height="320" colspan="2" align="center"> <table cellpadding="0" cellspacing="0" border="0" width="676" height="280"> 
                      <tr align="right"> 
                        <td height="10" colspan="2"></td> 
                      </tr> 
                      <tr> 
                        <td height="280" colspan="2"> <table cellpadding="0" cellspacing="0" border="0" width="100%" height="85%"> 
                            <tr> 
                              <td width="1" height="433"> </td> 
                              <td width="674" height="510" valign="top"><iframe src="QueryRel.asp" frameborder="0" width="100%" height="100%" name="mainFrame" scrolling="auto">对不起！你的浏览器不支持框架</iframe></td> 
                              <td width="1" height="450"> </td> 
                            </tr> 
                        </table></td> 
                      </tr> 
                      <tr align="right"> 
                        <td width="390" height="13">&nbsp;</td> 
                        <td width="286" align="center"><font color="#FF0000">您的咨询卡截止日期为：<%=rs("stoptime")%></font></td>
                      </tr> 
                  </table></td> 
                </tr> 
              </table></td> 
            <td width="27"> </td> 
          </tr> 
        </table>	</td>
  </tr>
  <tr>
    <td height="18">&nbsp;</td>
    </tr>
  </form> 
</table>
</body>
</html>
