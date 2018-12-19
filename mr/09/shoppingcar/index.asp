<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="top.asp"--><!--网站包含文件-->
<!--#include file="conn/conn.asp"--><!--数据库连接文件-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>购物车</title>
<style type="text/css">
<!--
.STYLE1 {	font-size: 10pt;
	font-weight: bold;
	color: #000000;
}
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}

.STYLE2 {color: #000000}
-->
</style>

</head>

<body>
<table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="492" valign="top"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="183" colspan="2" valign="top"><table width="100%" height="191" border="0" cellpadding="0" cellspacing="0" background="images/new1.gif">
          <tr>
            <td height="191"><table width="800" height="188" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="226" height="188"><!--#include file="login.asp"--></td>
                <td width="574" valign="top">
                <table width="390" height="191" border="0" align="right" cellpadding="0" cellspacing="0">
                    <tr>
                      <td height="177" align="left" valign="bottom">
<%
set rs1=server.CreateObject("adodb.recordset")				'创建记录集
sql1="select top 4 * from tb_news order by time1 desc"		'查询数据
rs1.open sql1,conn,1,3										'打开记录集
If(rs1.Eof)Then												'判断是否有商品信息
	Response.Write("暂无商品信息!")							'输出提示信息
Else
While(Not rs1.Eof)											
%>
                          <table width="88%"  border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr>
                              <td height="31" class="STYLE2"><img src="Images/stars.gif" /> <span style="font-size: 9pt">
                                <%
If(Len(rs1("content")) > 18)Then						'判断信息内容的长度是否大于18
Response.Write("<a href='Detail1.asp?id="&rs1("Id")&"'>"&Server.HTMLEncode(Mid(rs1("content"),1,18))&"....</a>")
%>
                                &nbsp;<%=rs1("time1")%>
                                <%Else
Response.Write("<a href='Detail1.asp?id="&rs1("Id")&"'>"&Server.HTMLEncode(rs1("content"))&"</a>")
%>
                                &nbsp;<%=rs1("time1")%>
                                <%
End If
%>
                              </span><br><hr size="1" color="D7D7D7"></td>
                            </tr>
                            <%
rs1.MoveNext				'向下移动记录指针
Wend
rs1.close					'关闭记录集
Set rs1 = Nothing
%>
                            <tr>
                              <td height="20" align="right" valign="top"><a href="Default1.asp"><img src="images/more.gif" width="35" height="13" border="0" /></a></td>
                            </tr>
                          </table>
                        <%end if%>
                        &nbsp; </td>
                    </tr>
                </table></td>
              </tr>
            </table></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td colspan="2" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" background="images/center.gif">
          <tr>
            <td valign="top"><table width="763" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="28" height="231">&nbsp;</td>
                <td width="550" valign="bottom"><!--#include file="bestnew.asp"--></td>
                <td width="185" valign="middle"><table width="128" height="58" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td><table width="168" height="53" border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="150"><marquee direction="up" onMouseOver="this.stop()" onMouseOut="this.start()" scrollamount="3" height="126">
                          &nbsp;
                          <%
set rs=server.CreateObject("adodb.recordset")				'创建记录集
sql="select top 5 * from tb_bestnew order by time1 desc"	'查询数据
rs.open sql,conn,1,3										'打开记录集
If(rs.Eof)Then												'判断是否有动态信息
	Response.Write("暂无最新动态信息!")						
Else
While(Not rs.Eof)				
%>
                          <table width="90%"  border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr>
                              <td height="31" class="STYLE2">&nbsp;<img src="Images/new7.gif" width="3" height="3" /><span style="font-size: 9pt">&nbsp;
                                <%
If(Len(rs("title")) > 7)Then						'判断标题的长度是否大于7
Response.Write("<a href='#'>"&Server.HTMLEncode(Mid(rs("title"),1,7))&"....</a>")
Else
Response.Write("<a href='#'>"&Server.HTMLEncode(rs("title"))&"</a>")
End If
%>
                              </span></td>
                            </tr>
                            <%
rs.MoveNext					'向下移动记录指针
Wend
rs.close					'关闭记录集
Set rs = Nothing			'释放内存变量空间
%>
                            <tr>
                              <td height="20" align="right" valign="bottom"><a href="#"><img src="Images/more.gif" width="35" height="13" border="0" /></a></td>
                            </tr>
                          </table>
                          <%end if%>
                          </marquee>                        </td>
                      </tr>
                    </table></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td height="302">&nbsp;</td>
                <td height="302"><!--#include file="tejia.asp"--></td>
                <td><table width="166" height="115" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td><!--#include file="Ballot.asp"--></td>
                  </tr>
                </table></td>
              </tr>
            </table></td>
          </tr>
        </table></td>
        </tr>
</table></td>
  </tr>
</table><!--#include file="bottom.asp"-->
</body>
</html>
