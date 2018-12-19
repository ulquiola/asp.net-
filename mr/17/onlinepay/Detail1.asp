<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="top.asp"-->
<!--#include file="Conn/conn.asp"--><!--数据库连接文件-->
<%
if request.QueryString("ID")<>"" then'判断获取的ID值是否空
	ID=request.QueryString("ID")'为ID变量赋值
end if
Set rs=Server.CreateObject("ADODB.Recordset")'创建记录集
sql="select * from tb_news where ID="&ID'查询数据
rs.open sql,conn,1,3'打开记录集
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
</head>
<style type="text/css">
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}
</style>
<body>
<table width="813" height="58" border="0" align="center" cellpadding="0" cellspacing="0">
      <table width="800" height="526" border="0" align="center" cellpadding="0" cellspacing="0" background="images/new15.gif">
        <tr>
          <td height="86" colspan="3">&nbsp;</td>
        </tr>
        <tr>
          <td width="46" height="289">&nbsp;</td>
          <td width="713" valign="top"><table width="693" height="70" border="0" align="center" cellpadding="0" cellspacing="0">

              <tr>
                <td width="693" height="70"><table width="592" align="center"  cellpadding="0" cellspacing="0" style="border:#F7AB58 1px solid;">
                  <tr>
                    <td width="590" height="24" align="center" nowrap="nowrap"><%=Server.HTMLEncode(rs("title"))%> </td>
                  </tr>
                  <tr>
                    <td height="24" align="right" nowrap="nowrap"> 发表时间：<%=rs("time1")%> <br />
                        <hr size="1" color="#F7AB58" />                    </td>
                  </tr>
                  <tr>
                    <td align="left" nowrap="nowrap"><br />
                        <%=replace(rs("Content"),chr(13),"<br>")'输出新闻内容%> </td>
                  </tr>
                  <tr>
                    <td height="13">&nbsp;</td>
                  </tr>
                </table>
                  <%
rs.close
Set rs = Nothing
%>
&nbsp;</td>
              </tr>
          </table></td>
          <td width="41">&nbsp;</td>
        </tr>
        <tr>
          <td height="13" colspan="3"><table width="100%" height="21" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="40" height="15">&nbsp;</td>
                <td width="19">&nbsp;</td>
                <td width="665">&nbsp;</td>
                <td width="40">&nbsp;</td>
                <td width="34">&nbsp;</td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td height="130" colspan="3">&nbsp;</td>
        </tr>
</table>
      <tr>
        <td width="813">
<map name="Map" id="Map"><area shape="rect" coords="647,30,707,55" href="default1.asp" />
</map></body>
</html>
