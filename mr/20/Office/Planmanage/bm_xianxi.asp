<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../conn/conn.asp"-->
<%
if request.QueryString("ID")<>"" then
ID=request.QueryString("ID")
end if
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from Tab_bm where ID="&ID
rs.open sql,conn,1,3
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>部门计划详细信息</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}
.STYLE2 {
	font-size: 10pt;
	color: #FFFFFF;
}
.STYLE3 {
	font-size: 9pt;
	color: #000000;
}
-->
</style></head>

<body>
<table width="458" height="459" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top" background="../Images/waichudeng.gif"><table width="442" height="448" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="42" valign="bottom">&nbsp;&nbsp;&nbsp;&nbsp;<span class="STYLE2">部门计划详细信息</span></td>
      </tr>
      <tr>
        <td height="406"><table width="427" height="298" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="72" height="20"><div align="center" class="STYLE3">部门名称</div></td>
            <td width="136"><input name="textfield" type="text" value="<%=rs("name1")%>" size="20">
              <div align="center" class="STYLE3"></div></td>
            <td width="57"><div align="center" class="STYLE3">计划主题</div></td>
            <td width="162"><div align="center" class="STYLE3">
              <div align="left">
                <input name="textfield" type="text" value="<%=rs("title")%>" size="20">
                </div>
            </div></td>
          </tr>
          <tr>
            <td height="187" valign="middle" class="STYLE3">&nbsp;&nbsp;
              <div align="center">计划内容<br>
              </div></td>
            <td height="187" colspan="3" valign="middle" class="STYLE3"><textarea name="content" cols="48" rows="17" id="content"><%=rs("content")%></textarea></td>
            </tr>
          <tr>
            <td height="15"><div align="center" class="STYLE3">发表时间</div></td>
            <td height="15"><div align="center" class="STYLE3">
              <div align="left">
                <input name="textfield" type="text" value="<%=rs("time1")%>" size="10">
              </div>
            </div></td>
            <td height="15" colspan="2">&nbsp;</td>
            </tr>
        </table>
        <br>
        <table width="131" height="39" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td align="center"><input type="submit" name="Submit" value=" 关闭窗口 " onclick="javascrip:window.close()"></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>