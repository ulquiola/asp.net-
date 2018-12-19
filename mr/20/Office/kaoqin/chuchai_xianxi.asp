<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../conn/conn.asp"-->
<%
if request.QueryString("ID")<>"" then
ID=request.QueryString("ID")
end if
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from Tab_chuchai where ID="&ID
rs.open sql,conn,1,3
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>出差登记详细信息</title>
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
        <td height="42" valign="bottom">&nbsp;&nbsp;&nbsp;&nbsp;<span class="STYLE2">出差登记详细信息</span></td>
      </tr>
      <tr>
        <td height="406"><table width="427" height="298" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="72" height="20"><div align="center" class="STYLE3">姓名</div></td>
            <td width="92"><input name="textfield" type="text" value="<%=rs("name1")%>" size="13">
              <div align="center" class="STYLE3"></div></td>
            <td width="87"><div align="center" class="STYLE3">所属部门</div></td>
            <td width="176"><div align="center" class="STYLE3">
              <div align="left">
                &nbsp;&nbsp;<input name="textfield" type="text" value="<%=rs("department")%>" size="13">
                </div>
            </div></td>
          </tr>
          <tr>
            <td height="187" colspan="4" valign="top" class="STYLE3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;出差地区/原因<strong>:</strong><br>
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("chuarea")%></td>
          </tr>
          <tr>
            <td height="15"><div align="center" class="STYLE3">开始时间</div></td>
            <td height="15"><div align="center" class="STYLE3"><input name="textfield" type="text" value="<%=rs("time1")%>" size="10"></div></td>
            <td height="15"><div align="center" class="STYLE3">终止时间</div></td>
            <td height="15"><div align="center" class="STYLE3"><input name="textfield" type="text" value="<%=rs("time2")%>" size="10"></div></td>
          </tr>
        </table><br>
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
