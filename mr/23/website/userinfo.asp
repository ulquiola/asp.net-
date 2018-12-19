<%@LANGUAGE="VBSCRIPT"%>
<% Session.Timeout=200 %>
<!--#include file="Conn/conn.asp" -->
<%
MR_authorizedUsers=""
MR_authFailedURL="timeout.asp"
MR_grantAccess=false
If Session("MR_Username") <> "" Then
  If (true Or CStr(Session("MR_UserAuthorization"))="") Or _
         (InStr(1,MR_authorizedUsers,Session("MR_UserAuthorization"))>=1) Then
    MR_grantAccess = true
  End If
End If
If Not MR_grantAccess Then
  MR_qsChar = "?"
  If (InStr(1,MR_authFailedURL,"?") >= 1) Then MR_qsChar = "&"
  MR_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MR_referrer = MR_referrer & "?" & Request.QueryString()
  MR_authFailedURL = MR_authFailedURL & MR_qsChar & "accessdenied=" & Server.URLEncode(MR_referrer)
  Response.Redirect(MR_authFailedURL)
End If
%>
<%
Dim Rs__MRColParam
Rs__MRColParam = "1"
if (Session("MR_username") <> "") then Rs__MRColParam = Session("MR_username")
%>
<%
set Rs = Server.CreateObject("ADODB.Recordset")
Rs.ActiveConnection = MR_website_STRING
Rs.Source = "SELECT * FROM website WHERE id = '" + Replace(Rs__MRColParam, "'", "''") + "'"
Rs.CursorType = 0
Rs.CursorLocation = 2
Rs.LockType = 3
Rs.Open()
Rs_numRows = 0
%>
<html>
<head>
<title>建站平台</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link 
href="css/top.css" rel=stylesheet type="text/css">
<style type="text/css">
<!--
.pt9 {  font-size: 9pt; font-family: "宋体"}
.pt105 {  font-size: 10.5pt; font-family: "宋体"}
.pt9n {  font-size: 9pt; line-height: 13pt; font-family: "宋体"}
.pt9nx {  font-size: 9pt; line-height: 15pt; font-family: "宋体"}
-->
</style>
<script>

var Temp;
var TimerId = null;
var TimerRunning = false;

Seconds = 0
Minutes = 0
Hours = 0

function showtime() 
{
if(Seconds >= 59) 
{
Seconds = 0
if(Minutes >= 59) 
{
Minutes = 0
if(Hours >= 23) 
{
Seconds = 0
Minutes = 0
Hours = 0
} 
else { 
++Hours 
}
} 
else { 
++Minutes 
}
} 
else { 
++Seconds 
}

if(Seconds != 1) { var ss="s" } else { var ss="" }
if(Minutes != 1) { var ms="s" } else { var ms="" }
if(Hours != 1) { var hs="s" } else { var hs="" }

Temp = '您在本页停留了 '+Hours+' 小时'+', '+Minutes+' 分'+', '+Seconds+' 秒'+''
window.status = Temp;
TimerId = setTimeout("showtime()", 1000);
TimerRunning = true;
}

var TimerId = null;
var TimerRunning = false;

function stopClock() {
if(TimerRunning) 
clearTimeout(TimerId);
TimerRunning = false;
}

function startClock() {
stopClock();
showtime();
}

function stat(txt) {
window.status = txt;
setTimeout("erase()", 2000);
}

function erase() {
window.status = "";
}

</SCRIPT>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" onLoad="startClock()">
<!-- #include file="top.asp"-->
<table width="776" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#679cdc">
  <tr>
    <td bgcolor="#FFFFFF"><table width="774" border="0" cellspacing="0" cellpadding="0" height="300">
        <tr valign="top">
          <td width="160" bgcolor="#A2DBEE" height="407"><!-- #include file="left.asp"-->
          </td>
          <td width="10" height="407"></td>
          <td width="604" height="407"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="24"> 您现在位置：首页-&gt;注册资料</td>
              </tr>
              <tr>
                <td height="24"><div align="center"><b><font color="#FF0000" class="pt105">
                    <% If (Request("diy")) <> (ok) Then %>
                    </font></b><font color="#FF0000" class="pt105"> <font color="#000000">注册资料更改成功!</font></font><b><font color="#FF0000" class="pt105">
                    <% End If %>
                    </font></b></div></td>
              </tr>
              <tr>
                <td><form>
                    <table width="80%" border="0" cellspacing="2" cellpadding="1" align="center" bgcolor="#A2DBEE">
                      <tr>
                        <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="#A2DBEE">
                            <tr>
                              <td bgcolor="#FFFFFF" width="26%" height="28"><div align="right">联系人：</div></td>
                              <td bgcolor="#FFFFFF" height="28"><%=(Rs.Fields.Item("name").Value)%> </td>
                            </tr>
                            <tr>
                              <td bgcolor="#FFFFFF" width="26%" height="28"><div align="right">网站名称：</div></td>
                              <td bgcolor="#FFFFFF" height="28"><%=(Rs.Fields.Item("title").Value)%> </td>
                            </tr>
                            <tr>
                              <td bgcolor="#FFFFFF" width="26%" height="28"><div align="right">用户地址：</div></td>
                              <td bgcolor="#FFFFFF" height="28"><%=(Rs.Fields.Item("address").Value)%> </td>
                            </tr>
                            <tr>
                              <td bgcolor="#FFFFFF" width="26%" height="28"><div align="right">电话号码：</div></td>
                              <td bgcolor="#FFFFFF" height="28"><%=(Rs.Fields.Item("phone").Value)%> </td>
                            </tr>
                            <tr>
                              <td bgcolor="#FFFFFF" width="26%" height="28"><div align="right">传真号码：</div></td>
                              <td bgcolor="#FFFFFF" height="28"><%=(Rs.Fields.Item("fax").Value)%> </td>
                            </tr>
                            <tr>
                              <td bgcolor="#FFFFFF" width="26%" height="28"><div align="right">手机号码：</div></td>
                              <td bgcolor="#FFFFFF" height="28"><%=(Rs.Fields.Item("mobile").Value)%> </td>
                            </tr>
                            <tr>
                              <td bgcolor="#FFFFFF" width="26%" height="28"><div align="right">电子信箱：</div></td>
                              <td bgcolor="#FFFFFF" height="28"><%=(Rs.Fields.Item("email").Value)%> </td>
                            </tr>
                            <tr>
                              <td bgcolor="#FFFFFF" width="26%" height="28"><div align="right"></div></td>
                              <td bgcolor="#FFFFFF" height="28"><input type="button" name="Button" onClick="location.href='userinfo1.asp'" value="修 改 资 料" class="pt9n" 
								 cursor:hand>
                              </td>
                            </tr>
                          </table></td>
                      </tr>
                    </table>
                  </form></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>
<!-- #include file="bottom.asp"-->
</body>
</html>
<%
Rs.Close()
%>
