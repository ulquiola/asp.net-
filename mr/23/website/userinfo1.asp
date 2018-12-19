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
MR_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MR_editAction = MR_editAction & "?" & Request.QueryString
End If
MR_abortEdit = false
MR_editQuery = ""
%>
<%
If (CStr(Request("MR_update")) <> "" And CStr(Request("MR_recordId")) <> "") Then
  MR_editConnection = MR_website_STRING
  MR_editTable = "website"
  MR_editColumn = "idd"
  MR_recordId = "" + Request.Form("MR_recordId") + ""
  MR_editRedirectUrl = "userinfo.asp?diy=ok"
  MR_fieldsStr  = "name|value|title|value|address|value|phone|value|fax|value|mobile|value|email|value"
  MR_columnsStr = "name|',none,''|title|',none,''|address|',none,''|phone|',none,''|fax|',none,''|mobile|',none,''|email|',none,''"
  MR_fields = Split(MR_fieldsStr, "|")
  MR_columns = Split(MR_columnsStr, "|")
  
  For i = LBound(MR_fields) To UBound(MR_fields) Step 2
    MR_fields(i+1) = CStr(Request.Form(MR_fields(i)))
  Next

  If (MR_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MR_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MR_editRedirectUrl = MR_editRedirectUrl & "?" & Request.QueryString
    Else
      MR_editRedirectUrl = MR_editRedirectUrl & "&" & Request.QueryString
    End If
  End If
End If
%>
<%
If (CStr(Request("MR_update")) <> "" And CStr(Request("MR_recordId")) <> "") Then
  MR_editQuery = "update " & MR_editTable & " set "
  For i = LBound(MR_fields) To UBound(MR_fields) Step 2
    FormVal = MR_fields(i+1)
    MR_typeArray = Split(MR_columns(i+1),",")
    Delim = MR_typeArray(0)
    If (Delim = "none") Then Delim = ""
    AltVal = MR_typeArray(1)
    If (AltVal = "none") Then AltVal = ""
    EmptyVal = MR_typeArray(2)
    If (EmptyVal = "none") Then EmptyVal = ""
    If (FormVal = "") Then
      FormVal = EmptyVal
    Else
      If (AltVal <> "") Then
        FormVal = AltVal
      ElseIf (Delim = "'") Then
        FormVal = "'" & Replace(FormVal,"'","''") & "'"
      Else
        FormVal = Delim + FormVal + Delim
      End If
    End If
    If (i <> LBound(MR_fields)) Then
      MR_editQuery = MR_editQuery & ","
    End If
    MR_editQuery = MR_editQuery & MR_columns(i) & " = " & FormVal
  Next
  MR_editQuery = MR_editQuery & " where " & MR_editColumn & " = " & MR_recordId

  If (Not MR_abortEdit) Then
    ' 更新数据
    Set MR_editCmd = Server.CreateObject("ADODB.CoMMand")
    MR_editCmd.ActiveConnection = MR_editConnection
    MR_editCmd.CoMMandText = MR_editQuery
    MR_editCmd.Execute
    MR_editCmd.ActiveConnection.Close

    If (MR_editRedirectUrl <> "") Then
      Response.Redirect(MR_editRedirectUrl)
    End If
  End If

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
                <td height="24"> 您现在位置：首页-&gt;注册资料-&gt;修改注册资料</td>
              </tr>
              <tr>
                <td height="24"><script language="JavaScript">
<!--
function Juge(theForm)
{
  if (theForm.name.value == "")
  {
    alert("请输入您的姓名!");
    theForm.name.focus();
    return (false);
  }
  if (checktext(theForm.name.value))
  {
    alert("请您输入您的姓名!");
    theForm.name.select();
    theForm.name.focus();
    return (false);
  }
  if (theForm.title.value == "")
  {
    alert("请输入网站名称!");
    theForm.title.focus();
    return (false);
  }
  if (checktext(theForm.title.value))
  {
    alert("请输入网站名称!");
    theForm.title.select();
    theForm.title.focus();
    return (false);
  }
    if (theForm.address.value == "")
  {
    alert("请输入公司地址!");
    theForm.address.focus();
    return (false);
  }
  if (checktext(theForm.address.value))
  {
    alert("请输入公司地址!");
    theForm.address.select();
    theForm.address.focus();
    return (false);
  }
  if (theForm.phone.value == "")
  {
    alert("请输入电话号码!");
    theForm.phone.focus();
    return (false);
  }
  if (checktext(theForm.phone.value))
  {
    alert("请输入电话号码!");
    theForm.phone.select();
    theForm.phone.focus();
    return (false);
  }
 if (theForm.email.value == "")
 {
    alert("请填写您的电子信箱！");
    theForm.email.focus();
    return (false);
  }  
  if (theForm.email.value.indexOf("@")<0 || theForm.email.value.indexOf(".")<0 || 0<=theForm.email.value.indexOf("@.") || 0<=theForm.email.value.indexOf(".@"))
  {
    alert("您的电子信箱填写有误！");
    theForm.email.focus();
    return (false);
  }

}
function checktext(text)
{
	allValid = true;

	for (i = 0;  i < text.length;  i++)
		{
			if (text.charAt(i) != " ")
			{
				allValid = false;
				break;
			}
		}
return allValid;
}
//-->
</script>
                </td>
              </tr>
              <tr>
                <td><form method="POST" action="<%=MR_editAction%>" name="" onSubmit="return Juge(this)">
                    <table width="80%" border="0" cellspacing="2" cellpadding="1" align="center" bgcolor="#A2DBEE">
                      <tr>
                        <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="#A2DBEE">
                            <tr>
                              <td bgcolor="#FFFFFF" width="26%"><div align="right">联系人：</div></td>
                              <td bgcolor="#FFFFFF"><input type="text" name="name" value="<%=(Rs.Fields.Item("name").Value)%>" size="32" maxlength="20" style="BORDER-RIGHT: rgb(0,0,0) 1px solid; BORDER-TOP: rgb(0,0,0) 1px solid; BORDER-LEFT: rgb(0,0,0) 1px solid; BORDER-BOTTOM: rgb(0,0,0) 1px solid; HEIGHT: 18px" class="pt9">
                                <font color="#FF0000"> *</font> </td>
                            </tr>
                            <tr>
                              <td bgcolor="#FFFFFF" width="26%"><div align="right">网站名称：</div></td>
                              <td bgcolor="#FFFFFF"><input type="text" name="title" value="<%=(Rs.Fields.Item("title").Value)%>" size="32" maxlength="50" style="BORDER-RIGHT: rgb(0,0,0) 1px solid; BORDER-TOP: rgb(0,0,0) 1px solid; BORDER-LEFT: rgb(0,0,0) 1px solid; BORDER-BOTTOM: rgb(0,0,0) 1px solid; HEIGHT: 18px" class="pt9">
                                <font color="#FF0000"> *</font> </td>
                            </tr>
                            <tr>
                              <td bgcolor="#FFFFFF" width="26%"><div align="right">用户地址：</div></td>
                              <td bgcolor="#FFFFFF"><input type="text" name="address" value="<%=(Rs.Fields.Item("address").Value)%>" size="32" maxlength="50" style="BORDER-RIGHT: rgb(0,0,0) 1px solid; BORDER-TOP: rgb(0,0,0) 1px solid; BORDER-LEFT: rgb(0,0,0) 1px solid; BORDER-BOTTOM: rgb(0,0,0) 1px solid; HEIGHT: 18px" class="pt9">
                                <font color="#FF0000"> *</font> </td>
                            </tr>
                            <tr>
                              <td bgcolor="#FFFFFF" width="26%"><div align="right">电话号码：</div></td>
                              <td bgcolor="#FFFFFF"><input type="text" name="phone" value="<%=(Rs.Fields.Item("phone").Value)%>" size="32" maxlength="30" style="BORDER-RIGHT: rgb(0,0,0) 1px solid; BORDER-TOP: rgb(0,0,0) 1px solid; BORDER-LEFT: rgb(0,0,0) 1px solid; BORDER-BOTTOM: rgb(0,0,0) 1px solid; HEIGHT: 18px" class="pt9">
                                <font color="#FF0000"> *</font> </td>
                            </tr>
                            <tr>
                              <td bgcolor="#FFFFFF" width="26%"><div align="right">传真号码：</div></td>
                              <td bgcolor="#FFFFFF"><input type="text" name="fax" value="<%=(Rs.Fields.Item("fax").Value)%>" size="32" maxlength="30" style="BORDER-RIGHT: rgb(0,0,0) 1px solid; BORDER-TOP: rgb(0,0,0) 1px solid; BORDER-LEFT: rgb(0,0,0) 1px solid; BORDER-BOTTOM: rgb(0,0,0) 1px solid; HEIGHT: 18px" class="pt9">
                              </td>
                            </tr>
                            <tr>
                              <td bgcolor="#FFFFFF" width="26%"><div align="right">手机号码：</div></td>
                              <td bgcolor="#FFFFFF"><input type="text" name="mobile" value="<%=(Rs.Fields.Item("mobile").Value)%>" size="32" maxlength="30" style="BORDER-RIGHT: rgb(0,0,0) 1px solid; BORDER-TOP: rgb(0,0,0) 1px solid; BORDER-LEFT: rgb(0,0,0) 1px solid; BORDER-BOTTOM: rgb(0,0,0) 1px solid; HEIGHT: 18px" class="pt9">
                              </td>
                            </tr>
                            <tr>
                              <td bgcolor="#FFFFFF" width="26%"><div align="right">电子信箱：</div></td>
                              <td bgcolor="#FFFFFF"><input type="text" name="email" value="<%=(Rs.Fields.Item("email").Value)%>" size="32" maxlength="50" style="BORDER-RIGHT: rgb(0,0,0) 1px solid; BORDER-TOP: rgb(0,0,0) 1px solid; BORDER-LEFT: rgb(0,0,0) 1px solid; BORDER-BOTTOM: rgb(0,0,0) 1px solid; HEIGHT: 18px" class="pt9">
                                <font color="#FF0000"> *</font> </td>
                            </tr>
                            <tr>
                              <td bgcolor="#FFFFFF" width="26%"><div align="right"></div></td>
                              <td bgcolor="#FFFFFF"><input type="submit" value="修 改 资 料" name="Submit" 
								 cursor:hand">
                                <input type="reset" value="恢 复 原 状" name="Reset" 
								 cursor:hand">
                              </td>
                            </tr>
                          </table></td>
                      </tr>
                    </table>
                    <input type="hidden" name="MR_update" value="true">
                    <input type="hidden" name="MR_recordId" value="<%= Rs.Fields.Item("idd").Value %>">
                  </form></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
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
