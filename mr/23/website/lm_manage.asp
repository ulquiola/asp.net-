<%@LANGUAGE="VBSCRIPT"%>
<% Session.Timeout=200 %>
<!--#include file="conn/conn.asp" -->
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
  MR_editRedirectUrl = "lm_manage.asp?diy=ok"
  MR_fieldsStr  = "cb01|value|tf01|value|cbxw|value|tf03|value|cb04|value|tf04|value|cb05|value|tf05|value|cb06|value|tf06|value|cb07|value|tf07|value"
  MR_columnsStr = "yin_1|none,'Y','N'|title1|',none,''|yin_news|none,'Y','N'|title_news|',none,''|yin_3|none,'Y','N'|title3|',none,''|yin_4|none,'Y','N'|title4|',none,''|yin_5|none,'Y','N'|title5|',none,''|yin_6|none,'Y','N'|title6|',none,''|yin_7|none,'Y','N'|title7|',none,''"
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
<title>自动建站</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link 
href="css/top.css" rel=stylesheet type="text/css">
<link rel="stylesheet" type="text/css" href="css/forms.css">
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
<!-- #include file="top.asp" -->
<table width="776" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#679cdc">
  <tr>
    <td bgcolor="#FFFFFF"><table width="774" border="0" cellspacing="0" cellpadding="0" height="300">
        <tr valign="top">
          <td width="160" bgcolor="#A2DBEE" height="407"><!-- #include file="left.asp"-->
          </td>
          <td width="10" height="407"></td>
          <td width="604" height="407"><form name="form1" method="POST" action="<%=MR_editAction%>">
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="24"> 您现在位置：首页-&gt;栏目设置</td>
                </tr>
                <tr>
                  <td height="24"><div align="center">
                      <script LANUGAGE="JavaScript">
<!--
function js_callpage(htmlurl) {
 var newwin=window.open(htmlurl,"homeWin","toolbar=no,top=0,left=60,width=660,height=500,menubar=no,scrollbars=no,resizable=yes,status=yes");
  newwin.focus();
  return false;
}
//-->
</script>
                      <% If (Request("diy")) <> (ok) Then %>
                      <font color="#FF0000"><span class="pt105"><font color="#000000">栏目设置成功!</font></span></font>
                      <% End If %>
                    </div></td>
                </tr>
                <tr>
                  <td valign="top"><table width="509" border="0" cellspacing="0" cellpadding="1" align="center" class="pt9">
                      <tr>
                        <td width="3%">&nbsp;</td>
                        <td width="23%"><div align="center">
                            <input type="checkbox" name="cb01">
                        </div></td>
                        <td width="42%"><input 
      style="BORDER-RIGHT: rgb(0,0,0) 1px solid; BORDER-TOP: rgb(0,0,0) 1px solid; BORDER-LEFT: rgb(0,0,0) 1px solid; BORDER-BOTTOM: rgb(0,0,0) 1px solid; HEIGHT: 18px" 
      size=20 name=tf01 maxlength="20" value="<%=(Rs.Fields.Item("title1").Value)%>" class="pt9">
                          <% If (Rs.Fields.Item("yin_1").Value) = ("Y") Then %>
                          <font color="#FF0000">被隐藏</font>
                          <% End If %>                        </td>
                        <td width="22%"><div align="left">
                            <input type="button" value="[内容编辑]" name="sm01" class=NoLineInput style="cursor:hand"  onClick="return js_callpage(this.href);" href="../editor/lm1_editor.asp" target="_blank">
                        </div></td>
                      </tr>
                      <tr>
                        <td width="3%">&nbsp;</td>
                        <td width="23%"><div align="center">
                            <input type="checkbox" name="cbxw" value="checkbox">
                        </div></td>
                        <td width="42%"><input 
      style="BORDER-RIGHT: rgb(0,0,0) 1px solid; BORDER-TOP: rgb(0,0,0) 1px solid; BORDER-LEFT: rgb(0,0,0) 1px solid; BORDER-BOTTOM: rgb(0,0,0) 1px solid; HEIGHT: 18px" 
      size=20 name=tf03 maxlength="20" value="<%=(Rs.Fields.Item("title_news").Value)%>" class="pt9">
                          <% If (Rs.Fields.Item("yin_news").Value) = ("Y") Then 'script %>
                          <font color="#FF0000">被隐藏</font>
                          <% End If %>                        </td>
                        <td width="22%"><div align="left">
                            <input type="button" value="[内容编辑]" name="sm03" class=NoLineInput style="cursor:hand"  onClick="return js_callpage(this.href);" href="../editor/news_editor.asp" target="_blank">
                        </div></td>
                      </tr>
                      <tr>
                        <td width="3%">&nbsp;</td>
                        <td width="23%"><div align="center">
                            <input type="checkbox" name="cb04" value="checkbox">
                        </div></td>
                        <td width="42%"><input 
      style="BORDER-RIGHT: rgb(0,0,0) 1px solid; BORDER-TOP: rgb(0,0,0) 1px solid; BORDER-LEFT: rgb(0,0,0) 1px solid; BORDER-BOTTOM: rgb(0,0,0) 1px solid; HEIGHT: 18px" 
      size=20 name=tf04 maxlength="20" value="<%=(Rs.Fields.Item("title3").Value)%>" class="pt9">
                          <% If (Rs.Fields.Item("yin_3").Value) = ("Y") Then %>
                          <font color="#FF0000">被隐藏</font>
                          <% End If %>                        </td>
                        <td width="22%"><div align="left">
                            <input type="button" value="[内容编辑]" name="sm04" class=NoLineInput style="cursor:hand"  onClick="return js_callpage(this.href);" href="../editor/lm3_editor.asp" target="_blank">
                        </div></td>
                      </tr>
                      <tr>
                        <td width="3%">&nbsp;</td>
                        <td width="23%"><div align="center">
                            <input type="checkbox" name="cb05" value="checkbox">
                        </div></td>
                        <td width="42%"><input 
      style="BORDER-RIGHT: rgb(0,0,0) 1px solid; BORDER-TOP: rgb(0,0,0) 1px solid; BORDER-LEFT: rgb(0,0,0) 1px solid; BORDER-BOTTOM: rgb(0,0,0) 1px solid; HEIGHT: 18px" 
      size=20 name=tf05 maxlength="20" value="<%=(Rs.Fields.Item("title4").Value)%>" class="pt9">
                          <% If (Rs.Fields.Item("yin_4").Value) = ("Y") Then %>
                          <font color="#FF0000">被隐藏</font>
                          <% End If %>                        </td>
                        <td width="22%"><div align="left">
                            <input type="button" value="[内容编辑]" name="sm05" class=NoLineInput style="cursor:hand"  onClick="return js_callpage(this.href);" href="../editor/lm4_editor.asp" target="_blank">
                        </div></td>
                      </tr>
                      <tr>
                        <td width="3%">&nbsp;</td>
                        <td width="23%"><div align="center">
                            <input type="checkbox" name="cb06" value="checkbox">
                        </div></td>
                        <td width="42%"><input 
      style="BORDER-RIGHT: rgb(0,0,0) 1px solid; BORDER-TOP: rgb(0,0,0) 1px solid; BORDER-LEFT: rgb(0,0,0) 1px solid; BORDER-BOTTOM: rgb(0,0,0) 1px solid; HEIGHT: 18px" 
      size=20 name=tf06 maxlength="20" value="<%=(Rs.Fields.Item("title5").Value)%>" class="pt9">
                          <% If (Rs.Fields.Item("yin_5").Value) = ("Y") Then %>
                          <font color="#FF0000">被隐藏</font>
                          <% End If%>                        </td>
                        <td width="22%"><div align="left">
                            <input type="button" value="[内容编辑]" name="sm06" class=NoLineInput style="cursor:hand"  onClick="return js_callpage(this.href);" href="../editor/lm5_editor.asp" target="_blank">
                        </div></td>
                      </tr>
                      <tr>
                        <td width="3%">&nbsp;</td>
                        <td width="23%"><div align="center">
                            <input type="checkbox" name="cb07" value="checkbox">
                        </div></td>
                        <td width="42%"><input 
      style="BORDER-RIGHT: rgb(0,0,0) 1px solid; BORDER-TOP: rgb(0,0,0) 1px solid; BORDER-LEFT: rgb(0,0,0) 1px solid; BORDER-BOTTOM: rgb(0,0,0) 1px solid; HEIGHT: 18px" 
      size=20 name=tf07 maxlength="20" value="<%=(Rs.Fields.Item("title6").Value)%>" class="pt9">
                          <% If (Rs.Fields.Item("yin_6").Value) = ("Y") Then %>
                          <font color="#FF0000">被隐藏</font>
                          <% End If %>                        </td>
                        <td width="22%"><div align="left">
                            <input type="button" value="[内容编辑]" name="sm07" class=NoLineInput style="cursor:hand"  onClick="return js_callpage(this.href);" href="../editor/lm6_editor.asp" target="_blank">
                        </div></td>
                      </tr>
                      <tr>
                        <td width="3%">&nbsp;</td>
                        <td width="23%"><div align="center">
                            <input type="checkbox" name="cb09" value="checkbox">
                        </div></td>
                        <td width="42%"><input 
      style="BORDER-RIGHT: rgb(0,0,0) 1px solid; BORDER-TOP: rgb(0,0,0) 1px solid; BORDER-LEFT: rgb(0,0,0) 1px solid; BORDER-BOTTOM: rgb(0,0,0) 1px solid; HEIGHT: 18px" 
      size=20 name=tf09 maxlength="20" value="<%=(Rs.Fields.Item("title7").Value)%>" class="pt9">
                          <% If (Rs.Fields.Item("yin_7").Value) = ("Y") Then %>
                          <font color="#FF0000">被隐藏</font>
                          <% End If%>                        </td>
                        <td width="22%"><div align="left"> </div></td>
                      </tr>
                      <tr>
                        <td width="3%" height="40">&nbsp;</td>
                        <td width="23%" height="40">&nbsp;</td>
                        <td width="42%" height="40"><input type=submit  value=" 确 定 " name=Submit></td>
                        <td width="22%" height="40"><input type=reset  value=" 重 置 " name=Submit2>                        </td>
                        <td width="10%" height="40"></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                </tr>
              </table>
              <input type="hidden" name="MR_update" value="true">
              <input type="hidden" name="MR_recordId" value="<%= Rs.Fields.Item("idd").Value %>">
            </form></td>
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
