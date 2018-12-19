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
  MR_editRedirectUrl = "webset.asp?diy=ok"
  MR_fieldsStr  = "title|value|titlefont|value|titlesize|value|TableBgcolorValue|value|aTableBgcolorValue|value|titleway|value|cssopen|value"
  MR_columnsStr = "title|',none,''|titlefont|',none,''|titlesize|',none,''|titlecolor|',none,''|titlecss|',none,''|titleway|',none,''|cssopen|',none,''"
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
.STYLE1 {color: #666666}
.STYLE4 {color: #000000}
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
<script type="text/javascript">
function Mycheck(){
  if (document.SetColorForm.title.readOnly){
    alert("此文本框为只读属性!");return false;
  }
  else return
    return true;
}
</script>
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
                <td height="24"><span class="STYLE1"> 您现在位置：首页-&gt;基本设置
                  <script language="JavaScript">
if (document.createAttribute) {
	document.writeln("<object id=\"dlgColHelper\" classid=\"clsid:3050f819-98b5-11cf-bb82-00aa00bdce0b\" width=\"0px\" height=\"0px\"></object>");
}

function SetTableBgcolor(el) {
	var return_value = web_getCol();
	if (return_value != null){
		el.style.backgroundColor = return_value;
		document.SetColorForm.TableBgcolorInput.style.backgroundColor = return_value;
		document.SetColorForm.TableBgcolorValue.value = return_value;
	}
}

function SetaTableBgcolor(el) {
	var return_value = web_getCol();
	if (return_value != null){
		el.style.backgroundColor = return_value;
		document.SetColorForm.aTableBgcolorInput.style.backgroundColor = return_value;
		document.SetColorForm.aTableBgcolorValue.value = return_value;
	}
}

function web_getCol() {
	if (document.createAttribute) {
		var retCol = dlgColHelper.ChooseColorDlg();
		retCol = retCol.toString(16);
		if( retCol == 0 )
			return null;
		if( retCol.length < 6 ) {
			var tmp = "000000".substring(0,6-retCol.length);
			retCol = tmp.concat(retCol);
		}
	}else{
		if (bascol != null) {
			retCol = bascol;
		}
	}

	var Hclr = retCol.charAt(0);
	if( Hclr != "#" )
		retCol = "#" + retCol;
	return retCol;
}
                    </script>
                  </span></td>
              </tr>
              <tr>
                <td height="20"><div align="center">
                    <% If Rs.Fields.Item("banner").Value <> ("000.bmp") Then %>
                    <span class="STYLE4">文字<font face="Arial, Helvetica, sans-serif">Banner</font>处于屏蔽状态,当前网站使用图片<font face="Arial, Helvetica, sans-serif">Banner</font>模式!</span>
                    <% End If %>
                  </div></td>
              </tr>
              <tr>
                <td><form method="POST" action="<%=MR_editAction%>" name="SetColorForm">
                    <table width="550" cellspacing="2" cellpadding="1" align="center" bgcolor="#A2DBEE">
                      <tr>
                        <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="#A2DBEE">
                            <tr bgcolor="#FFFFFF">
                              <td bgcolor="#FFFFFF" width="26%"><div align="right">网站名称：</div></td>
                              <td><input type="text" name="title" style="BORDER-RIGHT: rgb(0,0,0) 1px solid; BORDER-TOP: rgb(0,0,0) 1px solid; BORDER-LEFT: rgb(0,0,0) 1px solid; BORDER-BOTTOM: rgb(0,0,0) 1px solid; HEIGHT: 20px" value="<%=(Rs.Fields.Item("title").Value)%>" class="pt9" size="32" maxlength="30" readonly="true" onMouseDown="Mycheck()" onMouseUp="Mycheck()">
                              </td>
                            </tr>
                            <tr bgcolor="#FFFFFF">
                              <td bgcolor="#FFFFFF" width="26%"><div align="right">字体选择：</div></td>
                              <td><select name="titlefont" class="pt9">
                                  <option value="隶书"  <%If (Rs.Fields.Item("titlefont").Value = "隶书") Then Response.Write("SELECTED") : Response.Write("")%>>隶书</option>
                                  <option value="宋体"  <%If (Rs.Fields.Item("titlefont").Value = "宋体") Then Response.Write("SELECTED") : Response.Write("")%>>宋体</option>
                                  <option value="楷体_GB2312"  <%If (Rs.Fields.Item("titlefont").Value = "楷体_GB2312") Then Response.Write("SELECTED") : Response.Write("")%>>楷体</option>
                                  <option value="仿宋_GB2312"  <%If (Rs.Fields.Item("titlefont").Value = "仿宋_GB2312") Then Response.Write("SELECTED") : Response.Write("")%>>仿宋</option>
                                  <option value="幼圆"  <%If (Rs.Fields.Item("titlefont").Value = "幼圆") Then Response.Write("SELECTED") : Response.Write("")%>>幼圆</option>
                                </select>
                              </td>
                            </tr>
                            <tr bgcolor="#FFFFFF">
                              <td bgcolor="#FFFFFF" width="26%"><div align="right">字体大小：</div></td>
                              <td><input type="text" name="titlesize" maxlength="2" style="BORDER-RIGHT: rgb(0,0,0) 1px solid; BORDER-TOP: rgb(0,0,0) 1px solid; BORDER-LEFT: rgb(0,0,0) 1px solid; BORDER-BOTTOM: rgb(0,0,0) 1px solid; font-family:Fixedsys; WIDTH:50px; HEIGHT: 20px" value="<%=(Rs.Fields.Item("titlesize").Value)%>">
                                磅</td>
                            </tr>
                            <tr bgcolor="#FFFFFF">
                              <td bgcolor="#FFFFFF" width="26%"><div align="right">字体颜色：</div></td>
                              <td><input type=test name="TableBgcolorValue" 
				style="BORDER-RIGHT: rgb(0,0,0) 1px solid; BORDER-TOP: rgb(0,0,0) 1px solid; BORDER-LEFT: rgb(0,0,0) 1px solid; BORDER-BOTTOM: rgb(0,0,0) 1px solid; font-family:Fixedsys; WIDTH:80px; HEIGHT: 20px" value="<%=(Rs.Fields.Item("titlecolor").Value)%>">
                                <input type=button name="Input"
				style="background-color:<%=(Rs.Fields.Item("titlecolor").Value)%>" 
				id="TableBgcolorInput" value="    " 
				onClick="SetTableBgcolor(TableBgcolorInput)" 
				onMouseOver="this.style.cursor='hand';">
                              </td>
                            </tr>
                            <tr bgcolor="#FFFFFF">
                              <td bgcolor="#FFFFFF" width="26%"><div align="right">字影颜色：</div></td>
                              <td><input type=test name="aTableBgcolorValue" 
				style="BORDER-RIGHT: rgb(0,0,0) 1px solid; BORDER-TOP: rgb(0,0,0) 1px solid; BORDER-LEFT: rgb(0,0,0) 1px solid; BORDER-BOTTOM: rgb(0,0,0) 1px solid; font-family:Fixedsys; WIDTH:80px; HEIGHT: 20px" value="<%=(Rs.Fields.Item("titlecss").Value)%>">
                                <input type=button name="Input2"
				style="background-color:<%=(Rs.Fields.Item("titlecss").Value)%>" 
				id="aTableBgcolorInput" 
				onClick="SetaTableBgcolor(aTableBgcolorInput)" 
				onMouseOver="this.style.cursor='hand';" value="    ">
                              </td>
                            </tr>
                            <tr bgcolor="#FFFFFF">
                              <td bgcolor="#FFFFFF" width="26%"><div align="right">字影方向：</div></td>
                              <td><select name="titleway" class="pt9">
                                  <option value="270"  <%If (Rs.Fields.Item("titleway").Value = "270") Then Response.Write("SELECTED") : Response.Write("")%>>左</option>
                                  <option value="315"  <%If (Rs.Fields.Item("titleway").Value = "315") Then Response.Write("SELECTED") : Response.Write("")%>>左上</option>
                                  <option value="0"  <%If (Rs.Fields.Item("titleway").Value = "0") Then Response.Write("SELECTED") : Response.Write("")%>>上</option>
                                  <option value="45"  <%If (Rs.Fields.Item("titleway").Value = "45") Then Response.Write("SELECTED") : Response.Write("")%>>右上</option>
                                  <option value="90"  <%If (Rs.Fields.Item("titleway").Value = "90") Then Response.Write("SELECTED") : Response.Write("")%>>右</option>
                                  <option value="135"  <%If (Rs.Fields.Item("titleway").Value = "135") Then Response.Write("SELECTED") : Response.Write("")%>>右下</option>
                                  <option value="180"  <%If (Rs.Fields.Item("titleway").Value = "180") Then Response.Write("SELECTED") : Response.Write("")%>>下</option>
                                  <option value="225"  <%If (Rs.Fields.Item("titleway").Value = "225") Then Response.Write("SELECTED") : Response.Write("")%>>左下</option>
                                </select>
                              </td>
                            </tr>
                            <tr bgcolor="#FFFFFF">
                              <td bgcolor="#FFFFFF" width="26%"><div align="right">字影开关：</div></td>
                              <td><table align="left" cellpadding="0" cellspacing="0" border="0" width="120">
                                  <tr>
                                    <td><input type="radio" value="shadow" name="cssopen" <%If (CStr("shadow") = CStr(Rs.Fields.Item("cssopen").Value)) Then Response.Write("CHECKED") : Response.Write("")%>>
                                      开启
                                    <td><input type="radio" value="shadowclose" name="cssopen" <%If (CStr("shadowclose") = CStr(Rs.Fields.Item("cssopen").Value)) Then Response.Write("CHECKED") : Response.Write("")%>>
                                      关闭 
                                  </tr>
                                </table></td>
                            </tr>
                            <tr bgcolor="#FFFFFF">
                              <td bgcolor="#FFFFFF" width="26%">&nbsp;</td>
                              <td><input type=submit  value=" 确  定 " name=Submit>
                                <input type=reset o value=" 恢  复 " name=Submit2>
                              </td>
                            </tr>
                          </table></td>
                      </tr>
                    </table>
                    <input type="hidden" name="MR_update" value="true">
                    <input type="hidden" name="MR_recordId" value="<%= Rs.Fields.Item("idd").Value %>">
                  </form></td>
              </tr>
            </table>
            <iframe name="banner" align="center" width="580" height="210" noresize scrolling="no" frameborder="0" marginheight="0" marginwidth="0" src="webbanner.asp"></iframe></td>
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
