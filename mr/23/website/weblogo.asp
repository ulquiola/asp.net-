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
  MR_editRedirectUrl = "weblogo.asp"
  MR_fieldsStr  = "MyLogo|value|w|value|h|value"
  MR_columnsStr = "logo|',none,''|logo_w|',none,''|logo_h|',none,''"
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
<style type="text/css">
<!--
.pt9 {  font-size: 9pt; font-family: "宋体"}
.pt105 {  font-size: 10.5pt; font-family: "宋体"}
.pt9n {  font-size: 9pt; line-height: 13pt; font-family: "宋体"}
.pt9nx {  font-size: 9pt; line-height: 15pt; font-family: "宋体"}
-->
</style>
<script language="javascript"> 
function r(){
location.href="weblogo.asp";
}
</script> 

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
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" onLoad="startClock()" onLoad="d()">
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
                <td height="24"> 您现在位置：首页-&gt;网站图标</td>
              </tr>
              <tr>
                <td height="24"><div align="center">
                    <script>
<!--
function ShowOrderUser( UserName ){
	var NewWin = "OrderUserInfo.asp?OrderName=" + UserName;	
	showModalDialog(NewWin,"","dialogWidth:25;dialogHeight:14;dialogTop:100;dialogLeft:200;status:no;");

	return;
}
//-->
</script>
                    <script language="JavaScript">
<!--
function Juge(SetColorForm)
{
  if (SetColorForm.w.value == "")
  {
    alert("请输入图片的宽度,例如120");
    SetColorForm.w.focus();
    return (false);
  }
  if (SetColorForm.h.value == "")
  {
    alert("请输入图片的高度,例如60");
    SetColorForm.h.focus();
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
                    <% If Rs.Fields.Item("logo").Value = ("0.bmp") Then 'script %>
                    <font color="#FF0000" class="pt105"><font color="#000000">网站<font face="Arial, Helvetica, sans-serif">LOGO</font>处于屏蔽状态!</font></font><b><font color="#FF0000" class="pt105"> </font></b>
                    <% End If%>
                    <% If Rs.Fields.Item("logo").Value = ("") Then 'script %>
                    <font color="#FF0000"><span class="pt105"><font color="#000000">网站<font face="Arial, Helvetica, sans-serif">LOGO</font>设置错误,您没有选择或屏蔽<font face="Arial, Helvetica, sans-serif">LOGO</font>!</font></span></font>
                    <% End If%>
                  </div></td>
              </tr>
              <tr>
                <td><form name="SetColorForm" method="POST" action="<%=MR_editAction%>" onSubmit="return Juge(this)">
                    <table width="90%" border="0" cellspacing="2" cellpadding="1" align="center" bgcolor="#A2DBEE">
                      <tr>
                        <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="#A2DBEE">
                            <tr>
                              <td bgcolor="#FFFFFF" width="20%"><div align="right">网站图标：</div></td>
                              <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="2" cellpadding="6">
                                  <tr>
                                    <td colspan="2"><span id="lg"><img src="photos/<%=(Rs.Fields.Item("logo").Value)%>"></span></td>
                                  </tr>
                                  <tr>
                                    <td width="52%"><input type="button" onClick="openImageList('../editor/library.asp?action=select','lg',0,0);" 
				value="选择LOGO" title="选择一个图片" name="button">
                                      <input type="button" onClick="javascript:SetLogoEmpty();" value="不使用LOGO" name="button2">
                                      <script language="JavaScript">
var strInsertTo = "";
var lngWidth;
var lngHeigth;

//打开附件列表窗
function openImageList( href, InsertTo, ThisW, ThisH ) {
	strInsertTo = InsertTo;
	lngWidth = ThisW;
	lngHeigth = ThisH;
	var newwin=window.open(href,"","toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=500,height=200,top=0,left=0");
	newwin.focus();
	return;
}

//接收返回的图片链接地址
function receiveVal(val) {
	var str = "";
	var ArrExt = val.split(".");
	var Ext = ArrExt[ArrExt.length-1];
	Ext = Ext.toLowerCase();

	if( (Ext != "swf") && (Ext != "gif") && (Ext != "jpg") && (Ext != "bmp") ) {
		alert("您选择的可能不是图片文件，这样将无法正常显示！");
		return false;
	}


	// 测试文件是否放在站点的根目录
	var strTest = val;
	var isAtRoot = true;
	var ArrTest = strTest.split("/");
	if( strTest.charAt(0) == "/" ) {
		if( ArrTest.length>4 )
			isAtRoot = false;
	}else{
		if( ArrTest.length>3 )
			isAtRoot = false;
	}

	if( Ext == "swf" ) {
		str = "<embed src=\"" + val + "\" ";
		str += "pluginspage=\"http://www.macromedia.com/shockwave/download/index.cgi";
		str += "?P1_Prod_Version=ShockwaveFlash\" type=\"application/x-shockwave-flash\"";
		if( lngWidth != 0 && lngHeigth != 0 ) {
			str += " width=" + lngWidth + " height=" + lngHeigth;
		}
		str += "></embed>";
	}else{
		str = "<img src=\"" + val + "\" border=0";
		if( lngWidth != 0 && lngHeigth != 0 ) {
			str += " width=" + lngWidth + " height=" + lngHeigth;
		}
		str += ">";
	}

	var ArrFileName = val.split("/");
	var FileName = ArrFileName[ArrFileName.length-1];

	if( strInsertTo == "lg" ) {
		document.all.lg.innerHTML = str;
		document.all.SetColorForm.MyLogo.value = FileName;
	}
	if( strInsertTo == "bnr" ) {
		document.all.bnr.innerHTML = str;
		document.all.SetColorForm.MyBanner.value = FileName;
	}
}

function SetLogoEmpty() {
	var isGo = confirm("您的网站以后将不再显示LOGO,(需要时请再次选择),是否确定?");
	if( isGo ) {
		document.all.lg.innerHTML = "<img src=\"photos/0.bmp\">";
		document.all.SetColorForm.MyLogo.value = "0.bmp";
	}
}
</script>                                    </td>
                                    <td width="48%">上传图片宽小于120，高小于30</td>
                                  </tr>
                                </table></td>
                            </tr>
                            <tr>
                              <td bgcolor="#FFFFFF" width="20%"><div align="right">调整尺寸：</div></td>
                              <td bgcolor="#FFFFFF">宽度
                                <input type="text" name="w" style="BORDER-RIGHT: rgb(0,0,0) 1px solid; BORDER-TOP: rgb(0,0,0) 1px solid; BORDER-LEFT: rgb(0,0,0) 1px solid; BORDER-BOTTOM: rgb(0,0,0) 1px solid; HEIGHT: 17px" class="pt9" size="3" maxlength="3" value="<%=(Rs.Fields.Item("logo_w").Value)%>">
                                × 高度
                                <input type="text" name="h" style="BORDER-RIGHT: rgb(0,0,0) 1px solid; BORDER-TOP: rgb(0,0,0) 1px solid; BORDER-LEFT: rgb(0,0,0) 1px solid; BORDER-BOTTOM: rgb(0,0,0) 1px solid; HEIGHT: 17px" class="pt9" size="3" maxlength="3" value="<%=(Rs.Fields.Item("logo_h").Value)%>">
                                &nbsp;&nbsp;<font color="#999999">例如:60×30</font></td>
                            </tr>
                            <tr>
                              <td bgcolor="#FFFFFF" width="20%" height="34">&nbsp;</td>
                              <td bgcolor="#FFFFFF" height="34">&nbsp;
                                <input type="button" 
								onClick="javascript:location.reload()" value="  刷 新  ">
                                <input type=submit 
								value="  确 定  " name=Submit>
                              </td>
                            </tr>
                          </table></td>
                      </tr>
                    </table>
                    <input type="hidden" name="MR_update" value="true">
                                      <input type="hidden" name="MyLogo" value="<%=rs("logo")%>">
                    <input type="hidden" name="MR_recordId" value="<%= Rs.Fields.Item("idd").Value %>">
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
