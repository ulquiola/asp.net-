<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Conn/conn.asp" -->
<%
MR_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MR_editAction = MR_editAction & "?" & Request.QueryString
End If
MR_abortEdit = false
MR_editQuery = ""
%>
<%
MR_flag="MR_insert"
If (CStr(Request(MR_flag)) <> "") Then
  MR_dupKeyRedirect="user_exist.asp"
  MR_rsKeyConnection=MR_website_STRING
  MR_dupKeyUsernameValue = CStr(Request.Form("username"))
  MR_dupKeySQL="SELECT id FROM website WHERE id='" & MR_dupKeyUsernameValue & "'"
  MR_adodbRecordset="ADODB.Recordset"
  set MR_rsKey=Server.CreateObject(MR_adodbRecordset)
  MR_rsKey.ActiveConnection=MR_rsKeyConnection
  MR_rsKey.Source=MR_dupKeySQL
  MR_rsKey.CursorType=0
  MR_rsKey.CursorLocation=2
  MR_rsKey.LockType=3
  MR_rsKey.Open
  If Not MR_rsKey.EOF Or Not MR_rsKey.BOF Then 
    MR_qsChar = "?"
    If (InStr(1,MR_dupKeyRedirect,"?") >= 1) Then MR_qsChar = "&"
    MR_dupKeyRedirect = MR_dupKeyRedirect & MR_qsChar & "requsername=" & MR_dupKeyUsernameValue
    Response.Redirect(MR_dupKeyRedirect)
  End If
  MR_rsKey.Close
End If
%>
<%
If (CStr(Request("MR_insert")) <> "") Then
  MR_editConnection = MR_website_STRING
  MR_editTable = "website"
  MR_editRedirectUrl = "reg2.asp"
  MR_fieldsStr  = "username|value|password|value|repassword|value|name|value|address|value|phone|value|fax|value|yin_1|value|yin_3|value|yin_4|value|yin_5|value|yin_6|value|yin_7|value|yin_news|value|yin_zimu|value|mobile|value|templet|value|hit|value|titlesize|value|titlefont|value|titlecolor|value|titlecss|value|cssopen|value|on_off|value|zm|value|title|value|font1|value|email|value|ip|value|title1|value|title3|value|title4|value|title5|value|title6|value|title_news|value|title7|value|logo|value|banner|value"
  MR_columnsStr = "id|',none,''|password|',none,''|repassword|',none,''|name|',none,''|address|',none,''|phone|',none,''|fax|',none,''|yin_1|',none,''|yin_3|',none,''|yin_4|',none,''|yin_5|',none,''|yin_6|',none,''|yin_7|',none,''|yin_news|',none,''|yin_zimu|',none,''|mobile|',none,''|xuanze|',none,''|hit|',none,''|titlesize|',none,''|titlefont|',none,''|titlecolor|',none,''|titlecss|',none,''|cssopen|',none,''|on_off|',none,''|zimu|',none,''|title|',none,''|font1|',none,''|email|',none,''|ip|',none,''|title1|',none,''|title3|',none,''|title4|',none,''|title5|',none,''|title6|',none,''|title_news|',none,''|title7|',none,''|logo|',none,''|banner|',none,''"
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
If (CStr(Request("MR_insert")) <> "") Then
  MR_tableValues = ""
  MR_dbValues = ""
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
      MR_tableValues = MR_tableValues & ","
      MR_dbValues = MR_dbValues & ","
    End if
    MR_tableValues = MR_tableValues & MR_columns(i)
    MR_dbValues = MR_dbValues & FormVal
  Next
  MR_editQuery = "insert into " & MR_editTable & " (" & MR_tableValues & ") values (" & MR_dbValues & ")"
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
<html>
<head>
<title>注册1</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link 
href="css/css.css" rel=stylesheet type="text/css">
<style type="text/css">
<!--
.pt9 {  font-family: "宋体"; font-size: 9pt; line-height: normal}
.pt105 {  font-family: "宋体"; font-size: 10.5pt}
-->
</style>
</head>
<body bgcolor="#FF0000" text="#000000" leftmargin="0" topmargin="3">
<b><font color="#FFFFFF"> </font></b>
<form action="<%=MR_editAction%>" method="POST" name="" onSubmit="return checkUserInfo(this)">
  <table width="400" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td height="26" class="pt105"><b><font color="#FFFFFF">第一步：填写注册资料</font>
        <script language="JavaScript">
<!--
function checkUserInfo(theForm)
{
  if (theForm.username.value == "")
  {
    alert("请输入用户名!");
    theForm.username.focus();
    return (false);
  }
  if (checktext(theForm.username.value))
  {
    alert("请输入用户名!");
    theForm.username.select();
    theForm.username.focus();
    return (false);
  }
  if (theForm.username.value.length < 2)
  {
    alert("用户名长度最少要2个字符!");
    theForm.username.focus();
    return(false);
  }
  if (theForm.password.value == "")
  {
    alert("请输入密码!");
    theForm.password.focus();
    return (false);
  }
  if (checktext(theForm.password.value))
  {
    alert("请输入密码!");
    theForm.password.select();
    theForm.password.focus();
    return (false);
  }
  if (theForm.password.value.length < 4)
  {
    alert("密码长度最少要4位!");
    theForm.password.focus();
    return(false);
  }
  if (theForm.password.value != theForm.repassword.value)
  {
    alert("您两次输入的密码不一致!");
    theForm.password.focus();
    return(false);
  }
  if (theForm.name.value == "")
  {
    alert("请输入您的姓名!");
    theForm.name.focus();
    return (false);
  }
  if (checktext(theForm.name.value))
  {
    alert("请输入您的姓名!");
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
    alert("请输入联系地址!");
    theForm.address.focus();
    return (false);
  }
  if (checktext(theForm.address.value))
  {
    alert("请输入联系地址!");
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
      </b></td>
    </tr>
  </table>
  <table width="400" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#000000">
    <tr>
      <td bgcolor="#95BFEC"><table align="center" width="400" cellpadding="2" cellspacing="0" class="pt9" border="0">
          <tr bgcolor="#FFFFFF">
            <td nowrap align="right" width="24%" height="10"></td>
            <td height="10"></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td nowrap align="right" width="24%">用户名：&nbsp;</td>
            <td><input type="text" name="username" value="" size="12" maxlength="12" style="font-size: 9pt; border: 1 solid #000000">
              （12位以内英文或数字）<font color="#FF0000">*</font></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td nowrap align="right" width="24%">密&nbsp;&nbsp;码：&nbsp;</td>
            <td><input type="password" name="password" value="" size="12" maxlength="12" style="font-size: 9pt; border: 1 solid #000000">
              （12位以内英文或数字）<font color="#FF0000">*</font></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td nowrap align="right" width="24%">重复密码：&nbsp;</td>
            <td><input type="password" name="repassword" value="" size="12" maxlength="12" style="font-size: 9pt; border: 1 solid #000000">
              （重复输入以上密码）<font color="#FF0000">*</font></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td nowrap align="right" width="24%">联 系 人：&nbsp;</td>
            <td><input type="text" name="name" value="" size="20" maxlength="20" style="font-size: 9pt; border: 1 solid #000000">
              <font color="#FF0000">* </font></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td nowrap align="right" width="24%">网站名称：&nbsp;</td>
            <td><input type="text" name="title" value="" size="43" maxlength="50" style="font-size: 9pt; border: 1 solid #000000">
              <font color="#FF0000">*</font></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td nowrap align="right" width="24%">联系地址：&nbsp;</td>
            <td><input type="text" name="address" value="" size="43" maxlength="50" style="font-size: 9pt; border: 1 solid #000000">
              <font color="#FF0000">*</font></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td nowrap align="right" width="24%">电话号码：&nbsp;</td>
            <td><input type="text" name="phone" value="" size="30" maxlength="30" style="font-size: 9pt; border: 1 solid #000000">
              <font color="#FF0000">*</font></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td nowrap align="right" width="24%">邮编号码：&nbsp;</td>
            <td><input type="text" name="fax" value="" size="30" maxlength="30" style="font-size: 9pt; border: 1 solid #000000">
              <input type="hidden" name="yin_1" value="N">
              <input type="hidden" name="yin_3" value="N">
              <input type="hidden" name="yin_4" value="N">
              <input type="hidden" name="yin_5" value="N">
              <input type="hidden" name="yin_6" value="N">
              <input type="hidden" name="yin_7" value="N">
              <input type="hidden" name="yin_news" value="N">
              <input type="hidden" name="yin_zimu" value="N">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td nowrap align="right" width="24%">手机号码：&nbsp;</td>
            <td><input type="text" name="mobile" size="30" maxlength="30" style="font-size: 9pt; border: 1 solid #000000">
              <input type="hidden" name="templet" value="1">
              <input type="hidden" name="hit" value="1">
              <input type="hidden" name="titlesize" value="56">
              <input type="hidden" name="titlefont" value="隶书">
              <input type="hidden" name="titlecolor" value="#009b00">
              <input type="hidden" name="titlecss" value="#95ffaf">
              <input type="hidden" name="cssopen" value="shadow">
              <input type="hidden" name="on_off" value="关">
              <input type="hidden" name="zm" value="欢迎您光临我们的网站!">
              <input type="hidden" name="font1" value="pt9">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td nowrap align="right" width="24%">电子邮箱：&nbsp;</td>
            <td><input type="text" name="email" value="" size="43" maxlength="50" style="font-size: 9pt; border: 1 solid #000000">
              <font color="#FF0000">*</font>
              <input type="hidden" name="ip" value="<%= Request.ServerVariables("REMOTE_ADDR") %>">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td nowrap align="right" width="24%" height="34">&nbsp;</td>
            <td height="34">&nbsp;
              <input type="submit" value=" 确  定 " class="pt9n" name="submit" >
              <input type="reset" name="Reset" value=" 重  填 " class="pt9n">
              <input type="hidden" name="title1" value="我的档案">
              <input type="hidden" name="title3" value="我的文章">
              <input type="hidden" name="title4" value="我的感悟">
              <input type="hidden" name="title5" value="我的故事">
              <input type="hidden" name="title6" value="联系方式">
              <input type="hidden" name="title_news" value="新闻实事">
              <input type="hidden" name="title7" value="在线维护">
              <input type="hidden" name="logo" value="0.bmp">
              <input type="hidden" name="banner" value="000.bmp">
               </td>
          </tr>
        </table></td>
    </tr>
  </table>
  <input type="hidden" name="MR_insert" value="true">
</form>
</body>
</html>
