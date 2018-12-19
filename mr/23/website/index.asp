<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conn/conn.asp" -->
<%
MR_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MR_LoginAction = MR_LoginAction + "?" + Request.QueryString
MR_valUsername=CStr(Request.Form("username"))
If MR_valUsername <> "" Then
  MR_fldUserAuthorization=""
  MR_redirectLoginSuccess="templet.asp"
  MR_redirectLoginFailed="log_error.htm"
  MR_flag="ADODB.Recordset"
  set MR_rsUser = Server.CreateObject(MR_flag)
  MR_rsUser.ActiveConnection = MR_website_STRING
  MR_rsUser.Source = "SELECT id, password"
  If MR_fldUserAuthorization <> "" Then MR_rsUser.Source = MR_rsUser.Source & "," & MR_fldUserAuthorization
  MR_rsUser.Source = MR_rsUser.Source & " FROM website WHERE id='" & Replace(MR_valUsername,"'","''") &"' AND password='" & Replace(Request.Form("password"),"'","''") & "'"
  MR_rsUser.CursorType = 0
  MR_rsUser.CursorLocation = 2
  MR_rsUser.LockType = 3
  MR_rsUser.Open
  If Not MR_rsUser.EOF Or Not MR_rsUser.BOF Then 
    Session("MR_Username") = MR_valUsername
    If (MR_fldUserAuthorization <> "") Then
      Session("MR_UserAuthorization") = CStr(MR_rsUser.Fields.Item(MR_fldUserAuthorization).Value)
    Else
      Session("MR_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MR_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MR_rsUser.Close
    Response.Redirect(MR_redirectLoginSuccess)
  End If
  MR_rsUser.Close
  Response.Redirect(MR_redirectLoginFailed)
End If
%>
<html>
<head>
<title>自动建站</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link 
href="css/css.css" rel=stylesheet type="text/css">
<style type="text/css">
<!--
.pt9 {  font-family: "宋体"; font-size: 9pt}
.pt9n {  font-family: "宋体"; font-size: 9pt; line-height: 13pt}
-->
</style>
</head>
<body bgcolor="#FFFFFF" text="#000000">
<form name="form1" method="post" action="<%=MR_LoginAction%>">
  <table width="214" border="0" align="center">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="center"><table width="400" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="222" align="right" background="image/login.jpg"><table width="29%" border="0" cellspacing="0" cellpadding="0" align="right">
          <tr>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td><div align="center">用 户：
              <input type="text" name="username" class="pt9" size="14" maxlength="12">
            </div></td>
          </tr>
          <tr>
            <td height="2"></td>
          </tr>
          <tr>
            <td><div align="center">密 码：
              <input type="password" name="password" class="pt9" size="14" maxlength="12">
            </div></td>
          </tr>
          <tr>
            <td height="8"></td>
          </tr>
          <tr>
            <td><table width="200" border="0" cellspacing="0" cellpadding="0" align="center">
                <tr>
                  <td width="44"></td>
                  <td width="156"><a href="reg1.asp" onClick="window.open(this.href,'','location=no,menu=no,scrollbars=no,resizable=no,top=0,left=200,width=430,height=330');return false;"><img src="image/register.gif" alt="54" width="55" height="18" border="0" galleryimg="no"></a>
                      <input type="image" name="Submit" value="登 录" src="image/denglu.gif" width="55" height="18">
                  </td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
          </tr>
        </table></td>
      </tr>
    </table>
    </form>
</body>
</html>
