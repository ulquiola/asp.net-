<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn/conn.asp" -->
<!--#include file="js/Function.asp" -->
<%
Dim NewMsgList
Dim NewMsgList_numRows

Set NewMsgList = Server.CreateObject("ADODB.Recordset")
NewMsgList.ActiveConnection = conn
NewMsgList.Source="SELECT * FROM tb_talk where (fuserid="&request.QueryString("UserID")&" and ToUserID="&request.Cookies("UserID")&") and Isreaded = '0' order by id desc"
NewMsgList.CursorType = 0
NewMsgList.CursorLocation = 2
NewMsgList.LockType = 1
NewMsgList.Open()
NewMsgList_numRows = 0
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script>
function closeAlert(){
		try{
				MsgAlert.close()
		}
		catch(e){
			return;
		}
}
</script>
<style type="text/css">
<!--
.STYLE2 {font-size: 9pt}
-->
</style>
</head>

<body topmargin="2" onLoad="closeAlert()">
<span class="STYLE2">
<%
NewMsgID = ""
while not NewMsgList.EOF
mes = NewMsgList.fields.item("Message").value
%>
<font color="#FF0000"><%=NewMsgList("senddate")%>&nbsp;<%=NewMsgList("fusernickname")%> </font><br>
<font color="#3333FF"><%=mes%> </font><br>
<%
NewMsgID = NewMsgList("ID") & "," & NewMsgID
NewMsgList.movenext()
wend
if  len(NewMsgID)>2 then
NewMsgID=left(NewMsgID,len(NewMsgID)-1)
end if
%>
</span>
</body>
</html>
<%
NewMsgList.Close()
Set NewMsgList = Nothing

if NewMsgID <> "" then
    MM_editQuery  = "Update tb_talk set isReaded = '1' where id in ("&NewMsgID&")"
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = conn
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
    Set MM_editCmd = nothing
End if
%>
