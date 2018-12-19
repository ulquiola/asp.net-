<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn/conn.asp" -->
<!--#include file="js/Function.asp" -->
<%
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.ActiveConnection = conn 
rs.Source = "SELECT * from v_gmeber where GroupID="&request.querystring("ID")&""
rs.CursorType =0
rs.CursorLocation =2
rs.LockType =1
rs.Open() 
%>
<HTML>
<script>
function dodelete(){
	document.form1.action = "addfriend_del.asp" 
	document.form1.submit() 
}
</script>
<style type="text/css">
<!--
.STYLE2 {
	font-size: 9pt;
	color: #000000;
}
.STYLE3 {width:8%;font:8pt/14pt Georgia,Verdana,"sans-serif","宋体";color:#3C404D;text-align:center;
}
.STYLE4 {
	font-size: 9pt;
	color: #FF6600;
}
-->
</style>
<BODY>
<form name="form1" action="" method="post">
<table border="1" cellpadding="0" cellspacing="1" width="100%" id="AutoNumber1">
  <tr>
    <td width="25%" align="center"><font size="2" class="STYLE2">操作</font></td>
    <td width="25%" align="center"><span class="STYLE2">用户ID</span><font size="2">　</font></td>
    <td width="25%" align="center"><font size="2" class="STYLE2">用户名称</font></td>
    <td width="25%" align="center"><font size="2" class="STYLE2">用户昵称</font></td>
  </tr>
  <%if rs.EOF then%>
  <tr>
    <td width="100%" colspan="4"><font size="2"><span class="STYLE2">该组暂无好友!</span></td>
  </tr>
  <%end if%>
  <%while not rs.EOF%>
  <tr>
    <td width="25%" align="center"><font><input type="checkbox" name="MemberID" value="<%=rs("ID")%>"></font></td>
    <td width="25%" align="center"><font class="STYLE3"><%=rs("id")%></font></td>
    <td width="25%" align="center"><font  class="STYLE2"><%=rs("UserName")%> </font></td>
    <td width="25%" align="center" class="STYLE4"><font class="STYLE4"><%=rs("state")%> </font></td>
  </tr>
  <%
  rs.MoveNext()
  wend
  %>
</table>
<input type="hidden" name="groupID" value="<%=request("ID")%>">
</form>
</BODY>
</HTML>
