<!--#include file="conn/conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>昵称修改</title>
<script>
window.resizeTo("270","180")
function mycheck()
{
   if(document.form1.state.value=="")
   {
      alert("请输入昵称!")
      return
   }
   else
   {
     document.form1.action = "modifyname_deal.asp"
     document.form1.submit()
   }
}
</script>
<style type="text/css">
<!--
.STYLE3 {
	font-size: 9pt;
	color: #000000;
}
-->
</style>
</head>
<%
If request.cookies("UserID") <> "" then
	userid=request.cookies("UserID")
end if
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select state from tb_user where id="&userid
rs.open sql,conn,1,3
%>
<body topmargin="0" leftmargin="0" style="border:none" scroll=no>
<form name="form1" action="" method="POST">
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1" height="100%">
    <tr height="51"> 
      <td colspan="2"></td>
    </tr>
    <tr> 
      <td width="40%" align="right"><span class="STYLE3">用户昵称:</span></td>
      <td width="60%" align="left"><font size="2"> 
        <input name="state" type="text" id="text" style="width:95%" value="<%=rs("state")%>">
        </font></td>
    </tr>
    <tr> 
      <td width="100%" colspan="2" align="right"> <div align="right">
          <input type="button" value="提交" name="B3" onClick="mycheck()">
          <input type="button" value="取消" name="B3" onClick="window.close()">&nbsp;&nbsp;
        </div></td>
    </tr>
  </table>
</form>
</body>

</html>