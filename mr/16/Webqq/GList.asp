<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<!--#include file="js/Function.asp" -->
<%
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.ActiveConnection = conn 
rs.Source = "SELECT  * FROM tb_groupmain WHERE UserID = "&request.cookies("UserID")&""
rs.CursorType =0
rs.CursorLocation =2
rs.LockType =1
rs.Open() 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<STYLE>

.ItemHL
{
    BORDER-RIGHT: #006699 0px solid;
    BORDER-TOP: #006699 0px solid;
    FONT-SIZE: 10pt;
    BORDER-LEFT: #006699 0px solid;
    CURSOR: hand;
    BORDER-BOTTOM: #006699 0px solid;
    BACKGROUND-COLOR: #c5d5fc
}
.ItemLL
{
    BORDER-RIGHT: #ece9d8 0px solid;
    BORDER-TOP: #ece9d8 0px solid;
    FONT-SIZE: 10pt;
    BORDER-LEFT: #ece9d8 0px solid;
    CURSOR: hand;
    BORDER-BOTTOM: #ece9d8 0px solid;
    BACKGROUND-COLOR: #ece9d8
}
.btm
{
    FONT-SIZE: 10pt;
    CURSOR: hand;
    BACKGROUND-COLOR: #ece9d8
}
.STYLE2 {
	font-size: 9pt;
	color: #000000;
}
.STYLE5 {font-size: 9pt}
body {
	background-color: #f6f7fb;
}
</STYLE>
<script>
function css1(pp){
  pp.orgclassName = pp.className
  pp.className = "ItemHL"
}

function css2(pp){
   pp.className = pp.orgclassName
}

function sun(ID){
	window.parent.frames("right").location = "Group_xianxi.asp?ID=" + ID  
}
</script>
</head>

<body topmargin="2" leftmargin="3">
<form name="form1" method="post" action="DelGroup_deal.asp">
<table border="0" cellpadding="0" cellspacing="0" width="25%">
  <tr>
    <td width="10%" height="16"><img src="images/home.gif" WIDTH="23" HEIGHT="18"  align="absmiddle"></td>
    <td width="90%" height="16" nowrap colspan="2" valign="middle"><span class="STYLE2">我的群组</span></td>
  </tr>
  <%			
  if not rs.EOF then					'判断是否有记录信息
	  if rs.RecordCount = -1 then 
		While (Not rs.EOF)
		rs_Count = rs_Count + 1				'为rs_Count变量赋值
		rs.MoveNext()						
		Wend
		rs.MoveFirst()						'向下移动记录指针
	  end if
  end if
   %>
  <%
  if not rs.EOF then
  for i=1 to rs_Count 
  %>
  <tr>
    <td width="10%" height="16"></td>
    <td width="10%" height="16">
	<%if i=rs_Count Then%><img src="images/End.gif" height="20"><%else%><img src="images/Lines.gif" height="20"><%end if%></td>
    <td width="80%" height="16" nowrap class="STYLE2" style="cursor:hand"  onclick="sun('<%=rs("ID")%>')" onmouseover="css1(this)" onMouseOut="css2(this)"><font size="2">
      <input type="checkbox" name="groupID" value="<%=rs("ID")%>">
      </font><span class="STYLE5"><%=rs("GroupName")%> </span></td>
  </tr>
  <%
  rs.MoveNext()
  Next
  end if
  %>
</table>
</form>
</body>
</html>