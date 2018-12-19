<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<!--#include file="JS/Function.asp" -->
<%
        CurrUserID =  ",S" & Request.Cookies("UserID") & "E"
	    application("OnLineUserID") = replace(application("OnLineUserID"),CurrUserID,"")
	    application("OnLineUserID") = replace(application("OnLineUserID"),",SE","")
	    application("OnLineUserID") = application("OnLineUserID") & CurrUserID

Set rs1 = Server.CreateObject("ADODB.Recordset")
rs1.ActiveConnection = conn
rs1.Source = "select * from tb_Groupmain WHERE UserID="&request.Cookies("UserID")&""
rs1.CursorType = 0
rs1.CursorLocation = 2
rs1.LockType = 1
rs1.Open()
set rs=server.CreateObject("adodb.recordset")
sql="select * from tb_user"
rs.open sql,conn,1,3
%>
<html xmlns:sml>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="css/style.css">
<script language="javascript" src='JS/check.js'></script>
<title></title>
<script language="javascript">
	window.resizeTo("140","500")
	function mycheck1(){
	  window.event.returnValue = false;
	}
	
	setInterval("refrash()",10000)
         
        function mycheck2(){
          window.frames("frmrefrash").location ="touxiang.asp"
        }	
        
	function refrash(){
	   window.frames("frmrefrash").location.reload();
	//  window.location.reload();
	}
	
	function mycheckclear(){
		window.open("webQQ_clearApplication.asp?UserID=<%=request.Cookies("UserID")%>","mycheckclear","width=100,height=100")
	}
	function openmessage(){
		readmessage=window.open("blank.asp","readmessage","width=410,height=211")
		readmessage.moveTo(2000,400)
	}
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) { 
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
<style type="text/css">
<!--
.STYLE2 {
	font-size: 9pt;
	color: #000000;
}
-->
</style>
</head>

<body topmargin="0" leftmargin="0"  style="border:none" bgcolor="#ECE9D8" scroll=no class="body" onLoad="setTimeout('mycheck2()',5000);openmessage()" onUnload="mycheckclear();readmessage.close()">
<div id="Layer1" style="position:absolute; width:200px; height:115px; z-index:1; visibility: hidden;">
<iframe name="frmrefrash" src=""></iframe>
</div>
<div id="ZhuM">
<table border="1" width="100%" height="100%">
  <tr> 
    <td width="100%" height="54" colspan="2"><img src="Images\head\<%=Session("touxiang")%>.gif">[<%=Session("state")%>]<span class="STYLE2">在线</span></td>
  </tr>
  <tr>
    <td colspan="2" id="QQmenuTD">
<sml:QQsml> 
<QQsml> 
	<%while not rs1.EOF %>
	<MGROUP name="<%=rs1("Groupname")%>">
	    <%
		Set rs2 = Server.CreateObject("ADODB.Recordset")
		rs2.ActiveConnection = conn
		rs2.Source = "SELECT * FROM v_gmeber WHERE GroupID = "&rs1("ID")&"" 
		rs2.CursorType = 0
		rs2.CursorLocation = 2
		rs2.LockType = 1
		rs2.Open()
		%>
		<%
		tempstr=""
		tempstr1=""
		while not rs2.EOF
		strUserID = ",S"&rs2("UserID")&"E"
		 %>
		 <%

		 if instr(application("OnLineUserID"),strUserID)<>false then
		  tempstr1=tempstr1&"<User name="""&rs2("state")&""" UserID="""&rs2("UserID")&""" face="""&"Images\head\"&rs2("touxiang")&".gif"&""" online="""&"1"&"""></User>"
		 else
		 tempstr=tempstr&"<User name="""&rs2("state")&""" UserID="""&rs2("UserID")&""" face="""&"Images\head\"&rs2("touxiang")&".gif"&""" online="""&"0"&"""></User>"
		 end if
		 %>
	    <%
		rs2.movenext
		wend
		response.Write(tempstr1)
		response.Write(tempstr)
		%>
		<%
		rs2.close()
		set rs2 = nothing
		%>
	</MGROUP>  
	<%
	rs1.movenext()
	wend
	%>
   </QQsml> 
 </sml:QQsml> </td>
  </tr>
  <tr> 
    <td  height="9" onmouseover='biancss(this)' onmouseout='huifuCss(this)' class='btm' onClick="window.open('fdGroup.asp','fdGroup','width=450,height=300')"> 
      <div align="center">好友与组</div>    </td>
        <td  height="9" onmouseover='biancss(this)' onmouseout='huifuCss(this)' class='btm' onClick="window.open('daochu.asp','daochu','width=300,height=150')"> 
      <div align="center">导出信息</div>    </td>
  </tr>
  <tr> 
    <td width="50%" height="9" onmouseover='biancss(this)' onmouseout='huifuCss(this)' class='btm' onClick="window.open('modifyname.asp','modifyname','width=270,height=180')"> 
      <div align="center">更改昵称</div>    </td>
    <td width="50%" height="9" onmouseover='biancss(this)' onmouseout='huifuCss(this)' class='btm' onClick="window.open('message_delete.asp','message_delete','width=530,height=380')"> 
      <div align="center">清空记录</div></td>
  </tr>
  <tr> 
    <td width="50%" height="9" onmouseover='biancss(this)' onmouseout='huifuCss(this)' class='btm' onClick="window.open('modifyPassword.asp','modifyPassword','width=270,height=180')"> 
      <div align="center">修改密码</div>    </td>
    <td width="50%" height="18" onmouseover='biancss(this)' onmouseout='huifuCss(this)' class='btm' onClick="window.location.href='exit.asp';"> 
<div align="center">安全退出</div>	</td>
</tr>
<tr style="display:none">
 <td  height="9" onmouseover='biancss(this)' onmouseout='huifuCss(this)' class='btm' onClick="unline(this)" UserID="" HasNewMsg="0" id="unknowMsgTD"> 
      <div align="center"><Img src="images/msg.gif" align="absmiddle" style="display:none" id="unknowMsg">导出信息</div></td>
</tr>
</table>
</div>
</body>
</html>
<%
rs1.close()
set rs1 = nothing
%>