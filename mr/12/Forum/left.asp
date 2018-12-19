<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<%
set rs_Type=server.CreateObject("ADODB.RecordSet")
sql="select * from tb_Type"
rs_Type.open sql,conn,1,3
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>编程词典技术论坛</title>
<link rel="stylesheet" type="text/css" href="css/style.css">
<style type="text/css">
<!--
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
}
a:hover {
	text-decoration: none;
}
a:active {
	text-decoration: none;
}
-->
</style>
</head>
<script src="JS/fun.js"></script>
<% 
	set rs_1=server.CreateObject("adodb.recordset")
	sql1="select * from tb_User where username='"&session("username")&"'"
  	rs_1.open sql1,conn,1,3
%>
<body topmargin="0" leftmargin="0" bottommargin="0" style="overflow:auto" class="scrollbar">
<table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="24" colspan="2" valign="top" bgcolor="#FFFFFF"><table width="100%" height="24" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td background="images/back_1.gif">
		<div align="center">
		<%If Session("UserName")="" or Session("UserName")="Friend"  Then%>
		<font color="#FFFFFF">【<a href="Login.asp" target="mainwindow" class="a3">用户登录</a>】【<a href="register.asp" target="mainwindow" class="a3">注册</a>】
		</font>
	 <%Else%>
		<table width="100%" height="22" border="0" align="center" cellpadding="0" cellspacing="0">
		<tr>
		<td width="19%" valign="top"><div align="center"><img src="images/head/<%=rs_1("ICO")%>" width="20" height="20"></div></td>
		<td width="81%"> <font color="#FFFFFF">用户：<%=session("username")%>&nbsp;&nbsp;【<a href="Logout.asp" class="a3">安全退出</a>】</font></td>
		</tr>
		</table>
		<%End If%>
		</font></div></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td width="15%" valign="top" bgcolor="#B9DFFF"><table width="96%" height="1" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td></td>
        </tr>
      </table>
      <table border="0" align="right" cellpadding="0" cellspacing="0">
  <tr>
    <td width="28" height="40" style="cursor:hand" title="论坛版块" background="Images/bk1.gif"><div align="center"><img src="images/b1.gif" width="27" height="20" border="0"></div></td>
  </tr>
  <tr>
    <td width="28" height="1"></td>
  </tr>
  <tr>
    <td width="28" height="40" style="cursor:hand" title="用户信息" background="Images/bk2.gif"><div align="center"><a href="user_message.asp"><img src="images/b2.gif" width="24" height="23" border="0"></a></div></td>
  </tr>
  <tr>
    <td width="28" height="1"></td>
  </tr>
  <tr>
    <td width="28" height="40" style="cursor:hand" title="技术支持" background="Images/bk2.gif"><div align="center"><a href="skill_up.asp"><img src="images/b3.gif" width="27" height="27" border="0"></a></div></td>
  </tr>
  <tr>
    <td width="28" height="1"></td>
  </tr>
  <tr>
    <td width="28" height="40" style="cursor:hand" title="使用帮助" background="Images/bk2.gif"><div align="right"><a href="help.asp"><img src="images/b4.gif" width="23" height="24" border="0"></a></div></td>
  </tr>
</table>    </td>
    <td width="85%" valign="top" bgcolor="#FFFFFF"><table width="96%" height="2" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td></td>
      </tr>
    </table>
    <table width="96%" height="17" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td background="images/bbs_top.gif"><div align="center">编程词典技术论坛</div></td>
      </tr>
    </table>

<table width="90%" height="5" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td></td>
  </tr>
</table>
    <table width="88%" height="16" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="3%"><div align="center"><img src="images/bbs_biao.gif" width="16" height="14"></div></td>
        <td width="97%">&nbsp;论坛讨论区</td>
      </tr>
    </table>
    
	  <table width="89%"   border="0" style="margin:0px;" align="center" cellpadding="0" cellspacing="0">
            <%'rs_Type.movefirst()
			if rs_type.eof and rs_type.bof then%>
            <tr>
              <td colspan="2" align="center">暂无讨论区信息!</td>
            </tr>
            <%
			else
			m=1
			 for i=1 to rs_type.recordcount
				set rs_L=server.CreateObject("ADODB.RecordSet")
				sql_3="select P.*,T.typename as dd from tb_smalltype P inner join tb_Type T on P.typeid= T.id where P.typeid="&rs_Type("id")
				rs_L.open sql_3,conn,1,3
			 %>
            <tr>
			<td  height="15" style=" padding:0;float:left">
					<%
					If(rs_L.eof and rs_L.bof)Then
					%><img src="images/jian_null1.gif" width="38" height="18" align="absbottom">&nbsp;<%=server.HTMLEncode(rs_Type("typename"))%>
					<%
					Else
					%>
					<%IF i=rs_type.recordcount then%><a href="Javascript:ShowTR(img<%=m%>,OpenRep<%=m%>,<%=rs_Type("id")%>)"><img src="Images/uaplus.gif" name="img<%=m%>" width="38" height="18" border="0" id="img<%=m%>"></a>
					<%else%>
					<a href="Javascript:ShowTR(img<%=m%>,OpenRep<%=m%>,<%=rs_Type("id")%>)"><img src="Images/jia1.gif" alt="展开" name="img<%=m%>"  width="38" height="18" border="0" align="absbottom" id="img<%=m%>"></a>
					<%end if%>
			  <a href="Javascript:ShowTR(img<%=m%>,OpenRep<%=m%>,<%=rs_Type("id")%>)"><%=server.HTMLEncode(rs_Type("typename"))%></a>
	          <%
					End If
					%></td>
				<%if rs_L.recordcount>0 then%>
	    <tr id="OpenRep<%=m%>" style="display:none;" height="15px">
			   <td height="15px" colspan="7" valign="top">	
				<% for j=1 to rs_L.recordcount%>
			     <table width="92%" border="0" cellpadding="-2" cellspacing="-2">
                   <tr onMouseOver="this.style.background='#EEEEEE'" onMouseOut="this.style.background=''">
                      <td align="center" valign="top" height="15px">
					 <table border="0" align="left" cellpadding="0" cellspacing="0">
                  <tr>
				  
                    <td width="14px" height="15px" align="right" valign="top"><img src="Images/line.gif" width="11" height="19" align="baseline"></td>
                    <td align="center" height="15px" width="50px" valign="top">
                      <%IF J=rs_L.recordcount then%>    
                      <img src="Images/upangle.gif" width="16" height="16" align="top">
                      <%else%>      
                        <img src="Images/tshaped.gif" width="16" height="16" border="0" align="absbottom">
                      <%end if%>                        <img src="images/11.gif" width="16" height="16" border="0">
                    </td>
                    <td height="15"><a href="bbs_list.asp?typeid=<%=rs_L("id")%>" target="mainwindow" title="<%=rs_L("smalltypename")%>" class="a1"><%=rs_L("smalltypename")%></a></td>
                  </tr>
              </table></td>
                   </tr>
                 </table>
		         <%
				m=m+1  '注意，该条语句一定不能少
			  	rs_L.movenext
				next
				%></td> 
			   <%end if%>
	   </tr>
            <%
			 rs_Type.movenext
			 If rs_Type.Eof Then exit For 
			 next
			 end if
			 %>
<script language="javascript">
function ShowTR(objImg,objTr,id){
	if(objTr.style.display == "block"){
		objTr.style.display = "none";
		objImg.src = "Images/jia1.gif";
		objImg.alt = "展开";
		parent.mainwindow.location ="main.asp?typeid="+id;
		
	}else{
		objTr.style.display = "block";
		objImg.src = "Images/jian1.gif";
		objImg.alt = "折叠";
		parent.mainwindow.location ="main.asp?typeid="+id;
	
	}
}
</script>
      </table>
  </td>
  </tr>
</table>
</body>
</html>