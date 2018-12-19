<!--#include file="conn/conn.asp"-->
<%
Set rs_Type=Server.CreateObject("ADODB.RecordSet")
sql="select * from tab_Menu  where NodeID=0"
rs_Type.open sql,conn,1,3
%>
<title>树状导航菜单</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.STYLE2 {
	font-size: 9pt;
	color: #000000;
}
body {
	margin-top: 0px;
	margin-bottom: 0px;
}
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
<table width="100%" height="25"  border="0" cellpadding="0" cellspacing="0">
            <%
			rs_Type.movefirst()
			if rs_type.eof and rs_type.bof then
			%>
            <tr>
              <td align="center"><span class="STYLE2">暂无功能分类信息!</span></td>
            </tr>
            <%
			else
			m=1
			 for i=1 to rs_type.recordcount
				set rs_L=server.CreateObject("ADODB.RecordSet")
        	    sql="select * from tab_Menu where NodeID="&rs_type("ID")
				rs_L.open sql,conn,1,3
			 %>
            <tr>
			<td>
					<%
					If(rs_L.eof and rs_L.bof)Then
					%>
					&nbsp;<img src="Images/jian_null.gif" width="38" height="16" border="0">
					<%=server.HTMLEncode(rs_Type("MenuName"))%>
					<%Else%>
					&nbsp;<a href="Javascript:ShowTR(img<%=m%>,OpenRep<%=m%>)"><img src="Images/jia.gif" border="0" alt="展开" id="img<%=m%>"></a>
					<a href="Javascript:ShowTR(img<%=m%>,OpenRep<%=m%>)"><span class="STYLE2"><%=server.HTMLEncode(rs_Type("MenuName"))%></span></a>
					<%
					End If
					%>
                </td>
				<%
				if rs_L.recordcount>0 then
				%>
			  <tr id="OpenRep<%=m%>" style="display:none;">
			   <td colspan="6">				
				<% for j=1 to rs_L.recordcount%>
			     <table width="94%"  border="0" cellspacing="0" cellpadding="0">
                   <tr onMouseOver="this.style.background='#96F7F4'" onMouseOut="this.style.background=''">
                      <td height="25" align="center"><table width="90%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="7%" align="left">&nbsp;&nbsp;&nbsp;<img src="images/folder.gif" width="16" height="16" border="0"></td>
                    <td width="93%" align="left">&nbsp;<a href="<%=rs_L("LinkUrl")%>" target="mainFrame"><span class="STYLE2"><%=rs_L("MenuName")%></span></a></td>
                  </tr>
              </table></td>
                   </tr>
                 </table>
				 <%	
				    m=m+1  '注意，该条语句一定不能少
			  		rs_L.movenext
				next
				%>
			   </td>
			   <%end if%>
			  </tr>
<%
			 rs_Type.movenext
			 If rs_Type.Eof Then exit For 
			 next
			 end if
%>
<script language="javascript">
ShowTR(img1,OpenRep1)  //设置第1个结点为展开状态
function ShowTR(objImg,objTr)
{
	if(objTr.style.display == "block")
	{
		objTr.style.display = "none";
		objImg.src = "Images/jia.gif";
		objImg.alt = "展开";		
	}
	else
	{
		objTr.style.display = "block";
		objImg.src = "Images/jian.gif";
		objImg.alt = "折叠";		
	}
}
</script>
</tr>
</table>