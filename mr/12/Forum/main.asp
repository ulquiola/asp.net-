<!--#include file="Conn/conn.asp" -->
<%
If request.QueryString("typeid")<>"" Then
	session("id")=request.QueryString("typeid")
	set rs=server.CreateObject("ADODB.RecordSet")
	sql="select * from view_Type where ID="&session("id")
	rs.open sql,conn,1,3
	set rs1=server.CreateObject("ADODB.RecordSet")
	sql1="select * from view_smalltype where typeid="&session("id")
	rs1.open sql1,conn,1,3
	session("smalltypename")=rs1("smalltypename")
else
	response.write("&nbsp;")
end if
%>
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
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="css/style.css">
<!--#include file="top.asp"-->
<table width="96%" height="60" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#58BAE9">
  <tr>
    <td bgcolor="#FFFFFF" width="564" height="89"><img src="images/banner_l.gif" width="564" height="89" border="0"></td>
    <td width="100%" background="images/banner_c.gif" bgcolor="#FFFFFF">&nbsp;</td>
    <td bgcolor="#FFFFFF" width="19" height="89"><img src="images/banner_r.gif" width="19" height="89" /></td>
  </tr>
</table>
<table width="96%" height="8" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td bgcolor="#FFFFFF"></td>
  </tr>
</table>

<table width="96%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#349ED8">
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
      <tr>
        <td height="22" bgcolor="#58BAE9" >
		<% If request.QueryString("typeid")<>"" Then %> 
		<table width="100%" height="22" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="5%"><div align="center"><img src="images/bbs_biao.gif" width="16" height="14"></div></td>
            <td width="47%"><font color="#FFFFFF"><%=rs("tname")%></font></td>
            <td width="16%"><font color="#FFFFFF">创建时间：<%=FormatDateTime(rs("CreateTime"),2)%></font></td>
            <td width="4%"><div align="center"></div></td>
            <td width="15%"><font color="#FFFFFF">总版主：<%=rs("owner")%></font></td>
          </tr>
        </table>
		<% end if %>
		</td>
        </tr>
      <tr>
        <td height="22" >
		
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
          <tr>
            <td height="16" colspan="3" valign="bottom" bgcolor="#CCCCCC">&nbsp;&nbsp;版块名称</td>
            <td width="15%" bgcolor="#CCCCCC" valign="bottom"><div align="center">创建时间</div></td>
            <td width="15%" bgcolor="#CCCCCC" valign="bottom"><div align="center">版主</div></td>
            <td width="13%" bgcolor="#CCCCCC" valign="bottom"><div align="center">帖子统计</div></td>
          </tr>
		  
<% If request.QueryString("typeid")="" Then %> 
<table width="100%" height="93"  border="0" cellpadding="-2" cellspacing="-2"> 
 <tr><td height="93" align="center">很报歉！暂无任何版块信息！</td> 
</tr> 
</table>
 <%
			  Else
			rs1.pagesize=10
			page=CLng(Request("page"))
			if page<1 then page=1
			rs1.absolutepage=page
			for i=1 to rs1.pagesize
			if cbool(i mod 1) Then
%> 
<%Else%>		  
 <%
'获取主题总数
Set Rs_count=Server.CreateObject("ADODB.RecordSet")
SQL="Select count(*) as total From tb_Topic where TypeID="&rs1("ID")
Rs_count.Open SQL,conn,1,3
'获取今日主题数
Set Rs_Today=Server.CreateObject("ADODB.RecordSet")
SQL="Select count(*) as total_Today From tb_Topic where TypeID="&rs1("ID")&" and createTime>=#"&date()&"# and createTime < #"&date()+1&"#"
Rs_Today.Open SQL,conn,1,3
%>
		  <tr>
		  
            <td width="4%" height="50" rowspan="2"><div align="center"><img src="images/bbs_mark2.gif" width="22" height="22"></div></td>
            <td width="46%" rowspan="2" style="line-height:2">&nbsp;&nbsp;&nbsp;<a href="bbs_list.asp?typeid=<%=rs1("ID")%>">『<%=rs1("smalltypename")%>』</a><br>
              &nbsp;&nbsp;&nbsp;<font color="#999999"><%=rs1("account")%></font></td>
            <td width="7%" rowspan="2">&nbsp;</td>
            <td height="40" rowspan="2"><div align="center"><%=FormatDateTime(rs1("createobject"),2)%></div></td>
            <td height="40" rowspan="2"><div align="center"><%=rs1("username")%></div></td>
            <td height="19"><span class="word_grey">&nbsp;主题总数：<%=rs_count("total")%></span> </td>
          </tr>
		 
		  <tr>
		    <td height="20"><span class="word_grey">&nbsp;今日主题数：<%=rs_Today("total_Today")%></span></td>
		    </tr>
	<%      
	 		 End If
			  rs1.MoveNext
			  If rs1.Eof Then exit For 
			  Next
			  End If
			  %>
        </table>

		</td>
        </tr>
    </table>
</table> 
<%If request.QueryString("typeid")<>""  Then %> 
  <table width="96%" height="25"  border="0" align="center" cellpadding="-2" cellspacing="-2" bgcolor="FAFAFA"> 
    <tr> 
      <td valign="bottom">
	   <% if page<>1 then %> 
        <a href=?page=1 class="white">第一页</a> <a href=?page=<%=(page-1)%> class="white">上一页</a> 
        <%end if 
			if page<>rs1.pagecount then %> 
        <a href=?page=<%=(page+1)%> class="white">下一页</a> <a href=?page=<%=rs1.pagecount%> class="white">最后一页</a> 
        <%end if%> </td> 
      <%If not(rs1.Eof and rs1.Bof) Then%> 
      <td width="29%" align="right" class="word_grey">[<%=page%>/<%=rs1.PageCount%>]&nbsp;&nbsp;每页<span class="STYLE1"><%=rs1.PageSize%></span>条&nbsp;&nbsp;共<%=rs1.RecordCount%>条主题信息</td> 
      <%End If   
	%> 
    </tr> 
</table>
<%
rs1.close()
Set rs1=Nothing
%>
<% end if %>
	</td>
  </tr>
</table>
<table width="96%" height="8" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td bgcolor="#FFFFFF"></td>
  </tr>
</table>
<!--#include file="bottom.asp" -->
