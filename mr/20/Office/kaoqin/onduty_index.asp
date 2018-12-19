<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Conn/conn.asp"-->
<%
set rs=server.CreateObject("ADODB.RecordSet")
sql="select * from Tab_onduty order by ID desc"
rs.open sql,conn,1,3
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>上下班登记</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
}
.STYLE1 {font-size: 9px}
.STYLE5 {
	font-size: 10pt;
	color: #FFFFFF;
}
.STYLE6 {
	font-size: 10pt;
	color: #000000;
}
.STYLE7 {font-size: 9pt; color: #000000; }
-->
</style></head>

<body>
<table width="810" height="536" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top" background="../Images/qingjia.gif"> 
							<table width="762" height="503" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="661" height="52" valign="bottom"><table width="100" height="32" border="0">
          <tr>
            <td>&nbsp;<span class="STYLE5">上下班登记</span></td>
          </tr>
        </table></td>
        <td width="101" valign="bottom"><a href="#" onclick="Javascript:window.open('onduty_add.asp','','width=580,height=320')"><img src="../Images/waichuan5.GIF" width="80" height="21" border="0" /></a></td>
      </tr>
	 
      <tr>
        <td height="451" colspan="2" valign="top"><br><br>
			 <%
							If(rs.Eof)Then
								Response.Write("暂无登记信息!")
							Else
								
							%>
		<table width="724" height="109" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
		  <td height="39" colspan="7" background="../Images/waichubei.gif"><table width="720" height="26" border="0">

<tr>
                <td width="67"><div align="center" class="STYLE7">登记姓名</div></td>
                <td width="68"><div align="center" class="STYLE7">所属部门</div></td>
                <td width="92"><div align="center" class="STYLE7">登记类型</div></td>
                <td width="151"><div align="center" class="STYLE7">规定时间</div></td>
                <td width="122"><div align="center" class="STYLE7">登记时间</div></td>
                <td width="125"><div align="center" class="STYLE7">登记备注</div></td>
                <td width="65"><div align="center" class="STYLE7">是否登记</div></td>
                </tr>
            </table>			</td>
          </tr>
		
<%
		'分页
	  rs.pagesize=8
	  page1=clng(request("page1"))
	  if page1<1 then page1=1
	  rs.absolutepage=page1
	  for i=1 to rs.pagesize
	  k=rs("enrolremark")
	  %>
          <tr>
            <td width="71" height="31" class="STYLE7"><div align="center" class="STYLE7"><%=rs("name1")%></div></td>
            <td width="69" class="STYLE7"><div align="center" class="STYLE7"><%=rs("department")%></div></td>
            <td width="97"><div align="center" class="STYLE7"><%=rs("enroltype")%></div></td>
            <td width="156" class="STYLE7"><div align="center" class="STYLE7"><%=rs("definetime")%></div></td>
            <td width="129"><div align="center" class="STYLE7"><%=rs("enroltime")%>&nbsp;</div></td>
            <td width="129" class="STYLE7"><div align="center" class="STYLE7">
			  <% if len(k)>6 then%>
               <a href="#" onClick="Javascript:window.open('onduty_xianxi.asp?ID=<%=rs("ID")%>','','width=456,height=459')"><%=replace(left(k,6)&"....",chr(13),"<br>")%></a>
              <%else%> 
              <a href="#" onClick="Javascript:window.open('onduty_xianxi.asp?ID=<%=rs("ID")%>','','width=456,height=459')"><%=k%></a>
              <%end if %>
			</div></td>
            <td width="73"><div align="center"><%=rs("state")%></div></td>
          </tr>
	<%
	  rs.movenext
	  if rs.eof then exit for 
	  next
	  %>
          <tr>
            <td height="39" colspan="7" background="../Images/waichubei.gif"><table width="719" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td><div align="right" class="STYLE7">
                  <% if page1<>1 then%>
              <a href=<%=path1%>?page1=1>第一页</a> <a href=<%=path1%>?page1=<%=(page1-1)%>> 上一页</a>
              <%end if %>
              <%if page1<>rs.pagecount then%> 
              <a href=<%=path1%>?page1=<%=(page1+1)%>>下一页</a> <a href=<%=path1%>?page1=<%=rs.pagecount%> >最后一页</a>
              <%end if%></div></td>
              </tr>
          </table></td>
          </tr>
        </table>
		<%end if%>		</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
