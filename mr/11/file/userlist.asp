<!--#include file="Conn.asp"-->
<!--#include file="Function.asp"-->
<!--#include file="inc.asp"-->
<%Call Html("用户列表")
Dim Rs,Sql
If Session("UserName")="" Then
Response.Write("<script language=JavaScript>alert('您还没有登陆！');document.location.href='index.asp';</script>")
Response.End
End if
%>
<script language=JavaScript>
	function GetName(UserName)
	{
	var UserNameList=opener.document.form.ToUserName;
	 if(UserNameList.value=='')
	 {
	 UserNameList.value=UserName;
	 }
	 else
	 {
	 UserNameList.value=UserNameList.value+','+UserName;
	 }
	}
	</SCRIPT>
<table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""3"">
<tr align=""center"" bgcolor=""#3399ff"">
  <td width=""12%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">用户名</td>
  <td width=""6%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">性别</td>
  <td width=""20%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">单位</td>
  <td width=""12%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">部门</td>
  <td width=""16%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">电话</td>
  <td width=""10%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none""><a href=UserList.asp?orderby=group title=按组排列><font color=white>所属组</font></a></td>
  <td width=""13%"" style=""FONT-WEIGHT: bold; COLOR: white; HEIGHT: 25px; TEXT-ALIGN: center; TEXT-DECORATION: none"">操作</td>
</tr>
<%
		  dim totalPut
          dim Num,orderby,sqlStr,Gid
          dim strFileName,currentPage,MaxPerPage
		  orderby=Trim(Request("orderby"))
		  Gid=Trim(Request("GroupID"))
		   sqlStr=""
		   strFileName="UserList.asp?"
		  if orderby="group" then
		     sqlStr="order by GroupID"
			 strFileName="UserList.asp?orderby=group&"
		  end if
		  if Gid<>"" and isnumeric(Gid) then
		     sqlStr="where GroupID="&Gid&""
			 strFileName="UserList.asp?GroupID="&Gid&"&"
		  end if
          MaxPerPage=15
          if request("page")<>"" then
             currentPage=clng(request("page"))
          else
	         currentPage=1
          end if
		  If Not IsObject(Conn) Then ConnectionDatabase()
		  Set rs= Server.CreateObject("ADODB.Recordset")
		  sql="Select * from tb_Users "&sqlStr&" "
		  rs.open sql,conn,1,1
		  if  rs.eof then
			 response.write"<tr><td colspan=7>还没有任何用户……</td></tr>"
		  else
		       totalPut=rs.recordcount
	        if currentpage<1 then
   		       currentpage=1
   	        end if
   	           if (currentpage-1)*MaxPerPage>totalput then
   		       if (totalPut mod MaxPerPage)=0 then
     		      currentpage= totalPut \ MaxPerPage
	  	       else
	      	      currentpage= totalPut \ MaxPerPage + 1
   		       end if
   	           end if
		if currentPage=1 then
	    Num=0
       	do while not rs.eof
		response.Write("<tr class=bgcolor_1><td><a href=# onClick=""GetName('"&rs("UserName")&"')"" title=""选中此用户"">"&rs("UserName")&"</a></td><td align=center>")
		if rs("UserSex") then response.Write("男") else response.Write("女")
		response.Write("</td><td>"&rs("Company")&"</td><td>"&rs("Department")&"</td><td>"&rs("Tel")&"</td><td><a href=UserList.asp?GroupID="&Rs("GroupID")&" title=查看该组用户>"&GetGroupInfo(rs("GroupID"))(0)&"</a></td><td align=center>")
		response.Write("<a href=# onClick=""GetName('"&rs("UserName")&"')"">选中此用户</a></td></tr>")
       	Num=Num+1
	   	if Num>=MaxPerPage then exit do
	   	rs.movenext
	    loop
 	else
     	if (currentPage-1)*MaxPerPage<totalPut then
       	   	rs.move  (currentPage-1)*MaxPerPage
       		dim bookmark
           	bookmark=rs.bookmark
        Num=0
       	do while not rs.eof
		response.Write("<tr class=bgcolor_1><td><a href=# onClick=""GetName('"&rs("UserName")&"')"" title=""选中此用户"">"&rs("UserName")&"</a></td><td align=center>")
		if rs("UserSex") then response.Write("男") else response.Write("女")
		response.Write("</td><td>"&rs("Company")&"</td><td>"&rs("Department")&"</td><td>"&rs("Tel")&"</td><td><a href=UserList.asp?GroupID="&Rs("GroupID")&" title=查看该组用户>"&GetGroupInfo(rs("GroupID"))(0)&"</a></td><td align=center>")
		response.Write("<a href=# onClick=""GetName('"&rs("UserName")&"')"">选中此用户</a></td></tr>")
       	Num=Num+1
	   	if Num>=MaxPerPage then exit do
	   	rs.movenext
	    loop
       	else
	        currentPage=1
        Num=0
       	do while not rs.eof
		response.Write("<tr class=bgcolor_1><td><a href=# onClick=""GetName('"&rs("UserName")&"')"" title=""选中此用户"">"&rs("UserName")&"</a></td><td align=center>")
		if rs("UserSex") then response.Write("男") else response.Write("女")
		response.Write("</td><td>"&rs("Company")&"</td><td>"&rs("Department")&"</td><td>"&rs("Tel")&"</td><td><a href=UserList.asp?GroupID="&Rs("GroupID")&" title=查看该组用户>"&GetGroupInfo(rs("GroupID"))(0)&"</a></td><td align=center>")
		response.Write("<a href=# onClick=""GetName('"&rs("UserName")&"')"">选中此用户</a></td></tr>")
       	Num=Num+1
	   	if Num>=MaxPerPage then exit do
	   	rs.movenext
	    loop
          	
	    end if
	end if
	response.Write("<tr><td colspan=7>"&showpage(strFileName,totalput,MaxPerPage,CurrentPage,true,true,"个用户")&"</td></tr>")
end if
rs.close
set rs=nothing
response.Write("</table>")
Response.Write "<br><div align=center><INPUT onclick=""window.close()"" type=button value=关闭窗口 class=button></div>"
Call CloseDatabase()
%>
</body>
</html>
