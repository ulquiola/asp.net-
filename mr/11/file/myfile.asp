<!--#include file="Conn.asp"-->
<!--#include file="Function.asp"-->
<!--#include file="inc.asp"-->
<%Call Top("我的文件","MyFile.asp")
dim sql,rs
If Session("UserName")="" Then
Response.Write("<script language=JavaScript>alert('您还没有登陆！');document.location.href='Index.asp';</script>")
Response.End
End if
if Request.QueryString("Action")="del" then
If ChkPost=false Then
Response.Write("<script language=JavaScript>alert('请不要重外部提交信息！');document.location.href='index.asp';</script>")
Response.End
End If
dim id
Dim objFSO
id=Request.QueryString("id")
if id="" or not isnumeric(id) then
Response.Write("<script language=JavaScript>alert('参数错误或丢失！');document.location.href='javascript:window.history.go(-1);';</script>")
Response.End
end if
id=clng(id)
Sql="Select * From tb_Files Where id="&id&""
Set Rs=server.CreateObject("adodb.recordset")
If Not IsObject(Conn) Then ConnectionDatabase()
Rs.open Sql,Conn,1,3
If Rs.Eof Then
Rs.close
Set Rs=Nothing
CloseDatabase()
Response.Write("<script language=JavaScript>alert('文件不存在！');document.location.href='javascript:window.history.go(-1);';</script>")
response.End
Elseif Rs("UserName")=Session("UserName") Then
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
if objfso.FileExists(Server.MapPath("UpLoadFile/"&rs("FileUrl"))) then
objFSO.DeleteFile(Server.MapPath("UpLoadFile/"&rs("FileUrl")))
end if
Set objFSO =nothing
Rs.delete
Conn.Execute("Update tb_Users set UpFiles=UpFiles-1 Where UserName='"&Session("UserName")&"'")
else
rs("DelUserName")=rs("DelUserName")&","&Session("UserName")
rs.update
end if
rs.close
set rs=nothing
CloseDatabase()
Response.Write("<script language=JavaScript>alert('删除成功！');document.location.href='MyFile.asp';</script>")
End If
%>
<table width=785 align="center">
  <tbody>
    <tr>
      <td style="font-size: 12px; width: 200px; height: 21px; text-align: center"
    align=left rowspan=2 valign="top"><table class="table">
          <tbody>
            <tr>
              <td width="200" 
          colspan=3 bgcolor="#3399ff" style="font-weight: bold; width: 200px; color: white; height: 25px; text-align: center; text-decoration: none">用户登陆</td>
            </tr>
            <tr>
              <td style="width: 200px; height: 145px; text-align: center" 
          colspan=3 rowspan=2><table style="width: 200px; height: 126px; text-align: left">
                  <tbody>
                    <tr>
                      <td style="width: 188px" colspan=3 rowspan=3><%call login()%></td>
                    </tr>
                    <tr></tr>
                    <tr></tr>
                  </tbody>
                </table></td>
            </tr>
            <tr></tr>
          </tbody>
        </table></td>
      <td style="width:100%; height: 21px" align=center valign="top" colspan=2 rowspan=2><table style="width: 100%; height: 182px">
          <tbody>
            <tr>
              <td class="table" colspan=3 rowspan=3 valign="top"><%
		  response.Write("<table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""3"">"&_
               "<tr align=""center"" bgcolor=""#3399ff"">"&_
               "<td width=""25%"" style=""font-weight: bold; color: white; height: 25px; text-align: center; text-decoration: none"">文件名</td>"&_
               "<td width=""10%"" style=""font-weight: bold; color: white; height: 25px; text-align: center; text-decoration: none"">上传者</td>"&_
               "<td width=""10%"" style=""font-weight: bold; color: white; height: 25px; text-align: center; text-decoration: none"">文件大小</td>"&_
               "<td width=""20%"" style=""font-weight: bold; color: white; height: 25px; text-align: center; text-decoration: none"">上传时间</td>"&_
               "<td width=""10%"" style=""font-weight: bold; color: white; height: 25px; text-align: center; text-decoration: none"">上传ip</td>"&_
               "<td width=""6%"" style=""font-weight: bold; color: white; height: 25px; text-align: center; text-decoration: none"">操作</td>"&_
               "</tr>")
		  dim totalPut
          dim Num
          dim strFileName,currentPage,MaxPerPage
		  strFileName="MyFile.asp?"
          MaxPerPage=10
          if request("page")<>"" then
             currentPage=clng(request("page"))
          else
	         currentPage=1
          end if
		  If Not IsObject(Conn) Then ConnectionDatabase()
		  Set rs= Server.CreateObject("ADODB.Recordset")
		  sql="Select * from tb_Files where (UserName='"&Session("UserName")&"' or instr(','+GroupID+',',0)>0 or instr(','+GroupID+',',"&Session("GroupID")&")>0 or instr(','+Lcase(ToUserName)+',','"&Lcase(Session("UserName"))&"')>0) and (instr(','+DelUserName+',','"&Session("UserName")&"')=0 or isnull(DelUserName)) order by id desc "
		  rs.open sql,conn,1,1
		  if  rs.eof then
			 response.write"<tr><td colspan=6>还没有任何文件……</td></tr>"
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
		response.Write("<tr class=bgcolor_1><td><a href=Down.asp?id="&rs("id")&" target=_blank title="&htmlencode(rs("FileTitle"))&">"&getTopic(rs("FileTitle"),22)&"</a></td><td>"&rs("UserName")&"</td><td>"&myfilesize(rs("FileSize"))&"</td><td>"&rs("FileUpTime")&"</td><td>"&rs("IP")&"</td>")
		response.Write("<td align=center><a href=MyFile.asp?action=del&id="&rs("id")&" onClick=""return confirm('您确定要执行删除操作?')"">删除</a></td></tr>")
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
		response.Write("<tr class=bgcolor_1><td><a href=Down.asp?id="&rs("id")&" target=_blank title="&htmlencode(rs("FileTitle"))&">"&getTopic(rs("FileTitle"),22)&"</a></td><td>"&rs("UserName")&"</td><td>"&myfilesize(rs("FileSize"))&"</td><td>"&rs("FileUpTime")&"</td><td>"&rs("IP")&"</td>")
		response.Write("<td align=center><a href=MyFile.asp?action=del&id="&rs("id")&" onClick=""return confirm('您确定要执行删除操作?')"">删除</a></td></tr>")
       	Num=Num+1
	   	if Num>=MaxPerPage then exit do
	   	rs.movenext
	    loop
          
       	else
	        currentPage=1
        Num=0
       	do while not rs.eof
		response.Write("<tr class=bgcolor_1><td><a href=Down.asp?id="&rs("id")&" target=_blank title="&htmlencode(rs("FileTitle"))&">"&getTopic(rs("FileTitle"),22)&"</a></td><td>"&rs("UserName")&"</td><td>"&myfilesize(rs("FileSize"))&"</td><td>"&rs("FileUpTime")&"</td><td>"&rs("IP")&"</td>")
		response.Write("<td align=center><a href=MyFile.asp?action=del&id="&rs("id")&" onClick=""return confirm('您确定要执行删除操作?')"">删除</a></td></tr>")
       	Num=Num+1
	   	if Num>=MaxPerPage then exit do
	   	rs.movenext
	    loop
	    end if
	end if
	response.Write("<tr><td colspan=6>"&showpage(strFileName,totalput,MaxPerPage,CurrentPage,true,true,"个文件")&"</td></tr>")
end if
rs.close
set rs=nothing
response.Write("</table>")
		  %></td>
            </tr>
            <tr></tr>
            <tr></tr>
            <tr>
              <td class="titletd" colspan=3>&nbsp;</td>
            </tr>
          </tbody>
        </table></td>
    </tr>
    <tr></tr>
  </tbody>
</table>
<%Call Bottom()%>
