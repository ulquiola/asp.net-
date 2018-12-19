<!--#include file=drawimg.asp-->
<%
dim etdbvote'调查文件数据库链接地址
dim startDay,endDay
dim total
dim hourTotal
dim t,v,imgUrl,strSearchName,systemName,browserName,screenName,pageUrl

startDay=request("startday")
endDay=request("endday")
strSearchName=request("searchName")
systemName=request("systemName")
browserName=request("browserName")
screenName=request("screenName")
pageUrl=request("pageUrl")


'etdbvote=server.mapPath("manage/admin/vote/vote.mdb")

set conn=server.CreateObject("adodb.connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& server.MapPath("data/Stat.mdb")
conn.open connstr

sub UserOnLine()
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from online order by id desc"
	rs.Open sql,conn,1,1

	response.write "<table width=800 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
	response.write "  <tr>"
	response.write "<td background=""images/bk.gif"" width=200 height=""22"">日期</td>"
	response.write "<td background=""images/bk.gif"" width=200>来源页面</td>"
	response.write "<td background=""images/bk.gif"" width=300>访问页面</td>"
	response.write "</tr>"
	if not rs.eof then
		do while not rs.eof 
			response.write "<tr>"
			response.write "<td bgcolor=#ffffff>"
			response.write rs("visitDate")&" "&rs("visitTime")
			response.write "</td>"

			response.write "<td bgcolor=#ffffff>"
			response.write "<a href="""&rs("lastpage")&""" target=_blank>"&rs("lastpage")&"</a>"
			response.write "</td>"
			
			response.write "<td bgcolor=#ffffff>"
			response.write "<a href="""&rs("pageurl")&""" target=_blank>"&rs("pageurl")&"</a>"
			response.write "</td>"
			
			
			response.write "</tr>"

			
			rs.movenext
		loop

		response.write "</table>"
	else
		response.write "<tr><td  colspan=""4"">"
		response.write "暂无信息"
		response.write "</td></tr>"
		response.write "</table>"
	end if
end sub


'受访页历史
sub PageComeHistory()

	if startDay<>"" and endDay<>"" and pageUrl<>"" then

		dim sql
		response.Write "<div class=""Title"">来源页面:<a href="""&pageUrl&""">"&pageUrl&"</a> </div>"
		response.Write "<div style=""height:20px;padding-top:5px"">时间段: <font color=red>"&request("startday")&"</font> 到 <font color=red>"&request("endday")&"</font> 的统计 </div>"
		response.Write "<iframe src="""" id=""ifrImg"" name=""ifrImg"" width=""800"" height=""400"" frameborder=""0"" scrolling=""no""></iframe>"
		GetPageUrlTotal(pageUrl)
		sql="select * from ("
		for i=cdate(startDay) to cdate(endDay)
			sql=sql+"SELECT count(*) as c,#"&i&"# as visitDate FROM visitor where visitdate=#"&i&"# and pageUrl='"&pageUrl&"' "
			t=t+cstr(day(i))
			if i<>cdate(endDay) then
				sql=sql+"union "
				t=t+","
			end if
		next
		sql=sql+") order by visitDate"

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql,conn,1,1

		response.write "<table width=800 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">日期</td>"
		response.write "<td background=""images/bk.gif"" width=200>点击次数(PV)</td>"
		response.write "<td background=""images/bk.gif"" width=300>比例图</td>"
		response.write "</tr>"
		if not rs.eof then
			do while not rs.eof 
				response.write "<tr>"
				response.write "<td bgcolor=#ffffff>"
				response.write rs("visitDate")
				response.write "</td>"

				response.write "<td bgcolor=#ffffff>"
				response.write rs("c")
				response.write "</td>"
				
				response.write "<td bgcolor=#ffffff>"
				if total=0 then
					response.Write "0%"
				else
					response.write "<img src=""images/bar2.gif"" width="&rs("c")/total*100&" height=""10""/> "
					response.write FormatNumber(rs("c")/total,4)*100&"%"
				end if
				response.write "</td>"
				
				
				response.write "</tr>"
				v=v&rs("c")&","
				
				rs.movenext
			loop
			if v<>"" then
				v=left(v,(len(v)-1))
			end if
			rs.close
			response.write "</table>"
		else
			response.write "<tr><td  colspan=""4"">"
			response.write "暂无信息"
			response.write "</td></tr>"
			response.write "</table>"
		end if
		if total<>"0" then 
			imgUrl="img.asp?type=zhuxing&d="&t&"&v="&v
			response.write "<script>ifrImg.location.href='"&imgUrl&"';</script>"
		else
			response.write "<script>ifrImg.document.write(""暂无图形"");</script>"
		end if
	end if

end sub

'受访页面分析
sub PageCome()
'分页代码第一段－－－－－－－－－－－－－－－
	if startday<>"" and endday<>"" then
		response.Write "<div class=""Title"">受访页面分析: </div>"
		response.Write "<div style=""height:20px;padding-top:5px"">时间段: <font color=red>"&request("startday")&"</font> 到 <font color=red>"&request("endday")&"</font> 的统计 </div>"
		GetTotal()
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql = "select count(*) as c,pageUrl from  visitor where  visitDate>=#"&startday&"# and visitDate<=#"&endday&"# group by pageUrl order by pageUrl"

		rs.Open sql,conn,1,1
		
		if not rs.eof then
			dim currentpage
			dim totalput,f
			const maxperpage=20
			rs.movefirst
			if (trim(request.querystring("page")))<>"" then
			currentpage=clng(request.querystring("page"))
			else
			currentpage=1
			end if
			totalput=rs.recordcount
			if (currentpage<>1) then
				if (currentpage-1)*maxperpage<totalput then
				rs.move(currentpage-1)*maxperpage
				end if
			end if
			dim n,k
			if (totalput mod maxperpage)=0 then
			n=totalput\maxperpage
			else
			n=totalput\maxperpage+1
			end if
			
			response.write "<table width=800 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
			response.write "  <tr>"
			response.write "<td background=""images/bk.gif"" width=200 height=""22"">访问页</td>"
			response.write "<td background=""images/bk.gif"" width=200>点击次数(PV)</td>"
			response.write "<td background=""images/bk.gif"" width=300>比例图</td>"
			response.write "<td background=""images/bk.gif"" width=300>历史</td>"
			response.write "</tr>"
		'－－－－－－－－－－－－－－－分页代码第一段
		'循环数据段－－－－－－－－－
				
				for f=1 to MaxPerPage
					if not rs.eof then
	
					response.write "<tr>"
					response.write "<td bgcolor=#ffffff>"
					response.write rs("pageUrl")
					response.write "</td>"
	
					response.write "<td bgcolor=#ffffff>"
					response.write rs("c")
					response.write "</td>"
					
					response.write "<td bgcolor=#ffffff>"
					if total=0 then
						response.Write "0%"
					else
						response.write "<img src=""images/bar2.gif"" width="&rs("c")/total*100&" height=""10""/> "
						response.write FormatNumber(rs("c")/total,4)*100&"%"
					end if
					response.write "</td>"
					
					response.write "<td bgcolor=#ffffff>"
					response.write "<a href=""pageComeHistory.asp?startDay="&startDay&"&endday="&endDay&"&pageUrl="
					if IsNull(rs("pageUrl"))=false then
						response.Write server.urlencode(rs("pageUrl"))
					end if
					response.Write """>历史</a>"
					response.write "</td>"
					
					response.write "</tr>"
	
					rs.movenext
					end if
				next
				response.write "</table>"
			rs.close
		'－－－－－－－－循环数据段
		
		'分页代码第二段－－－－－－－－－
				response.write "<table width='95%' >"
				response.write "<tr><td colspan=3 align=right>"
				k=currentpage
				response.write "共"&totalput&"条&nbsp;&nbsp;"
				response.write "共"&n&"页&nbsp;&nbsp;"
				response.write "当前页:"&currentpage&"&nbsp;&nbsp;"
				if k<>1 then
				response.write "共"&maxperpage&"页&nbsp;&nbsp;"
				response.write "当前页:"&currentpage&"&nbsp;&nbsp;"
				response.write "<a href='?startday="&startday&"&endday="&endday&"&page=1'>首页</a>&nbsp;&nbsp;"
				response.write "<a href='?startday="&startday&"&endday="&endday&"&page="&k-1&"'>上一页</a>&nbsp;&nbsp;"
				else
				response.write "首页&nbsp;&nbsp;上一页&nbsp;&nbsp;"
				end if
				if k<>n then
				response.write "<a href='?startday="&startday&"&endday="&endday&"&page="&k+1&"'>下一页</a>&nbsp;&nbsp;"
				response.write "<a href='?startday="&startday&"&endday="&endday&"&page="&n&"'>尾页</a>"
				else
				response.write "下一页&nbsp;&nbsp;尾页"
				end if
				response.write "</td></tr></table>"
		else
			response.write ""
		end if
	end if
'－－－－－－－－分页代码第二段
end sub


'分辨率历史
sub ScreenrHistory()
	if startDay<>"" and endDay<>"" and screenName<>"" then
	
		dim sql
		response.Write "<div class=""Title"">分辨率:"&screenName&" </div>"
		response.Write "<div style=""height:20px;padding-top:5px"">时间段: <font color=red>"&request("startday")&"</font> 到 <font color=red>"&request("endday")&"</font> 的统计 </div>"
		response.Write "<iframe src="""" id=""ifrImg"" name=""ifrImg"" width=""800"" height=""400"" frameborder=""0"" scrolling=""no""></iframe>"
		GetScreenTotal(screenName)
		sql="select * from ("
		for i=cdate(startDay) to cdate(endDay)
			sql=sql+"SELECT count(*) as c,#"&i&"# as visitDate FROM visitor where visitdate=#"&i&"# and screen='"&screenName&"' "
			t=t+cstr(day(i))
			if i<>cdate(endDay) then
				sql=sql+"union "
				t=t+","
			end if
		next
		sql=sql+") order by visitDate"

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql,conn,1,1

		response.write "<table width=800 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">日期</td>"
		response.write "<td background=""images/bk.gif"" width=200>点击次数(PV)</td>"
		response.write "<td background=""images/bk.gif"" width=300>比例图</td>"
		response.write "</tr>"
		if not rs.eof then
			do while not rs.eof 
				response.write "<tr>"
				response.write "<td bgcolor=#ffffff>"
				response.write rs("visitDate")
				response.write "</td>"

				response.write "<td bgcolor=#ffffff>"
				response.write rs("c")
				response.write "</td>"
				
				response.write "<td bgcolor=#ffffff>"
				if total=0 then
					response.Write "0%"
				else
					response.write "<img src=""images/bar2.gif"" width="&rs("c")/total*100&" height=""10""/> "
					response.write FormatNumber(rs("c")/total,4)*100&"%"
				end if
				response.write "</td>"
				
				
				response.write "</tr>"
				v=v&rs("c")&","
				
				rs.movenext
			loop
			if v<>"" then
				v=left(v,(len(v)-1))
			end if
			rs.close
			response.write "</table>"
		else
			response.write "<tr><td  colspan=""4"">"
			response.write "暂无信息"
			response.write "</td></tr>"
			response.write "</table>"
		end if
		if total<>"0" then 
			imgUrl="img.asp?type=zhuxing&d="&t&"&v="&v
			response.write "<script>ifrImg.location.href='"&imgUrl&"';</script>"
		else
			response.write "<script>ifrImg.document.write(""暂无图形"");</script>"
		end if
	end if

end sub

'浏览器历史
sub BrowserHistory()
	if startDay<>"" and endDay<>"" and browserName<>"" then
	
		dim sql
		response.Write "<div class=""Title"">浏览器名称:"&browserName&" </div>"
		response.Write "<div style=""height:20px;padding-top:5px"">时间段: <font color=red>"&request("startday")&"</font> 到 <font color=red>"&request("endday")&"</font> 的统计 </div>"
		response.Write "<iframe src="""" id=""ifrImg"" name=""ifrImg"" width=""800"" height=""400"" frameborder=""0"" scrolling=""no""></iframe>"
		GetBrowserTotal(browserName)
		sql="select * from ("
		for i=cdate(startDay) to cdate(endDay)
			sql=sql+"SELECT count(*) as c,#"&i&"# as visitDate FROM visitor where visitdate=#"&i&"# and browser='"&browserName&"' "
			t=t+cstr(day(i))
			if i<>cdate(endDay) then
				sql=sql+"union "
				t=t+","
			end if
		next
		sql=sql+") order by visitDate"

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql,conn,1,1

		response.write "<table width=800 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">日期</td>"
		response.write "<td background=""images/bk.gif"" width=200>点击次数(PV)</td>"
		response.write "<td background=""images/bk.gif"" width=300>比例图</td>"
		response.write "</tr>"
		if not rs.eof then
			do while not rs.eof 
				response.write "<tr>"
				response.write "<td bgcolor=#ffffff>"
				response.write rs("visitDate")
				response.write "</td>"

				response.write "<td bgcolor=#ffffff>"
				response.write rs("c")
				response.write "</td>"
				
				response.write "<td bgcolor=#ffffff>"
				if total=0 then
					response.Write "0%"
				else
					response.write "<img src=""images/bar2.gif"" width="&rs("c")/total*100&" height=""10""/> "
					response.write FormatNumber(rs("c")/total,4)*100&"%"
				end if
				response.write "</td>"
				
				
				response.write "</tr>"
				v=v&rs("c")&","
				
				rs.movenext
			loop
			if v<>"" then
				v=left(v,(len(v)-1))
			end if
			rs.close
			response.write "</table>"
		else
			response.write "<tr><td  colspan=""4"">"
			response.write "暂无信息"
			response.write "</td></tr>"
			response.write "</table>"
		end if
		if total<>"0" then 
			imgUrl="img.asp?type=zhuxing&d="&t&"&v="&v
			response.write "<script>ifrImg.location.href='"&imgUrl&"';</script>"
		else
			response.write "<script>ifrImg.document.write(""暂无图形"");</script>"
		end if
	end if

end sub

'搜索引擎历史访问
sub OpSystemHistory()
	if startDay<>"" and endDay<>"" and systemName<>"" then
	
		dim sql
		response.Write "<div class=""Title"">操作系统:"&systemName&" </div>"
		response.Write "<div style=""height:20px;padding-top:5px"">时间段: <font color=red>"&request("startday")&"</font> 到 <font color=red>"&request("endday")&"</font> 的统计 </div>"
		response.Write "<iframe src="""" id=""ifrImg"" name=""ifrImg"" width=""800"" height=""400"" frameborder=""0"" scrolling=""no""></iframe>"
		GetSystemTotal(systemName)
		sql="select * from ("
		for i=cdate(startDay) to cdate(endDay)
			sql=sql+"SELECT count(*) as c,#"&i&"# as visitDate FROM visitor where visitdate=#"&i&"# and OpSystem='"&systemName&"' "
			t=t+cstr(day(i))
			if i<>cdate(endDay) then
				sql=sql+"union "
				t=t+","
			end if
		next
		sql=sql+") order by visitDate"

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql,conn,1,1

		response.write "<table width=800 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">日期</td>"
		response.write "<td background=""images/bk.gif"" width=200>点击次数(PV)</td>"
		response.write "<td background=""images/bk.gif"" width=300>比例图</td>"
		response.write "</tr>"
		if not rs.eof then
			do while not rs.eof 
				response.write "<tr>"
				response.write "<td bgcolor=#ffffff>"
				response.write rs("visitDate")
				response.write "</td>"

				response.write "<td bgcolor=#ffffff>"
				response.write rs("c")
				response.write "</td>"
				
				response.write "<td bgcolor=#ffffff>"
				if total=0 then
					response.Write "0%"
				else
					response.write "<img src=""images/bar2.gif"" width="&rs("c")/total*100&" height=""10""/> "
					response.write FormatNumber(rs("c")/total,4)*100&"%"
				end if
				response.write "</td>"
				
				
				response.write "</tr>"
				v=v&rs("c")&","
				
				rs.movenext
			loop
			if v<>"" then
				v=left(v,(len(v)-1))
			end if
			rs.close
			response.write "</table>"
		else
			response.write "<tr><td  colspan=""4"">"
			response.write "暂无信息"
			response.write "</td></tr>"
			response.write "</table>"
		end if
		if total<>"0" then 
			imgUrl="img.asp?type=zhuxing&d="&t&"&v="&v
			response.write "<script>ifrImg.location.href='"&imgUrl&"';</script>"
		else
			response.write "<script>ifrImg.document.write(""暂无图形"");</script>"
		end if
	end if

end sub

'访客分辨率
sub ScreenApart()
	if startDay<>"" and endDay<>"" then

		response.Write "<div class=""Title"">分辨率分布 </div>"
		GetTotal()
		set rs=conn.execute("select count(id) as c,screen from visitor where  visitDate>=#"&startday&"# and visitDate<=#"&endday&"#  group by screen order by screen")
		response.write "<table width=800 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">分辨率</td>"
		response.write "<td background=""images/bk.gif"" width=300>所占比例(来访次数)</td>"
		response.write "<td background=""images/bk.gif"" width=100>历史</td>"

		response.write "</tr>"
		if not rs.eof then
			do while not rs.eof
				response.write "<tr>"
				response.write "<td bgcolor=#ffffff>"
				response.write rs("screen")
				response.write "</td>"

				response.write "<td bgcolor=#ffffff>"
				if total=0 then
					response.Write "0% (0)"
				else
					response.write "<img src=""images/bar2.gif"" width="&rs("c")/total*100&" height=""10""/> "
					response.write FormatNumber(rs("c")/total,4)*100&"% ("&rs("c")&")"
				end if
				response.write "</td>"

				
				response.write "<td bgcolor=#ffffff>"
				response.write "<a href=""screenHistory.asp?startDay="&startDay&"&endday="&endDay&"&screenName="&rs("screen")&""">历史</a>"
				response.write "</td>"
				
				
				response.write "</tr>"
				
				
				rs.movenext
			loop
			rs.close
			response.write "</table>"
		else
			response.write "<tr><td  colspan=""4"">"
			response.write "暂无信息"
			response.write "</td></tr>"
			response.write "</table>"
		end if
	end if
end sub
'访客浏览器
sub OpBrowser()
	if startDay<>"" and endDay<>"" then

		response.Write "<div class=""Title"">浏览器分布 </div>"
		GetTotal()
		set rs=conn.execute("select count(id) as c,Browser from visitor where visitDate>=#"&startday&"# and visitDate<=#"&endday&"#  group by Browser order by Browser")
		response.write "<table width=800 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">浏览器名称</td>"
		response.write "<td background=""images/bk.gif"" width=300>所占比例(来访次数)</td>"
		response.write "<td background=""images/bk.gif"" width=100>历史</td>"

		response.write "</tr>"
		if not rs.eof then
			do while not rs.eof
				response.write "<tr>"
				response.write "<td bgcolor=#ffffff>"
				response.write rs("Browser")
				response.write "</td>"

				response.write "<td bgcolor=#ffffff>"
				if total=0 then
					response.Write "0% (0)"
				else
					response.write "<img src=""images/bar2.gif"" width="&rs("c")/total*100&" height=""10""/> "
					response.write FormatNumber(rs("c")/total,4)*100&"% ("&rs("c")&")"
				end if
				response.write "</td>"

				
				response.write "<td bgcolor=#ffffff>"
				response.write "<a href=""BrowserHistory.asp?startDay="&startDay&"&endday="&endDay&"&browserName="
				if IsNull(rs("Browser"))=false then
					response.write replace(rs("Browser")," ","+")
				end if
				response.write """>历史</a>"
				response.write "</td>"
				
				
				response.write "</tr>"
				
				
				rs.movenext
			loop
			rs.close
			response.write "</table>"
		else
			response.write "<tr><td  colspan=""4"">"
			response.write "暂无信息"
			response.write "</td></tr>"
			response.write "</table>"
		end if
	end if
end sub


'访客操作系统
sub OpSystem()
	if startDay<>"" and endDay<>"" then
	response.Write "<div style=""height:20px;padding-top:5px"">客户端明细 - 从 <font color=red>"&request("startday")&"</font> 到 <font color=red>"&request("endday")&"</font> 的统计 </div>"
		response.Write "<div class=""Title"">访客操作系统 </div>"
		GetTotal()
		set rs=conn.execute("select count(id) as c,opsystem from visitor where  visitDate>=#"&startday&"# and visitDate<=#"&endday&"#  group by opsystem order by opsystem")
		response.write "<table width=800 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">操作系统名称</td>"
		response.write "<td background=""images/bk.gif"" width=300>所占比例(来访次数)</td>"
		response.write "<td background=""images/bk.gif"" width=100>历史</td>"

		response.write "</tr>"
		if not rs.eof then
			do while not rs.eof
				response.write "<tr>"
				response.write "<td bgcolor=#ffffff>"
				response.write rs("opsystem")
				response.write "</td>"

				response.write "<td bgcolor=#ffffff>"
				if total=0 then
					response.Write "0% (0)"
				else
					response.write "<img src=""images/bar2.gif"" width="&rs("c")/total*100&" height=""10""/> "
					response.write FormatNumber(rs("c")/total,4)*100&"% ("&rs("c")&")"
				end if
				response.write "</td>"

				
				response.write "<td bgcolor=#ffffff>"
				response.write "<a href=""systemHistory.asp?startDay="&startDay&"&endday="&endDay&"&systemName="
				if IsNull(rs("opsystem"))=false then
					response.write replace(rs("opsystem")," ","+")
				end if
				response.write """>历史</a>"
				response.write "</td>"
				
				
				response.write "</tr>"
				
				
				rs.movenext
			loop
			rs.close
			response.write "</table>"
		else
			response.write "<tr><td  colspan=""4"">"
			response.write "暂无信息"
			response.write "</td></tr>"
			response.write "</table>"
		end if
	end if
end sub

'搜索引擎
sub SearchCome()
	if startDay<>"" and endDay<>"" then
		GetSearchTotal()
		response.Write "<div style=""height:20px;padding-top:5px"">搜索引擎来访分布 - 从 <font color=red>"&request("startday")&"</font> 到 <font color=red>"&request("endday")&"</font> 的统计 </div>"
		response.Write "<div class=""Title"">总计从搜索引擎来访："&total&" 次(PV) </div>"
		response.Write "<iframe src="""" id=""ifrImg"" name=""ifrImg"" width=""800"" height=""400"" frameborder=""0"" scrolling=""no""></iframe>"
		set rs=conn.execute("select count(id) as id,searchName from visitor where isSearch=0 and  visitDate>=#"&startday&"# and visitDate<=#"&endday&"#  group by searchName order by searchName")
		response.write "<table width=800 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">搜索引擎名称</td>"
		response.write "<td background=""images/bk.gif"" width=300>比例(来访次数)</td>"
		response.write "<td background=""images/bk.gif"" width=100>历史</td>"
		response.write "<td background=""images/bk.gif"" width=200>展开</td>"
		response.write "</tr>"
		if not rs.eof then
			do while not rs.eof
				response.write "<tr>"
				response.write "<td bgcolor=#ffffff>"
				response.write "<a href=""searchname.asp?startDay="&startDay&"&endday="&endDay&"&searchname="&rs("searchName")&""">"&rs("searchName")&"</a>"
				response.write "</td>"

				response.write "<td bgcolor=#ffffff>"
				response.write "<img src=""images/bar2.gif"" width="&rs("id")/total*100&" height=""10""/> "
				response.write FormatNumber(rs("id")/total,4)*100&"%"
				response.write "</td>"

				
				response.write "<td bgcolor=#ffffff>"
				response.write "<a href=""searchname.asp?startDay="&startDay&"&endday="&endDay&"&searchname="&rs("searchName")&""">历史</a>"
				response.write "</td>"
				
				response.write "<td bgcolor=#ffffff>"
				response.write "<a href =""javascript:"" onclick=""SeeSearchKeyword(this);"">查看搜索引擎>></a>"
				response.write "</td>"
				
				response.write "</tr>"
				
				response.Write "<tr style=""display:none;""><td colspan=""4"" bgcolor=#ffffff align=center>"
				AloneKeyword rs("searchName"),rs("id")
				response.Write "</td></tr>"
				
				t=t&rs("searchName")&","
				v=v&rs("id")&","
				
				rs.movenext
			loop
			if t<>"" then
				t=left(t,(len(t)-1))
				v=left(v,(len(v)-1))
			end if
			rs.close
			response.write "</table>"
		else
			response.write "<tr><td  colspan=""4"">"
			response.write "暂无信息"
			response.write "</td></tr>"
			response.write "</table>"
		end if
		if total<>"0" then 
			imgUrl="img.asp?type=shanxing&t="&t&"&v="&v
			response.write "<script>ifrImg.location.href='"&imgUrl&"';</script>"
		else
			response.write "<script>ifrImg.document.write(""暂无图形"");</script>"
		end if
	end if
end sub
'搜索引擎历史访问
sub SearchName()

	if startDay<>"" and endDay<>"" and strSearchName<>"" then
	
		dim sql
		response.Write "<div style=""height:20px;padding-top:5px"">搜索引擎来访分布 - 从 <font color=red>"&request("startday")&"</font> 到 <font color=red>"&request("endday")&"</font> 的统计 </div>"
		response.Write "<iframe src="""" id=""ifrImg"" name=""ifrImg"" width=""800"" height=""400"" frameborder=""0"" scrolling=""no""></iframe>"
		GetSearchNameDayTotal(strSearchName)
		sql="select * from ("
		for i=cdate(startDay) to cdate(endDay)
			sql=sql+"SELECT count(*) as c,#"&i&"# as visitDate FROM visitor where visitdate=#"&i&"# and searchname='"&strSearchName&"' "
			t=t+cstr(day(i))
			if i<>cdate(endDay) then
				sql=sql+"union "
				t=t+","
			end if
		next
		sql=sql+") order by visitDate"

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open sql,conn,1,1

		response.write "<table width=800 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">日期</td>"
		response.write "<td background=""images/bk.gif"" width=300>比例(来访次数)</td>"
		response.write "<td background=""images/bk.gif"" width=100>比例图</td>"
		response.write "</tr>"
		if not rs.eof then
			do while not rs.eof 
				response.write "<tr>"
				response.write "<td bgcolor=#ffffff>"
				response.write rs("visitDate")
				response.write "</td>"

				response.write "<td bgcolor=#ffffff>"
				response.write rs("c")
				response.write "</td>"
				
				response.write "<td bgcolor=#ffffff>"
				if total=0 then
					response.Write "0%"
				else
					response.write "<img src=""images/bar2.gif"" width="&rs("c")/total*100&" height=""10""/> "
					response.write FormatNumber(rs("c")/total,4)*100&"%"
				end if
				response.write "</td>"
				
				
				response.write "</tr>"
				v=v&rs("c")&","
				
				rs.movenext
			loop
			if v<>"" then
				v=left(v,(len(v)-1))
			end if
			rs.close
			response.write "</table>"
		else
			response.write "<tr><td  colspan=""4"">"
			response.write "暂无信息"
			response.write "</td></tr>"
			response.write "</table>"
		end if
		if total<>"0" then 
			imgUrl="img.asp?type=zhuxing&d="&t&"&v="&v
			response.write "<script>ifrImg.location.href='"&imgUrl&"';</script>"
		else
			response.write "<script>ifrImg.document.write(""暂无图形"");</script>"
		end if
	end if

end sub
'单独搜索引擎关键字
sub AloneKeyword(sname,ccount)
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql="select count(1) as c,keyword from visitor where  visitDate>=#"&startday&"# and visitDate<=#"&endday&"# and issearch=0 and searchname='"&sname&"' group by keyword"
		rs.Open sql,conn,1,1
		response.write "<table width=600 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td  bgcolor=#ffffff width=200 height=""22"">关键字</td>"
		response.write "<td  bgcolor=#ffffff width=300>次数</td>"
		response.write "<td  bgcolor=#ffffff width=300>比例</td>"
		response.write "</tr>"
		if not rs.eof then
			do while not rs.eof 
				response.write "<tr>"
				response.write "<td bgcolor=#ffffff>"
				response.write rs("keyword")
				response.write "</td>"

				response.write "<td bgcolor=#ffffff>"
				response.write rs("c")
				response.write "</td>"
				
	
				response.write "<td bgcolor=#ffffff align=left>"
				if ccount=0 then
					response.Write "0%"
				else
					response.write "<img src=""images/bar2.gif"" width="&rs("c")/ccount*100&" height=""10""/> "
					response.write FormatNumber(rs("c")/ccount,4)*100&"%"
				end if
				response.write "</td>"
				
				response.write "</tr>"
				
				rs.movenext
			loop

			rs.close
			response.write "</table>"
		else
			response.write "<tr><td  colspan=""4"">"
			response.write "暂无信息"
			response.write "</td></tr>"
			response.write "</table>"
		end if
end sub
'关键字查询
sub SearchKeyword()
	if startDay<>"" and endDay<>"" then
		GetSearchTotal()
		response.Write "<div style=""height:20px;padding-top:5px"">搜索关键字 - 从 <font color=red>"&request("startday")&"</font> 到 <font color=red>"&request("endday")&"</font> 的统计 </div>"
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql="select count(1) as c,keyword from visitor where  visitDate>=#"&startday&"# and visitDate<=#"&endday&"# and issearch=0 group by keyword"
		rs.Open sql,conn,1,1
		
		response.write "<table width=800 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">搜索关键字</td>"
		response.write "<td background=""images/bk.gif"" width=300>次数</td>"
		response.write "<td background=""images/bk.gif"" width=300>比例</td>"
		response.write "<td background=""images/bk.gif"" width=100></td>"
		response.write "</tr>"
		if not rs.eof then
			do while not rs.eof 
				response.write "<tr>"
				response.write "<td bgcolor=#ffffff>"
				response.write rs("keyword")
				response.write "</td>"

				response.write "<td bgcolor=#ffffff>"
				response.write rs("c")
				response.write "</td>"
				
	
				response.write "<td bgcolor=#ffffff>"
				if total=0 then
					response.Write "0%"
				else
					response.write "<img src=""images/bar2.gif"" width="&rs("c")/total*100&" height=""10""/> "
					response.write FormatNumber(rs("c")/total,4)*100&"%"
				end if
				response.write "</td>"
				
				response.write "<td bgcolor=#ffffff style=""width:300px"">"
				response.write "<a href =""javascript:"" onclick=""SeeSearchKeyword(this);"">查看搜索引擎>></a>"
				response.write "</td>"	
				response.write "</tr>"
				
				response.Write "<tr style=""display:none;""><td colspan=""4"" bgcolor=#ffffff align=center>"
				SearchNameKeyword rs("keyword"),rs("c")
				response.Write "</td></tr>"
				
				rs.movenext
			loop

			rs.close
			response.write "</table>"
		else
			response.write "<tr><td  colspan=""4"">"
			response.write "暂无信息"
			response.write "</td></tr>"
			response.write "</table>"
		end if
	end if
end sub
'关键字搜索引擎
sub SearchNameKeyword(keyword,keycount)
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql="select count(1) as c,searchname from visitor where  visitDate>=#"&startday&"# and visitDate<=#"&endday&"# and issearch=0 and keyword='"&keyword&"' group by searchname"
		rs.Open sql,conn,1,1
		response.write "<table width=600 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td  bgcolor=#ffffff width=200 height=""22"">搜索引擎名称</td>"
		response.write "<td  bgcolor=#ffffff width=300>次数</td>"
		response.write "<td  bgcolor=#ffffff width=300>比例</td>"
		response.write "</tr>"
		if not rs.eof then
			do while not rs.eof 
				response.write "<tr>"
				response.write "<td bgcolor=#ffffff>"
				response.write rs("searchname")
				response.write "</td>"

				response.write "<td bgcolor=#ffffff>"
				response.write rs("c")
				response.write "</td>"
				
	
				response.write "<td bgcolor=#ffffff align=left>"
				if keycount=0 then
					response.Write "0%"
				else
					response.write "<img src=""images/bar2.gif"" width="&rs("c")/keycount*100&" height=""10""/> "
					response.write FormatNumber(rs("c")/keycount,4)*100&"%"
				end if
				response.write "</td>"
				
				response.write "</tr>"
				
				rs.movenext
			loop

			rs.close
			response.write "</table>"
		else
			response.write "<tr><td  colspan=""4"">"
			response.write "暂无信息"
			response.write "</td></tr>"
			response.write "</table>"
		end if
end sub
'搜索引擎最近访问
sub SearchNearVisitor()
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT top 15 *FROM visitor where issearch=0 order by visitdate desc,visittime desc"
	rs.Open sql,conn,1,1
			response.write "<table width=600 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td  background=""images/bk.gif"" width=200 height=""22"">搜索引擎</td>"
		response.write "<td  background=""images/bk.gif"" width=300>关键字</td>"
		response.write "<td  background=""images/bk.gif"" width=300>时间</td>"
		response.write "<td  background=""images/bk.gif"" width=300>来源</td>"
		response.write "</tr>"
		if not rs.eof then
			do while not rs.eof 
				response.write "<tr>"
				response.write "<td bgcolor=#ffffff>"
				response.write rs("searchname")
				response.write "</td>"

				response.write "<td bgcolor=#ffffff>"
				response.write rs("keyword")
				response.write "</td>"
				
	
				response.write "<td bgcolor=#ffffff align=left>"
				response.write rs("visitdate")&" "&rs("visittime")
				response.write "</td>"
				
				response.write "<td bgcolor=#ffffff align=left>"
				response.write "<a href ="""& rs("lastpage")&""" target=""_blank"">点击这里</a>"
				response.write "</td>"
				
				response.write "</tr>"
				
				rs.movenext
			loop

			rs.close
			response.write "</table>"
		else
			response.write "<tr><td  colspan=""4"">"
			response.write "暂无信息"
			response.write "</td></tr>"
			response.write "</table>"
		end if
end sub
'某个受访页总数
sub GetPageUrlTotal(pageUrl)
	Set rs = Server.CreateObject("ADODB.Recordset")
	Sql = "select count(1) as c from visitor where visitDate>=#"&startday&"# and visitDate<=#"&endday&"# and pageUrl='"&pageUrl&"'"
	rs.Open Sql,conn,1,1
	if not rs.eof then
		total= rs("c")
	else
		response.write "0"
	end if
end sub

'某个系统访问总数
sub GetSystemTotal(systemName)
	Set rs = Server.CreateObject("ADODB.Recordset")
	Sql = "select count(1) as c from visitor where visitDate>=#"&startday&"# and visitDate<=#"&endday&"# and OpSystem='"&systemName&"'"
	rs.Open Sql,conn,1,1
	if not rs.eof then
		total= rs("c")
	else
		response.write "0"
	end if
end sub

'某个分辨率总数
sub GetScreenTotal(screenName)
	Set rs = Server.CreateObject("ADODB.Recordset")
	Sql = "select count(1) as c from visitor where visitDate>=#"&startday&"# and visitDate<=#"&endday&"# and screen='"&screenName&"'"
	rs.Open Sql,conn,1,1
	if not rs.eof then
		total= rs("c")
	else
		response.write "0"
	end if
end sub

'浏览器访问总数
sub GetBrowserTotal(browserName)
	Set rs = Server.CreateObject("ADODB.Recordset")
	Sql = "select count(1) as c from visitor where visitDate>=#"&startday&"# and visitDate<=#"&endday&"# and browser='"&browserName&"'"
	rs.Open Sql,conn,1,1
	if not rs.eof then
		total= rs("c")
	else
		response.write "0"
	end if
end sub
'搜索引擎历史访问总数
sub GetSearchNameDayTotal(strSearchName)
	Set rs = Server.CreateObject("ADODB.Recordset")
	Sql = "select count(1) as c from visitor where visitDate>=#"&startday&"# and visitDate<=#"&endday&"# and searchname='"&strSearchName&"'"
	rs.Open Sql,conn,1,1
	if not rs.eof then
		total= rs("c")
	else
		response.write "0"
	end if
end sub

'日访问量表格
sub DayVisit()

	if startDay<>"" and endDay<>"" then
		GetTotal()
		set rs=conn.execute("select count(id) as id,visitDate from visitor where visitDate>=#"&startday&"# and visitDate<=#"&endday&"# group by visitDate order by visitDate desc")
		
		response.Write "<div style=""height:20px;padding-top:5px"">时段分析 - 从 <font color=red>"&request("startday")&"</font> 到 <font color=red>"&request("endday")&"</font> 的统计 </div>"
		
		response.Write "<div class=""Title"">日访问量分布</div>"
		
		response.write "<table width=800 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">日期</td>"
		response.write "<td background=""images/bk.gif"" width=200>PV访问量</td>"
		response.write "<td background=""images/bk.gif"" width=200>唯一IP</td>"
		response.write "<td background=""images/bk.gif"" width=200>比例</td>"
		response.write "</tr>"
		
		if not rs.eof then
			do while not rs.eof
				response.write "<tr>"
				response.write "<td bgcolor=#ffffff>"
				response.write rs("visitDate")
				response.write "</td>"

				response.write "<td bgcolor=#ffffff>"
				response.write rs("id")
				response.write "</td>"

				
				response.write "<td bgcolor=#ffffff>"
				AloneIp cstr(rs("visitDate"))
				'response.write rs("visitDate")
				response.write "</td>"
				
				response.write "<td bgcolor=#ffffff>"
				response.write "<img src=""images/bar2.gif"" width="&rs("id")/total*100&" height=""10""/> "
				response.write FormatNumber(rs("id")/total,4)*100&"%"
				response.write "</td>"
				
				response.write "</tr>"
				rs.movenext
			loop
			rs.close
			response.write "</table>"
		else
			response.write "<tr><td  colspan=""4"">"
			response.write "暂无信息"
			response.write "</td></tr>"
			response.write "</table>"
		end if
	end if
end sub
'日总数
sub GetTotal()
	Set rs = Server.CreateObject("ADODB.Recordset")
	Sql = "select count(id) as d from visitor where visitDate>=#"&startday&"# and visitDate<=#"&endday&"#"
	rs.Open Sql,conn,1,1
	if not rs.eof then
		total= rs("d")
	else
		total= "0"
	end if
end sub
'搜速引擎访问总数
sub GetSearchTotal()
	Set rs = Server.CreateObject("ADODB.Recordset")
	Sql = "select count(id) as d from visitor where isSearch=0 and visitDate>=#"&startday&"# and visitDate<=#"&endday&"#"
	rs.Open Sql,conn,1,1
	if not rs.eof then
		total= rs("d")
	else
		total= "0"
	end if
end sub
'日唯一ip
sub AloneIp(today)
	Set rs = Server.CreateObject("ADODB.Recordset")
	Sql = "select count(id) as d,userIp from visitor where visitDate=#"&today&"# group by userIp"
	rs.Open Sql,conn,1,1
	if not rs.eof then
		response.write rs.recordcount
	else
		response.write "0"
	end if

end sub

'时段统计
sub HoursVisit()
	if startDay<>"" and startDay<>"" then
		
		response.Write "<div class=""Title"">24小时访问量分布</div>"
		response.Write "<iframe src="""" id=""ifrImg"" name=""ifrImg"" width=""800"" height=""400"" frameborder=""0"" scrolling=""no""></iframe>"
		response.write "<table width=800 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">时间段</td>"
		response.write "<td background=""images/bk.gif"" width=200>PV访问量</td>"
		response.write "<td background=""images/bk.gif"" width=200>唯一IP</td>"
		response.write "<td background=""images/bk.gif"" width=200>比例</td>"
		response.write "</tr>"
		for i=0 to 23
			if i=23 then  
				HourVisit "23:00:00", "00:00:00" 
			else
				HourVisit Cstr(i)+":00:00",Cstr(i+1)&":00:00" 
		    end if
		next
		response.write "</table>"
	end if
	if total<>"0" then 
		imgUrl="img.asp?type=zhuxing&dataType=hours&v="&v
		response.write "<script>ifrImg.location.href='"&imgUrl&"';</script>"
	else
		response.write "<script>ifrImg.document.write(""暂无图形"");</script>"
	end if
end sub

'小时计算
sub HourVisit(time1,time2)
	dim sql
	sql="select count(id) as d from visitor where visitDate>=#"&startDay&"# and visitDate<=#"&endDay&"# and visitTime>=#"&time1&"#"
	if time2<>"00:00:00" then
		sql=sql&" and visitTime<=#"&time2&"#  "
	end if
	set rs=conn.execute(sql)
	
	if not rs.eof and total<>"0" then
		do while not rs.eof 
			response.write "<tr>"
			response.write "<td bgcolor=#ffffff>"
			response.write time1&" 至 "&time2
			response.write "</td>"

			response.write "<td bgcolor=#ffffff>"
			response.write rs("d")
			response.write "</td>"

			
			response.write "<td bgcolor=#ffffff>"
			AloneHoursIp time1,time2
			response.write "</td>"
			
			response.write "<td bgcolor=#ffffff>"
			response.write "<img src=""images/bar2.gif"" width="&rs("d")/total*100&" height=""10""/> "
			response.write FormatNumber(rs("d")/total,4)*100&"%"
			response.write "</td>"
			
			response.write "</tr>"
			if v="" then
				v=rs("d") 
			else
				v=v&","&rs("d")
			end if
			rs.movenext
		loop
		rs.close

	else
		response.write "<tr>"
		response.write "<td bgcolor=#ffffff>"
		response.write time1&" 至 "&time2
		response.write "</td>"

		response.write "<td bgcolor=#ffffff>"
		response.write "0"
		response.write "</td>"

		
		response.write "<td bgcolor=#ffffff>"
		response.write "0"
		response.write "</td>"
		
		response.write "<td bgcolor=#ffffff>"
		response.write "0"
		response.write "</td>"
		
		response.write "</tr>"
		if v="" then
			v=rs("d") 
		else
			v=v&","&rs("d")
		end if
	end if
		
end sub

'时间段唯一ip
sub AloneHoursIp(time1,time2)
	set rs = Server.CreateObject("ADODB.Recordset")
	sql = "select count(id) as id,userIp from visitor where visitDate>=#"&startDay&"# and visitDate<=#"&endDay&"# and visitTime>=#"&time1&"# and visitTime<=#"&time2&"#  group by userIp"
	rs.Open sql,conn,1,1
	if not rs.eof then
		response.write rs.recordcount
	else
		response.write "0"
	end if
	rs.close
end sub
%>