<!--#include file=drawimg.asp-->
<%
dim etdbvote'�����ļ����ݿ����ӵ�ַ
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
	response.write "<td background=""images/bk.gif"" width=200 height=""22"">����</td>"
	response.write "<td background=""images/bk.gif"" width=200>��Դҳ��</td>"
	response.write "<td background=""images/bk.gif"" width=300>����ҳ��</td>"
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
		response.write "������Ϣ"
		response.write "</td></tr>"
		response.write "</table>"
	end if
end sub


'�ܷ�ҳ��ʷ
sub PageComeHistory()

	if startDay<>"" and endDay<>"" and pageUrl<>"" then

		dim sql
		response.Write "<div class=""Title"">��Դҳ��:<a href="""&pageUrl&""">"&pageUrl&"</a> </div>"
		response.Write "<div style=""height:20px;padding-top:5px"">ʱ���: <font color=red>"&request("startday")&"</font> �� <font color=red>"&request("endday")&"</font> ��ͳ�� </div>"
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
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">����</td>"
		response.write "<td background=""images/bk.gif"" width=200>�������(PV)</td>"
		response.write "<td background=""images/bk.gif"" width=300>����ͼ</td>"
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
			response.write "������Ϣ"
			response.write "</td></tr>"
			response.write "</table>"
		end if
		if total<>"0" then 
			imgUrl="img.asp?type=zhuxing&d="&t&"&v="&v
			response.write "<script>ifrImg.location.href='"&imgUrl&"';</script>"
		else
			response.write "<script>ifrImg.document.write(""����ͼ��"");</script>"
		end if
	end if

end sub

'�ܷ�ҳ�����
sub PageCome()
'��ҳ�����һ�Σ�����������������������������
	if startday<>"" and endday<>"" then
		response.Write "<div class=""Title"">�ܷ�ҳ�����: </div>"
		response.Write "<div style=""height:20px;padding-top:5px"">ʱ���: <font color=red>"&request("startday")&"</font> �� <font color=red>"&request("endday")&"</font> ��ͳ�� </div>"
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
			response.write "<td background=""images/bk.gif"" width=200 height=""22"">����ҳ</td>"
			response.write "<td background=""images/bk.gif"" width=200>�������(PV)</td>"
			response.write "<td background=""images/bk.gif"" width=300>����ͼ</td>"
			response.write "<td background=""images/bk.gif"" width=300>��ʷ</td>"
			response.write "</tr>"
		'��������������������������������ҳ�����һ��
		'ѭ�����ݶΣ�����������������
				
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
					response.Write """>��ʷ</a>"
					response.write "</td>"
					
					response.write "</tr>"
	
					rs.movenext
					end if
				next
				response.write "</table>"
			rs.close
		'����������������ѭ�����ݶ�
		
		'��ҳ����ڶ��Σ�����������������
				response.write "<table width='95%' >"
				response.write "<tr><td colspan=3 align=right>"
				k=currentpage
				response.write "��"&totalput&"��&nbsp;&nbsp;"
				response.write "��"&n&"ҳ&nbsp;&nbsp;"
				response.write "��ǰҳ:"&currentpage&"&nbsp;&nbsp;"
				if k<>1 then
				response.write "��"&maxperpage&"ҳ&nbsp;&nbsp;"
				response.write "��ǰҳ:"&currentpage&"&nbsp;&nbsp;"
				response.write "<a href='?startday="&startday&"&endday="&endday&"&page=1'>��ҳ</a>&nbsp;&nbsp;"
				response.write "<a href='?startday="&startday&"&endday="&endday&"&page="&k-1&"'>��һҳ</a>&nbsp;&nbsp;"
				else
				response.write "��ҳ&nbsp;&nbsp;��һҳ&nbsp;&nbsp;"
				end if
				if k<>n then
				response.write "<a href='?startday="&startday&"&endday="&endday&"&page="&k+1&"'>��һҳ</a>&nbsp;&nbsp;"
				response.write "<a href='?startday="&startday&"&endday="&endday&"&page="&n&"'>βҳ</a>"
				else
				response.write "��һҳ&nbsp;&nbsp;βҳ"
				end if
				response.write "</td></tr></table>"
		else
			response.write ""
		end if
	end if
'������������������ҳ����ڶ���
end sub


'�ֱ�����ʷ
sub ScreenrHistory()
	if startDay<>"" and endDay<>"" and screenName<>"" then
	
		dim sql
		response.Write "<div class=""Title"">�ֱ���:"&screenName&" </div>"
		response.Write "<div style=""height:20px;padding-top:5px"">ʱ���: <font color=red>"&request("startday")&"</font> �� <font color=red>"&request("endday")&"</font> ��ͳ�� </div>"
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
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">����</td>"
		response.write "<td background=""images/bk.gif"" width=200>�������(PV)</td>"
		response.write "<td background=""images/bk.gif"" width=300>����ͼ</td>"
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
			response.write "������Ϣ"
			response.write "</td></tr>"
			response.write "</table>"
		end if
		if total<>"0" then 
			imgUrl="img.asp?type=zhuxing&d="&t&"&v="&v
			response.write "<script>ifrImg.location.href='"&imgUrl&"';</script>"
		else
			response.write "<script>ifrImg.document.write(""����ͼ��"");</script>"
		end if
	end if

end sub

'�������ʷ
sub BrowserHistory()
	if startDay<>"" and endDay<>"" and browserName<>"" then
	
		dim sql
		response.Write "<div class=""Title"">���������:"&browserName&" </div>"
		response.Write "<div style=""height:20px;padding-top:5px"">ʱ���: <font color=red>"&request("startday")&"</font> �� <font color=red>"&request("endday")&"</font> ��ͳ�� </div>"
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
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">����</td>"
		response.write "<td background=""images/bk.gif"" width=200>�������(PV)</td>"
		response.write "<td background=""images/bk.gif"" width=300>����ͼ</td>"
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
			response.write "������Ϣ"
			response.write "</td></tr>"
			response.write "</table>"
		end if
		if total<>"0" then 
			imgUrl="img.asp?type=zhuxing&d="&t&"&v="&v
			response.write "<script>ifrImg.location.href='"&imgUrl&"';</script>"
		else
			response.write "<script>ifrImg.document.write(""����ͼ��"");</script>"
		end if
	end if

end sub

'����������ʷ����
sub OpSystemHistory()
	if startDay<>"" and endDay<>"" and systemName<>"" then
	
		dim sql
		response.Write "<div class=""Title"">����ϵͳ:"&systemName&" </div>"
		response.Write "<div style=""height:20px;padding-top:5px"">ʱ���: <font color=red>"&request("startday")&"</font> �� <font color=red>"&request("endday")&"</font> ��ͳ�� </div>"
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
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">����</td>"
		response.write "<td background=""images/bk.gif"" width=200>�������(PV)</td>"
		response.write "<td background=""images/bk.gif"" width=300>����ͼ</td>"
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
			response.write "������Ϣ"
			response.write "</td></tr>"
			response.write "</table>"
		end if
		if total<>"0" then 
			imgUrl="img.asp?type=zhuxing&d="&t&"&v="&v
			response.write "<script>ifrImg.location.href='"&imgUrl&"';</script>"
		else
			response.write "<script>ifrImg.document.write(""����ͼ��"");</script>"
		end if
	end if

end sub

'�ÿͷֱ���
sub ScreenApart()
	if startDay<>"" and endDay<>"" then

		response.Write "<div class=""Title"">�ֱ��ʷֲ� </div>"
		GetTotal()
		set rs=conn.execute("select count(id) as c,screen from visitor where  visitDate>=#"&startday&"# and visitDate<=#"&endday&"#  group by screen order by screen")
		response.write "<table width=800 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">�ֱ���</td>"
		response.write "<td background=""images/bk.gif"" width=300>��ռ����(���ô���)</td>"
		response.write "<td background=""images/bk.gif"" width=100>��ʷ</td>"

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
				response.write "<a href=""screenHistory.asp?startDay="&startDay&"&endday="&endDay&"&screenName="&rs("screen")&""">��ʷ</a>"
				response.write "</td>"
				
				
				response.write "</tr>"
				
				
				rs.movenext
			loop
			rs.close
			response.write "</table>"
		else
			response.write "<tr><td  colspan=""4"">"
			response.write "������Ϣ"
			response.write "</td></tr>"
			response.write "</table>"
		end if
	end if
end sub
'�ÿ������
sub OpBrowser()
	if startDay<>"" and endDay<>"" then

		response.Write "<div class=""Title"">������ֲ� </div>"
		GetTotal()
		set rs=conn.execute("select count(id) as c,Browser from visitor where visitDate>=#"&startday&"# and visitDate<=#"&endday&"#  group by Browser order by Browser")
		response.write "<table width=800 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">���������</td>"
		response.write "<td background=""images/bk.gif"" width=300>��ռ����(���ô���)</td>"
		response.write "<td background=""images/bk.gif"" width=100>��ʷ</td>"

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
				response.write """>��ʷ</a>"
				response.write "</td>"
				
				
				response.write "</tr>"
				
				
				rs.movenext
			loop
			rs.close
			response.write "</table>"
		else
			response.write "<tr><td  colspan=""4"">"
			response.write "������Ϣ"
			response.write "</td></tr>"
			response.write "</table>"
		end if
	end if
end sub


'�ÿͲ���ϵͳ
sub OpSystem()
	if startDay<>"" and endDay<>"" then
	response.Write "<div style=""height:20px;padding-top:5px"">�ͻ�����ϸ - �� <font color=red>"&request("startday")&"</font> �� <font color=red>"&request("endday")&"</font> ��ͳ�� </div>"
		response.Write "<div class=""Title"">�ÿͲ���ϵͳ </div>"
		GetTotal()
		set rs=conn.execute("select count(id) as c,opsystem from visitor where  visitDate>=#"&startday&"# and visitDate<=#"&endday&"#  group by opsystem order by opsystem")
		response.write "<table width=800 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">����ϵͳ����</td>"
		response.write "<td background=""images/bk.gif"" width=300>��ռ����(���ô���)</td>"
		response.write "<td background=""images/bk.gif"" width=100>��ʷ</td>"

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
				response.write """>��ʷ</a>"
				response.write "</td>"
				
				
				response.write "</tr>"
				
				
				rs.movenext
			loop
			rs.close
			response.write "</table>"
		else
			response.write "<tr><td  colspan=""4"">"
			response.write "������Ϣ"
			response.write "</td></tr>"
			response.write "</table>"
		end if
	end if
end sub

'��������
sub SearchCome()
	if startDay<>"" and endDay<>"" then
		GetSearchTotal()
		response.Write "<div style=""height:20px;padding-top:5px"">�����������÷ֲ� - �� <font color=red>"&request("startday")&"</font> �� <font color=red>"&request("endday")&"</font> ��ͳ�� </div>"
		response.Write "<div class=""Title"">�ܼƴ������������ã�"&total&" ��(PV) </div>"
		response.Write "<iframe src="""" id=""ifrImg"" name=""ifrImg"" width=""800"" height=""400"" frameborder=""0"" scrolling=""no""></iframe>"
		set rs=conn.execute("select count(id) as id,searchName from visitor where isSearch=0 and  visitDate>=#"&startday&"# and visitDate<=#"&endday&"#  group by searchName order by searchName")
		response.write "<table width=800 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">������������</td>"
		response.write "<td background=""images/bk.gif"" width=300>����(���ô���)</td>"
		response.write "<td background=""images/bk.gif"" width=100>��ʷ</td>"
		response.write "<td background=""images/bk.gif"" width=200>չ��</td>"
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
				response.write "<a href=""searchname.asp?startDay="&startDay&"&endday="&endDay&"&searchname="&rs("searchName")&""">��ʷ</a>"
				response.write "</td>"
				
				response.write "<td bgcolor=#ffffff>"
				response.write "<a href =""javascript:"" onclick=""SeeSearchKeyword(this);"">�鿴��������>></a>"
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
			response.write "������Ϣ"
			response.write "</td></tr>"
			response.write "</table>"
		end if
		if total<>"0" then 
			imgUrl="img.asp?type=shanxing&t="&t&"&v="&v
			response.write "<script>ifrImg.location.href='"&imgUrl&"';</script>"
		else
			response.write "<script>ifrImg.document.write(""����ͼ��"");</script>"
		end if
	end if
end sub
'����������ʷ����
sub SearchName()

	if startDay<>"" and endDay<>"" and strSearchName<>"" then
	
		dim sql
		response.Write "<div style=""height:20px;padding-top:5px"">�����������÷ֲ� - �� <font color=red>"&request("startday")&"</font> �� <font color=red>"&request("endday")&"</font> ��ͳ�� </div>"
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
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">����</td>"
		response.write "<td background=""images/bk.gif"" width=300>����(���ô���)</td>"
		response.write "<td background=""images/bk.gif"" width=100>����ͼ</td>"
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
			response.write "������Ϣ"
			response.write "</td></tr>"
			response.write "</table>"
		end if
		if total<>"0" then 
			imgUrl="img.asp?type=zhuxing&d="&t&"&v="&v
			response.write "<script>ifrImg.location.href='"&imgUrl&"';</script>"
		else
			response.write "<script>ifrImg.document.write(""����ͼ��"");</script>"
		end if
	end if

end sub
'������������ؼ���
sub AloneKeyword(sname,ccount)
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql="select count(1) as c,keyword from visitor where  visitDate>=#"&startday&"# and visitDate<=#"&endday&"# and issearch=0 and searchname='"&sname&"' group by keyword"
		rs.Open sql,conn,1,1
		response.write "<table width=600 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td  bgcolor=#ffffff width=200 height=""22"">�ؼ���</td>"
		response.write "<td  bgcolor=#ffffff width=300>����</td>"
		response.write "<td  bgcolor=#ffffff width=300>����</td>"
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
			response.write "������Ϣ"
			response.write "</td></tr>"
			response.write "</table>"
		end if
end sub
'�ؼ��ֲ�ѯ
sub SearchKeyword()
	if startDay<>"" and endDay<>"" then
		GetSearchTotal()
		response.Write "<div style=""height:20px;padding-top:5px"">�����ؼ��� - �� <font color=red>"&request("startday")&"</font> �� <font color=red>"&request("endday")&"</font> ��ͳ�� </div>"
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql="select count(1) as c,keyword from visitor where  visitDate>=#"&startday&"# and visitDate<=#"&endday&"# and issearch=0 group by keyword"
		rs.Open sql,conn,1,1
		
		response.write "<table width=800 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">�����ؼ���</td>"
		response.write "<td background=""images/bk.gif"" width=300>����</td>"
		response.write "<td background=""images/bk.gif"" width=300>����</td>"
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
				response.write "<a href =""javascript:"" onclick=""SeeSearchKeyword(this);"">�鿴��������>></a>"
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
			response.write "������Ϣ"
			response.write "</td></tr>"
			response.write "</table>"
		end if
	end if
end sub
'�ؼ�����������
sub SearchNameKeyword(keyword,keycount)
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql="select count(1) as c,searchname from visitor where  visitDate>=#"&startday&"# and visitDate<=#"&endday&"# and issearch=0 and keyword='"&keyword&"' group by searchname"
		rs.Open sql,conn,1,1
		response.write "<table width=600 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td  bgcolor=#ffffff width=200 height=""22"">������������</td>"
		response.write "<td  bgcolor=#ffffff width=300>����</td>"
		response.write "<td  bgcolor=#ffffff width=300>����</td>"
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
			response.write "������Ϣ"
			response.write "</td></tr>"
			response.write "</table>"
		end if
end sub
'���������������
sub SearchNearVisitor()
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT top 15 *FROM visitor where issearch=0 order by visitdate desc,visittime desc"
	rs.Open sql,conn,1,1
			response.write "<table width=600 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td  background=""images/bk.gif"" width=200 height=""22"">��������</td>"
		response.write "<td  background=""images/bk.gif"" width=300>�ؼ���</td>"
		response.write "<td  background=""images/bk.gif"" width=300>ʱ��</td>"
		response.write "<td  background=""images/bk.gif"" width=300>��Դ</td>"
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
				response.write "<a href ="""& rs("lastpage")&""" target=""_blank"">�������</a>"
				response.write "</td>"
				
				response.write "</tr>"
				
				rs.movenext
			loop

			rs.close
			response.write "</table>"
		else
			response.write "<tr><td  colspan=""4"">"
			response.write "������Ϣ"
			response.write "</td></tr>"
			response.write "</table>"
		end if
end sub
'ĳ���ܷ�ҳ����
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

'ĳ��ϵͳ��������
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

'ĳ���ֱ�������
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

'�������������
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
'����������ʷ��������
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

'�շ��������
sub DayVisit()

	if startDay<>"" and endDay<>"" then
		GetTotal()
		set rs=conn.execute("select count(id) as id,visitDate from visitor where visitDate>=#"&startday&"# and visitDate<=#"&endday&"# group by visitDate order by visitDate desc")
		
		response.Write "<div style=""height:20px;padding-top:5px"">ʱ�η��� - �� <font color=red>"&request("startday")&"</font> �� <font color=red>"&request("endday")&"</font> ��ͳ�� </div>"
		
		response.Write "<div class=""Title"">�շ������ֲ�</div>"
		
		response.write "<table width=800 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">����</td>"
		response.write "<td background=""images/bk.gif"" width=200>PV������</td>"
		response.write "<td background=""images/bk.gif"" width=200>ΨһIP</td>"
		response.write "<td background=""images/bk.gif"" width=200>����</td>"
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
			response.write "������Ϣ"
			response.write "</td></tr>"
			response.write "</table>"
		end if
	end if
end sub
'������
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
'���������������
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
'��Ψһip
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

'ʱ��ͳ��
sub HoursVisit()
	if startDay<>"" and startDay<>"" then
		
		response.Write "<div class=""Title"">24Сʱ�������ֲ�</div>"
		response.Write "<iframe src="""" id=""ifrImg"" name=""ifrImg"" width=""800"" height=""400"" frameborder=""0"" scrolling=""no""></iframe>"
		response.write "<table width=800 border=0 cellpadding=1 cellspacing=1 bgcolor=#cccccc>"
		response.write "  <tr>"
   		response.write "<td background=""images/bk.gif"" width=200 height=""22"">ʱ���</td>"
		response.write "<td background=""images/bk.gif"" width=200>PV������</td>"
		response.write "<td background=""images/bk.gif"" width=200>ΨһIP</td>"
		response.write "<td background=""images/bk.gif"" width=200>����</td>"
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
		response.write "<script>ifrImg.document.write(""����ͼ��"");</script>"
	end if
end sub

'Сʱ����
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
			response.write time1&" �� "&time2
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
		response.write time1&" �� "&time2
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

'ʱ���Ψһip
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