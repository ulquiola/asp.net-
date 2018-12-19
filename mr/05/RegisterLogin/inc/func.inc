<%
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
dim reback		'函数返回值
'函数作用：验证查询记录是否存在
'参数：
'sql：查询语句
sub chkform(sql)
	set rs = conn.execute(sql)
	if not rs.eof then
		reback = "1"			'当为1时，有记录
	else
		reback = "2"			'等于2时，无记录
	end if
end sub
'函数作用：添加记录
'参数
'sql：添加记录的数据库
'fields1：添加的字段名数组
'value1：添加的字段值数组
sub addrst(sql,fields1,value1)
	set rs = server.CreateObject("adodb.recordset")
	rs.open sql,conn,3,3
	rs.addnew
	if isArray(fields1) = false then
  		rs(""&fields1&"") = value1
	else
		num = UBound(fields1)-1
		for i=0 to num
			rs(""&fields1(i)&"") = value1(i)
		Next
	end if
	rs.update()
	reback = "1"
end sub
'函数作用：更新数据表记录
'参数：
'sql：查询语句
'fields1：修改字段
'value1：字段值
sub modirst(sql,fields1,value1)
	set rs = server.CreateObject("adodb.recordset")
	rs.open sql,conn,3,3
	if isArray(fields1) = false then
  		rs(""&fields1&"") = value1
	else
		num = UBound(fields1) - 1
		for i=0 to num
			rs(""&fields1(i)&"") = value1(i)
		Next
	end if	
	rs.update()
	reback = "1"
end sub
'函数作用 查询数据，返回结果集（单一记录）
'参数
'sql：查询数组
'arr：查询字段数组
sub rebackrst(sql,arr)
	set rs = server.CreateObject("adodb.recordset")
	rs.open sql,conn,3,3
	if not rs.eof then
		if isArray(arr) = false then
			reback = rs(""&arr&"")
		else
			num = UBound(arr) -1
			reback = rs(""&arr(0)&"")
			for i = 1 to num
				reback = reback & ","&rs(""&arr(i)&"")
			Next
		end if
	else
		reback = "2"
	end if
end sub
'显示提示消息窗口
sub msg(str,locat)
	if isEmpty(locat) then
		response.Write("<script>alert('"&str&"');history.go(-1);</script>")
	else 
		response.Write("<script>alert('"&str&"');location='"&locat&"';</script>")
	end if
end sub
%>