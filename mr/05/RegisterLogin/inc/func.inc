<%
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
dim reback		'��������ֵ
'�������ã���֤��ѯ��¼�Ƿ����
'������
'sql����ѯ���
sub chkform(sql)
	set rs = conn.execute(sql)
	if not rs.eof then
		reback = "1"			'��Ϊ1ʱ���м�¼
	else
		reback = "2"			'����2ʱ���޼�¼
	end if
end sub
'�������ã����Ӽ�¼
'����
'sql�����Ӽ�¼�����ݿ�
'fields1�����ӵ��ֶ�������
'value1�����ӵ��ֶ�ֵ����
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
'�������ã��������ݱ���¼
'������
'sql����ѯ���
'fields1���޸��ֶ�
'value1���ֶ�ֵ
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
'�������� ��ѯ���ݣ����ؽ��������һ��¼��
'����
'sql����ѯ����
'arr����ѯ�ֶ�����
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
'��ʾ��ʾ��Ϣ����
sub msg(str,locat)
	if isEmpty(locat) then
		response.Write("<script>alert('"&str&"');history.go(-1);</script>")
	else 
		response.Write("<script>alert('"&str&"');location='"&locat&"';</script>")
	end if
end sub
%>