<!--#include file="conn/conn.asp"-->
<!--#include file="inc/func.inc"-->
<%
	name = request.QueryString("name")
	pwd = request.QueryString("pwd")
	if name = "" or pwd = "" then
		call msg("�û������벻����Ϊ��",null)
		response.End()
	end if
	sql = "select count from tb_member where name = '"&name&"'"
	call chkform(sql)
	'���ȼ����û���
	if reback = 1 then		'����û������ڣ������������
		sql = sql&" and password = '"&pwd&"'"
		call chkform(sql)
		flds = "count"
		if reback = 1 then			'�û�������������ȷ
			session("name") = name
			response.Cookies("name") = name
			response.Cookies("name").Expires = now() + 600
			response.Cookies("count").Expires = now() - 1
			vle = 0
		else
			resql = "select count from tb_member where name = '"&name&"'"
			flds = "count"
			call rebackrst(resql,flds)
			if reback < 3 then
				vle = reback + 1
				call modirst(resql,flds,vle)
				reback = 2
			else
				reback = 4
			end if
		end if
	else					'�û��������ڣ���ʹ��cookie��¼��¼����
		if request.Cookies("count")="" then
			response.Cookies("count") = 1
			response.Cookies("count").Expires = now() + 600
			reback = 2
		elseif request.Cookies("count") < 3 then
			response.Cookies("count") = request.Cookies("count") + 1
			response.Cookies("count").Expires = now() + 600
			reback = 2
		else
			reback = 3
		end if
	end if
response.Write(reback)
%>
