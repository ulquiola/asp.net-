<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../../Conn/conn.asp"--><!--�������ݿ������ļ�-->
<!--#include file="../include/md5.asp" -->
<%

			sub click()
			if session("session('jiaoyanma')") <> ""Then
				Sql_2 = "delete * from tb_jiaoyanma Where jiaoyanma = '"&session("session('jiaoyanma')")&"'"
				Set rs_2 = conn.Execute(Sql_2)
			end if
			end sub
%>
<%
	id=request("id")
	dim rndnum,verifycode
	Randomize
	Do While Len(rndnum)<8
	num1=CStr(Chr((57-48)*rnd+48))
	rndnum=rndnum&num1
	loop
	set rs=server.CreateObject("adodb.recordset")
	Sql = "select * from tb_StuResult where id="&request("id")
	rs.Open Sql,conn,1,3

	'������¼��������ȡ����������ӵ����ݱ�tb_Student��
	set rs1=server.CreateObject("adodb.recordset")
	Sql1 = "select * from tb_Student "
	rs1.Open Sql1,conn,1,3
	rs1.AddNew
	rs1("UserID")=rs("����")
	rs1("PWD") = rndnum
	rs1("PWD2") = md5(trim(rndnum))
	rs1("Jointime")=now()
	rs1.Update
	call click()
	response.Write("<script language=javascript>alert('������ȡ������Ϊ��"&rndnum&"����������ã�');window.close();</script>")
	rs1.Close
	Set rs1 = Nothing
%>
