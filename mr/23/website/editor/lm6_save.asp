<%
	Dim Recordset__MRColParam
	Recordset__MRColParam = "1"
	if (Session("MR_username") <> "") then
		Recordset__MRColParam = Session("MR_username")
	end if

	dim rContent
	rContent = request("content")
	rContent = replace(rContent,"'","''")

	dim conn
	dim connstr
	Set conn = Server.CreateObject("ADODB.Connection")
	connstr="DBQ="+server.mappath("/database/db_website.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
	conn.Open connstr
	conn.execute "update [website] set main6='" & rContent & "' where id='" & Recordset__MRColParam & "'"
	conn.close
	set conn=nothing
%>
<body bgcolor="#CCCCCC">
<script>alert('���ϸ��³ɹ�!�������"Ԥ����վ"��ť�ۿ�Ч��');window.close();window.opener.location="../lm_manage.asp";</script>
</body>