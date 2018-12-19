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
<script>alert('资料更新成功!请点击左侧"预览网站"按钮观看效果');window.close();window.opener.location="../lm_manage.asp";</script>
</body>