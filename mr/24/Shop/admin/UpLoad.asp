<html>
<head>
<title>������ϴ���</title>
</head>
<body>
<!--#include FILE="upload.inc"-->
<%
formName=request("formName")
EditName=request("EditName")

set upload=new upload_5xsoft
set file=upload.file("file1")
if file.fileSize>0 then
	filenameend=file.filename
	filenameend=split(filenameend,".")
	n=UBound(filenameend)
	filename=filename&filenameend(n)
	
	const filepath="/upfile/"		'�ϴ�·��
	const filepathname = "/upfile/"
	dtNow=Now()
	randomize
	ranNum=int(90000*rnd)+10000
	filename1=year(dtNow) & right("0" & month(dtNow),2) & right("0" & day(dtNow),2) & right("0" & hour(dtNow),2) & right("0" & minute(dtNow),2) & right("0" & second(dtNow),2) & ranNum &"."&fileExt
	filename=filepath&filename1&filenameend(n)
	filelstname=filepathname&filename1&filenameend(n)

	if file.fileSize>512000 then
		response.write "<script language='javascript'>"
		response.write "alert('���ϴ����ļ�̫���ϴ����ɹ��������ļ�����ܳ���500K��');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	end if
	
	if LCase(filenameend(n))<>"gif" and LCase(filenameend(n))<>"jpg" then
		response.write "<script language='javascript'>"
		response.write "alert('�������ϴ���ѡ����ļ���ʽ������������ϴ���');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	end if
	
	file.saveAs Server.mappath(filename)
	session("tupian")=filename1&filenameend(n)
	response.Write("<script>window.alert('�ļ��ϴ��ɹ�!');window.location.href='upfile.asp';</script>")
else
	response.Write("<script>window.alert('����ѡ����Ҫ�ϴ����ļ�');window.close();</script>")
end if
set file=nothing
set upload=nothing
%>
</body> 
</html>