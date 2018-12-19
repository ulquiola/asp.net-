<html>
<head>
<title>无组件上传类</title>
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
	
	const filepath="/upfile/"		'上传路径
	const filepathname = "/upfile/"
	dtNow=Now()
	randomize
	ranNum=int(90000*rnd)+10000
	filename1=year(dtNow) & right("0" & month(dtNow),2) & right("0" & day(dtNow),2) & right("0" & hour(dtNow),2) & right("0" & minute(dtNow),2) & right("0" & second(dtNow),2) & ranNum &"."&fileExt
	filename=filepath&filename1&filenameend(n)
	filelstname=filepathname&filename1&filenameend(n)

	if file.fileSize>512000 then
		response.write "<script language='javascript'>"
		response.write "alert('您上传的文件太大，上传不成功，单个文件最大不能超过500K！');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	end if
	
	if LCase(filenameend(n))<>"gif" and LCase(filenameend(n))<>"jpg" then
		response.write "<script language='javascript'>"
		response.write "alert('不允许上传您选择的文件格式，请检查后重新上传！');"
		response.write "location.href='javascript:history.go(-1)';"
		response.write "</script>"
		response.end
	end if
	
	file.saveAs Server.mappath(filename)
	session("tupian")=filename1&filenameend(n)
	response.Write("<script>window.alert('文件上传成功!');window.location.href='upfile.asp';</script>")
else
	response.Write("<script>window.alert('请先选择你要上传的文件');window.close();</script>")
end if
set file=nothing
set upload=nothing
%>
</body> 
</html>