<!--#include file=drawimg.asp-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns:v = "urn:schemas-microsoft-com:vml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<STYLE>
v\:* {
	BEHAVIOR: url(#default#VML)
}
</STYLE>
<title>无标题文档</title>
</head>

<body>

  <%

	item_v=request("v")
	dataType=request("dataType")
	if dataType="hours" then 
		t=array("0","1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16","17","18","19","20","21","22","23")
	end if
	if request("t")<>"" then
		t=split(request("t"),",")
	end if
	if request("d")<>"" then
		t=split(request("d"),",")
	end if

	v=split(item_v,",")

	if request("type")="shanxing" then 
		DrawArc t ,v
	elseif request("type")="zhuxing" then
		DrawRects t ,v
	else
		DrawLines t ,v
	end if
%>

</body>
</html>
