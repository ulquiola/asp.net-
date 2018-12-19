<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
</head>
<body>
<%
'从一个完整路径中取出文件名称
function getFileName(strPath)
	If strPath<>"" Then
		getFileName = mid(strPath,instrrev(strPath,"\")+1)
	Else
		response.Write("请选择上传的文件！")
		response.End()
	End If
end function
dim Obj_Come, Obj_Go, Data_B, Data_C
dim posB, posE, posSB, posSE,Enter 
dim strPath, strFileName, strContentType
Enter = chrb(13)&chrb(10)   '定义一个单字节的回车换行符
set Obj_Come = server.CreateObject("adodb.stream")
Obj_Come.Type = 1  '指定返回数据类型为二进制类型
Obj_Come.Mode = 3  '指定打开模式为读写
Obj_Come.Open 
Obj_Come.Write Request.BinaryRead(Request.TotalBytes)  '以二进制码方式读取客户端POST数据
Obj_Come.Position = 0
Data_B = Obj_Come.Read 
set Obj_Go = server.CreateObject("adodb.stream")
Obj_Go.Type = 1  '指定返回数据类型为二进制类型
Obj_Go.Mode = 3  '指定打开模式为读写
Obj_Go.Open 
posB = 1
posB = instrb(posB,Data_B,Enter)
posE = instrb(posB+1,Data_B,Enter)
Obj_Come.Position = posB+1
Obj_Come.CopyTo Obj_Go,posE-posB-2
Obj_Go.Position = 0
Obj_Go.Type = 2
Obj_Go.Charset = "gb2312"
Data_C = Obj_Go.ReadText 
Obj_Go.Close 
posSB = 1
posSB = instr(posSB,Data_C,"filename=""") + len("filename=""")
posSE = instr(posSB,Data_C,"""")
if posSE > posSB then
	strPath = mid(Data_C,posSB,posSE-posSB)  '获取文件完整路径
	posB = posE
	posE = instrb(posB+1,Data_B,Enter)
	Obj_Go.Type = 1
	Obj_Go.Mode = 3
	Obj_Go.Open 
	Obj_Come.Position = posB
	Obj_Come.CopyTo Obj_Go,posE-posB-1
	Obj_Go.Position = 0
	Obj_Go.Type = 2
	Obj_Go.Charset = "gb2312"
	Data_C = Obj_Go.ReadText 
	Obj_Go.Close 
	strContentType = mid(Data_C,16)   '获取文件类型
	posB = posE+2
	posE = instrb(posB+1,Data_B,Enter)
	Obj_Go.Type = 1
	Obj_Go.Mode = 3
	Obj_Go.Open 
	Obj_Come.Position = posB+1
	Obj_Come.CopyTo Obj_Go,posE-posB-2
	Obj_Go.Position = 0
	Data_C = Obj_Go.Read 

	Randomize
	ranNum=int(90000*rnd)+10000
	dtNow=Now()
	filename1=year(dtNow) & right("0" & month(dtNow),2) & right("0" & day(dtNow),2) & right("0" & hour(dtNow),2) & right("0" & minute(dtNow),2) & right("0" & second(dtNow),2) & ranNum &"."
	'根据文件类型上传到不同的文件夹中
	file_type=right(getFileName(strPath),3)
	If file_type="mp3" or file_type="MP3" or file_type="WAV"or file_type="wav" or file_type="MID" or file_type="mid" or file_type="RMI" or file_type="rmi" Then
	  If Obj_Go.size>51200 Then
	    Response.Write("<script>alert('上传文件过大,单个文件不能超过50K,请重新上传!');window.close();</script>")
		Response.End()
	  Else
	    Obj_Go.SaveToFile server.MapPath("/music")&"\" & filename1&file_type,2
	    Session("upfile_size")=Obj_Go.size
	  End IF
	Else
	  If Obj_Go.size>102400 Then
	    Response.Write("<script>alert('上传文件过大,单个文件不能超过100K,请重新上传!');window.close();</script>")
		Response.End()
	  Else
	    Obj_Go.SaveToFile server.MapPath("/Upfile")&"\" & filename1&file_type,2
	  End If 
	End If
	Session("pic")=filename1&file_type '设置上传的文件名,并存放在Session变量中
	Obj_Go.Close 
End if
set Obj_Go = nothing
Obj_Come.Close 
%>
<script language="javascript">
alert("上传成功!");
</script>
</body>
</html>
