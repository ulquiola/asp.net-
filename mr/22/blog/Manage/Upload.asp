<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
</head>
<body>
<%
'��һ������·����ȡ���ļ�����
function getFileName(strPath)
	If strPath<>"" Then
		getFileName = mid(strPath,instrrev(strPath,"\")+1)
	Else
		response.Write("��ѡ���ϴ����ļ���")
		response.End()
	End If
end function
dim Obj_Come, Obj_Go, Data_B, Data_C
dim posB, posE, posSB, posSE,Enter 
dim strPath, strFileName, strContentType
Enter = chrb(13)&chrb(10)   '����һ�����ֽڵĻس����з�
set Obj_Come = server.CreateObject("adodb.stream")
Obj_Come.Type = 1  'ָ��������������Ϊ����������
Obj_Come.Mode = 3  'ָ����ģʽΪ��д
Obj_Come.Open 
Obj_Come.Write Request.BinaryRead(Request.TotalBytes)  '�Զ������뷽ʽ��ȡ�ͻ���POST����
Obj_Come.Position = 0
Data_B = Obj_Come.Read 
set Obj_Go = server.CreateObject("adodb.stream")
Obj_Go.Type = 1  'ָ��������������Ϊ����������
Obj_Go.Mode = 3  'ָ����ģʽΪ��д
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
	strPath = mid(Data_C,posSB,posSE-posSB)  '��ȡ�ļ�����·��
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
	strContentType = mid(Data_C,16)   '��ȡ�ļ�����
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
	'�����ļ������ϴ�����ͬ���ļ�����
	file_type=right(getFileName(strPath),3)
	If file_type="mp3" or file_type="MP3" or file_type="WAV"or file_type="wav" or file_type="MID" or file_type="mid" or file_type="RMI" or file_type="rmi" Then
	  If Obj_Go.size>51200 Then
	    Response.Write("<script>alert('�ϴ��ļ�����,�����ļ����ܳ���50K,�������ϴ�!');window.close();</script>")
		Response.End()
	  Else
	    Obj_Go.SaveToFile server.MapPath("/music")&"\" & filename1&file_type,2
	    Session("upfile_size")=Obj_Go.size
	  End IF
	Else
	  If Obj_Go.size>102400 Then
	    Response.Write("<script>alert('�ϴ��ļ�����,�����ļ����ܳ���100K,�������ϴ�!');window.close();</script>")
		Response.End()
	  Else
	    Obj_Go.SaveToFile server.MapPath("/Upfile")&"\" & filename1&file_type,2
	  End If 
	End If
	Session("pic")=filename1&file_type '�����ϴ����ļ���,�������Session������
	Obj_Go.Close 
End if
set Obj_Go = nothing
Obj_Come.Close 
%>
<script language="javascript">
alert("�ϴ��ɹ�!");
</script>
</body>
</html>
