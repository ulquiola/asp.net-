<%
'�������ã�������ת��Ϊ�ַ���
'����
'binarystr��Ҫת���Ķ����ƴ���
function changename(binarystr)
	for i = 0 to lenb(binarystr) - 1
		temp = midb(binarystr,i+1,1)
		if ascb(temp)>127 then
			changename = changename&chr(ascw(midb(binarystr,i+2,1)&temp))
			i = i + 1
		else
			changename = changename&chr(ascb(temp))
		end if 
	next
end function
'�������ã��жϺ�׺
'����
'chkstr���ļ���׺
'chktype������ĺ�׺��1����2
function chkpostfix(chkstr,chktype)
	tmparr = split(chktype,",")
	lennum = ubound(tmparr)
	chkpostfix = false
	for i=0 to lennum
		if chkstr = tmparr(i) then
			chkpostfix = true
		end if
	next
end function
function randname()
	Randomize
	dtNow=Now()		'��ȡ��ǰϵͳ������
	randname=year(dtNow) & right("0" & month(dtNow),2) & right("0" & day(dtNow),2) & right("0" & hour(dtNow),2) & right("0" & minute(dtNow),2) & right("0" & second(dtNow),2)	
end function

'�������ã��ϴ��ļ�
'����
'uptype���ϴ����͡�
'mixbyte������ֽ���
function upfilepath(uptype,mixbyte)
	filesize = request.TotalBytes
	if filesize > mixbyte then
		response.Write("<script>alert('�ϴ��ļ�̫��');history.go(-1);</script>")
		response.End()
	end if
	filedata = request.BinaryRead(filesize)
	newline=chrB(13) & chrB(10)
	divider=leftB(filedata,clng(instrb(filedata,newline))-1)
	'�ļ�����ʼλ��
	data1 = instrb(filedata,newline)+2
	'���ݿ�ʼλ��
	datastart = instrb(filedata,newline&newline)+4
	'���ݽ���λ��
	dataend = instrb(datastart,filedata,divider)-len(newline)
	'������Ҫ������
	mydata=midb(filedata,datastart,dataend-datastart)
	'���ݳ���
	num = dataend - datastart
	
	''''''''''''''''''''''''''''''
	'��ȡ�ļ�����
	''''''''''''''''''''''''''''''
	'filename="�Ķ�������ʽ
	filename=chrb(102)&chrb(105)&chrb(108)&chrb(101)&chrb(110)&chrb(97)&chrb(109)&chrb(101)&chrb(61)&chrb(34)
	'Content-Type:�Ķ�������ʽ
	filetype=chrb(67)&chrb(111)&chrb(110)&chrb(116)&chrb(101)&chrb(110)&chrb(116)&chrb(45)&chrb(84)&chrb(121)&chrb(112)&chrb(101)&chrb(58)
	'�ļ�������Ϣ
	namedata = midb(filedata,data1,datastart)
	namestart = instrb(namedata,filename)+10
	nameend = instrb(namestart+2,namedata,newline)-1
	'�ϴ��ļ�������Ϣ
	nametemp = midb(namedata,namestart,nameend-namestart)
	'����ת����������������ת��Ϊ�ַ���
	nametemp = changename(nametemp)
	'firstline�������õ�˵���������ݣ�chr(92) ��ʾ"\" 
	nametmp=instrrev(nametemp,chr(92)) 
	getname=mid(nametemp,nametmp+1,lenb(nametemp)-nametmp-1) '����ļ����� 
	
	'�ϴ��ļ�������Ϣ
	typestart = instrb(namedata,filetype)+13
	typeend = instrb(typestart+2,namedata,newline&newline)
	'�ϴ��ļ����͵Ķ�������ʽ
	typetemp = midb(namedata,typestart,typeend-typestart)
	'����ת����������������ת��Ϊ�ַ���
	typetemp = changename(typetemp)
	'�ϴ��ļ�������Ϣ
	typetmp = instrrev(typetemp,chr(47))
	gettype = mid(typetemp,typetmp+1,lenb(typetemp)-typetmp-1)
	'�ж��ļ���׺
	if chkpostfix(gettype,uptype) = false then
		response.Write("<script>alert('�ϴ����ʹ���');history.go(-1);</script>")
		response.End()
	end if
	'�����ļ�
	set fromdata = server.CreateObject("adodb.stream")
	fromdata.Mode = 3
	fromdata.Type = 1
	fromdata.Open
	set senddata = server.CreateObject("adodb.stream")
	senddata.Mode = 3
	senddata.Type = 1
	senddata.Open
	
	'�ļ�����
	fromdata.Write filedata
	'�ļ���ʼλ��
	fromdata.position = datastart-1
	'�����ļ���numΪ�ļ�����
	fromdata.copyto senddata,num
	'�����ļ�·��
	filepath = randname&getname
	savepath = server.MapPath(""&left(request.ServerVariables("SCRIPT_NAME"),instrrev(request.ServerVariables("SCRIPT_NAME"),"/")))&"/selfpic/"&filepath
	senddata.savetoFile savepath,2
	senddata.close
	set senddata = nothing
		fromdata. Close 
	Set fromdata = nothing 
	upfilepath = filepath
end function
%>