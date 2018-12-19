<%
'函数作用：二进制转换为字符串
'参数
'binarystr：要转换的二进制代码
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
'函数作用：判断后缀
'参数
'chkstr：文件后缀
'chktype：允许的后缀，1或者2
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
	dtNow=Now()		'获取当前系统的日期
	randname=year(dtNow) & right("0" & month(dtNow),2) & right("0" & day(dtNow),2) & right("0" & hour(dtNow),2) & right("0" & minute(dtNow),2) & right("0" & second(dtNow),2)	
end function

'函数作用：上传文件
'参数
'uptype：上传类型。
'mixbyte：最大字节数
function upfilepath(uptype,mixbyte)
	filesize = request.TotalBytes
	if filesize > mixbyte then
		response.Write("<script>alert('上传文件太大');history.go(-1);</script>")
		response.End()
	end if
	filedata = request.BinaryRead(filesize)
	newline=chrB(13) & chrB(10)
	divider=leftB(filedata,clng(instrb(filedata,newline))-1)
	'文件名开始位置
	data1 = instrb(filedata,newline)+2
	'数据开始位置
	datastart = instrb(filedata,newline&newline)+4
	'数据结束位置
	dataend = instrb(datastart,filedata,divider)-len(newline)
	'真正想要的数据
	mydata=midb(filedata,datastart,dataend-datastart)
	'数据长度
	num = dataend - datastart
	
	''''''''''''''''''''''''''''''
	'获取文件名称
	''''''''''''''''''''''''''''''
	'filename="的二进制形式
	filename=chrb(102)&chrb(105)&chrb(108)&chrb(101)&chrb(110)&chrb(97)&chrb(109)&chrb(101)&chrb(61)&chrb(34)
	'Content-Type:的二进制形式
	filetype=chrb(67)&chrb(111)&chrb(110)&chrb(116)&chrb(101)&chrb(110)&chrb(116)&chrb(45)&chrb(84)&chrb(121)&chrb(112)&chrb(101)&chrb(58)
	'文件描述信息
	namedata = midb(filedata,data1,datastart)
	namestart = instrb(namedata,filename)+10
	nameend = instrb(namestart+2,namedata,newline)-1
	'上传文件名称信息
	nametemp = midb(namedata,namestart,nameend-namestart)
	'调用转换函数，将二进制转换为字符串
	nametemp = changename(nametemp)
	'firstline即上面获得的说明部分数据，chr(92) 表示"\" 
	nametmp=instrrev(nametemp,chr(92)) 
	getname=mid(nametemp,nametmp+1,lenb(nametemp)-nametmp-1) '获得文件名称 
	
	'上传文件类型信息
	typestart = instrb(namedata,filetype)+13
	typeend = instrb(typestart+2,namedata,newline&newline)
	'上传文件类型的二进制形式
	typetemp = midb(namedata,typestart,typeend-typestart)
	'调用转换函数，将二进制转换为字符串
	typetemp = changename(typetemp)
	'上传文件类型信息
	typetmp = instrrev(typetemp,chr(47))
	gettype = mid(typetemp,typetmp+1,lenb(typetemp)-typetmp-1)
	'判断文件后缀
	if chkpostfix(gettype,uptype) = false then
		response.Write("<script>alert('上传类型错误');history.go(-1);</script>")
		response.End()
	end if
	'复制文件
	set fromdata = server.CreateObject("adodb.stream")
	fromdata.Mode = 3
	fromdata.Type = 1
	fromdata.Open
	set senddata = server.CreateObject("adodb.stream")
	senddata.Mode = 3
	senddata.Type = 1
	senddata.Open
	
	'文件复制
	fromdata.Write filedata
	'文件开始位置
	fromdata.position = datastart-1
	'拷贝文件，num为文件长度
	fromdata.copyto senddata,num
	'保存文件路径
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