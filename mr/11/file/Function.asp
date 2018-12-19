<%
Sub Login()
Response.write"<TABLE cellSpacing=0 cellPadding=0 width=200 align=center border=0><TBODY>"
Response.write"<TR><TD align='center'>"
If Session("UserName")="" Then
Response.write"<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
Response.write"<form name='form1' method='post' action='Login.asp'>"
Response.write"<tr><td width='34%' align='right'>用户名：</td>"
Response.write"<td width='66%'>"
Response.write"<input name='UserName' type='text' id='UserName' size='15' maxlength='20' style='height:20px;width:120px'>"
Response.write"</td></tr><tr><td align='right'>密　码：</td>"
Response.write"<td><input name='UserPwd' type='password' id='UserPwd' size='15' maxlength='20' style='height:20px;width:120px'>"
Response.write"</td></tr><tr><td height='3' colspan='2'></td></tr><tr>"
Response.write"<td><input name='action' type='hidden' id='action' value='login' /></td>"
Response.write"<td><input name='Login' type='submit' id='Login' value='登陆'>&nbsp;&nbsp;"
Response.write"<input name='reset' type='reset' id='reset' value='重写'>"
Response.write"</td></tr></form></table>"
Else
If Not IsObject(Conn) Then ConnectionDatabase()
Response.write"<table width='90%' border='0' cellspacing='0' cellpadding='0'>"
Response.write"<tr><td width='52%' align='center'>欢 迎 您：</td>"
Response.write"<td width='66%'>"&Session("UserName")&"</td></tr><tr>"
Response.write"<td align='center'>所 在 组：</td><td>"
Response.write GetGroupInfo(Session("GroupID"))(0)
Response.write"</td><tr><td align='center'>登陆次数：</td><td><b>"&Session("UserLogins")&"</b></td>"
Response.write"</tr><tr><td align='center'>上传文件：</td><td><font color=red>"
Response.write Conn.execute("Select UpFiles from tb_Users Where UserName='"&Session("UserName")&"'")(0)
Response.write"</font>个</td></tr><tr height=40><td align='right'>相 关 操 作：</td><td align='right'></td>"
Response.write"<tr><td colspan='2' align='center'><a href='UpLoad.asp'><font color=red>文 件 上 传</font></a></td></tr><tr><td  colspan='2' align='center'><a href='MyFile.asp'>我 的 文 件</a></td></tr>"
Response.write"<tr><td colspan='2' align='center'><a href='Edit.asp'>修 改 信 息</a></td></tr>"
If Session("UserAdmin") Then Response.Write"<tr><td colspan='2' align='center'><a href=Admin.asp><font color=blue>系 统 管 理</font></a></td></tr>"
Response.write"<tr><td colspan='2' align='center'><input name='Logout' type='button' id='Logout' value='退出'"
Response.write" onClick=""location.href='Login.asp?action=logout'"">"
Response.write"</td></tr></table>"
End If
Response.write"</TD></TR></TBODY></TABLE>"
End Sub

Function GetGroupInfo(GroupID)
dim rs,Str
if isnull(GroupID) then
Str="未加入组$$"
else
If Not IsObject(Conn) Then ConnectionDatabase()
set rs=Conn.execute("Select GroupName,UpFileType,UpFilesize from tb_Groups where GroupID="&clng(GroupID))
if rs.eof then
Str="未加入组$$"
else
Str=rs(0)&"$"&rs(1)&"$"&rs(2)
end if
set rs=nothing
end if
GetGroupInfo=split(Str,"$")
End Function

function HTMLEncode(fString)
if not isnull(fString) then
	fString = replace(fString, ">", "&gt;")
	fString = replace(fString, "<", "&lt;")
	fString = Replace(fString, CHR(32), "&nbsp;")
	fString = Replace(fString, CHR(9), "&nbsp;")
	fString = Replace(fString, CHR(34), "&quot;")
	fString = Replace(fString, CHR(39), "&#39;")
	fString = Replace(fString, vbCrLf & vbCrLf, "<P>")
	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10), "<BR>")

	HTMLEncode = fString
end if
end function
function getTopic(str,strlen)
	if str="" then
		getTopic=""
		exit function
	end if
	dim l,t,c, i
	l=len(str)
	t=0
	for i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		if c>255 then
			t=t+2
		else
			t=t+1
		end if
		if t>=strlen then
			getTopic=left(str,i) & "…"
			exit for
		else
			getTopic=str
		end if
	next
	getTopic=getTopic
end function

function myfilesize(xxx)
dim show11
show11=xxx & " Byte" 
if xxx>=1024 then
   xxx=(xxx/1024)
   show11=round(xxx,2) & " KB"
end if
if xxx>=1024 then
   xxx=(xxx/1024)
   show11=round(xxx,2) & " MB"
end if
myfilesize=show11
end function

function isInteger(para)
     on error resume next
     dim str
     dim l,i
     if isNUll(para) then 
        isInteger=false
        exit function
     end if
     str=cstr(para)
     if trim(str)="" then
        isInteger=false
        exit function
     end if
     l=len(str)
     for i=1 to l
         if mid(str,i,1)>"9" or mid(str,i,1)<"0" then
            isInteger=false 
            exit function
         end if
     next
     isInteger=true
     if err.number<>0 then err.clear
end function

function ChkPost()
dim server_v1,server_v2
	chkpost=false
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
	if mid(server_v1,8,len(server_v2))<>server_v2 then
		chkpost=false
	else
		chkpost=true
	end if
end function

Function Checkstr(Str)
		If Isnull(Str) Then
			CheckStr = ""
			Exit Function 
		End If
		Str = Replace(Str,Chr(0),"")
		CheckStr = Replace(Str,"'","''")
End Function

function IsValidTel(para)
	on error resume next
	dim str
	dim l,i
	if isNUll(para) then 
		IsValidTel=false
		exit function
	end if
	str=cstr(para)
	if len(trim(str))<7 then
		IsValidTel=false
		exit function
	end if
	l=len(str)
	for i=1 to l
		if not (mid(str,i,1)>="0" and mid(str,i,1)<="9" or mid(str,i,1)="-") then
			IsValidTel=false 
			exit function
		end if
	next
	IsValidTel=true
	if err.number<>0 then err.clear
end function
Public Function showpage(sfilename,totalnumber,maxperpage,CurrentPage,ShowTotal,ShowAllPages,strUnit)
	dim n, i,strTemp,strUrl
	if totalnumber mod maxperpage=0 then
    	n= totalnumber \ maxperpage
  	else
    	n= totalnumber \ maxperpage+1
  	end if
	strUrl=sfilename
  	strTemp= "<table align='center'><tr><td>"
	if ShowTotal=true then 
		strTemp=strTemp & "共 <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
  	if CurrentPage<2 then
    		strTemp=strTemp & "首页 上一页&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'>首页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>上一页</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "下一页 尾页"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>下一页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'>尾页</a>"
  	end if
   	strTemp=strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/页"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;转到：<select name='page' size='1' onchange=""javascript:window.location='" & strUrl & "page=" & "'+this.options[this.selectedIndex].value;"">"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">第" & i & "页</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</td></tr></table>"
	showpage= strTemp
end function
%>