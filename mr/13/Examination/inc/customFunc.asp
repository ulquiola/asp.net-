<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
'*****************************
'功能:转换输入文本中的回车和空格
'调用方法:convert(输入字符串)
'返回值:字符
'*****************************
function convert(str)
	if str<>"" then
		str=replace(str,chr(13),"<br>")
		str=replace(str,chr(32),"&nbsp;")
	end if
	convert=str
end function
'*****************************
'功能:关闭记录集
'调用方法:closeRS(记录集名称)
'*****************************
function closeRS(rsName)
	rsName.close
	set rsName=nothing
end function
'*****************************
'功能:关闭数据库连接
'调用方法:closeConn(连接名称)
'*****************************
function closeConn(connName)
	connName.close
	set connName=nothing
end function
'*****************************
'功能:弹出提示对话框并重定向网页
'调用方法:message "提示内容","重定向网页的url地址"
'*****************************
function message(content,url)
	response.Write("<script language='javascript'>alert('"&content&"');window.location.href='"&url&"';</script>")
	response.End()
end function
'*****************************
'功能:打开指定大小的新窗口并设置窗口居中显示
'调用方法:call openwin(url,width,height)
'说明：如果w的值为0，则新窗口尺寸不做控制
'*****************************
function openwin(url,w,h,no)
	response.Write("<script language='javascript'>function openwin"&no&"(){")
	response.Write("if ("&w&"=='0'){var winhdc=window.open('"&url&"');")
	response.Write(	"var width=0;var height=0;}else{")
	response.Write("var winhdc=window.open('"&url&"','','width="&w&",height="&h&"');")
	response.Write("var width=(screen.width-"&w&")/2;")
	response.Write("var height=(screen.height-"&h&")/2;}")
	response.Write("winhdc.moveTo(width,height);")
	response.Write("}</script>")
end function
function filter_Str(InString)
'***********************************
'功能：过滤输入字符串中的危险符号
'调用方法：filter_Str("String")
'***********************************
	NewStr=Replace(InString,"'","''")
	NewStr=Replace(NewStr,"<","&lt;")
	NewStr=Replace(NewStr,">","&gt;")
	NewStr=Replace(NewStr,"chr(60)","&lt;")
	NewStr=Replace(NewStr,"chr(37)","&gt;")
	NewStr=Replace(NewStr,"""","&quot;")
	NewStr=Replace(NewStr,";",";;")
	NewStr=Replace(NewStr,"--","-")
	NewStr=Replace(NewStr,"/*"," ")
	NewStr=Replace(NewStr,"%"," ")
	filter_Str=NewStr
end function
%>
