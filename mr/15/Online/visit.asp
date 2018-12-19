<%   
set conn=server.CreateObject("adodb.connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& server.MapPath("data/Stat.mdb")
conn.open connstr


dim pageUrl,userIp,lastpage,searchName,keyword,visitDate,visitTime,IsSearch,OpSystem,browser,screen
pageUrl=request("l")		'获取html的访问地址
userIp=GetIp()				'获取ip
lastpage=request("r")		'获取链接地址
GetSearchName(pageUrl)

visitDate=date()
visitTime=time()
OpSystem=request("aN")
browser=request("b")
screen=request("s")
Stat()
response.Write(pageUrl&" "&userIp&" "&lastpage&" "&visitDate&" "&vistiTime)
response.End()

function Stat()
	if pageUrl<>""  then
	sql="delete from online where DateDiff('s','"&visitDate&" "&visitTime&"',now())>600"'删除十分钟之前的数据
	conn.execute(sql)
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from online where  userIp='"&userIp&"' and pageUrl='"&pageUrl&"'"	'在online表汇总查找相关的IP记录
	rs.Open sql,conn,1,1
	if not rs.eof then	'如果有记录
		sql="update online set visitDate='"&date()&"',visitTime='"&time()&"' where userIp='"&userIp&"' and pageUrl='"&pageUrl&"'"
		conn.execute(sql)
	else
		sql=" insert into online (pageUrl,userIp,lastpage,searchName,keyword,visitDate,visitTime,IsSearch,OpSystem,browser,screen)values"&"('"&pageUrl&"','"&userIp&"','"&lastpage&"','"&searchName&"','"&keyword&"','"&visitDate&"','"&visitTime&"','"&IsSearch&"','"&OpSystem&"','"&browser&"','"&screen&"')"
		conn.execute(sql)

	end if
	sql=" insert into visitor (pageUrl,userIp,lastpage,searchName,keyword,visitDate,visitTime,IsSearch,OpSystem,browser,screen)values"&"('"&pageUrl&"','"&userIp&"','"&lastpage&"','"&searchName&"','"&keyword&"','"&visitDate&"','"&visitTime&"','"&IsSearch&"','"&OpSystem&"','"&browser&"','"&screen&"')"
	conn.execute(sql)
	end if 
end function
'获取客户端IP
sub GetSearchName(pageUrl)
	if Instr(1, pageUrl, "baidu.com")>0 then
		searchName="百度"
		IsSearch=0
		CheckExp searchName,pageUrl 
	elseif Instr(1, pageUrl, "google.com")>0 then
		searchName="google"
		IsSearch=0
		CheckExp searchName,pageUrl 
	elseif Instr(1, pageUrl, "soso.com")>0 then
		searchName="搜搜"
		IsSearch=0
		CheckExp searchName,pageUrl 
	elseif pageUrl="" then
		searchName="其他网站"
		IsSearch=0
	else
		searchName="直接输入网址"
		IsSearch=1
	end if

end sub

'获取客户端IP
function GetIp()
	dim realip,proxy
	realip = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
	proxy = Request.ServerVariables("REMOTE_ADDR")
	if realip = "" then
		GetIp = proxy
	else
		GetIp = realip
	end if
end function

sub CheckExp(searchType,word) 
	Dim reg,s   
	Set reg   = New RegExp   
	reg.IgnoreCase = True   
	reg.Global = true   
	
	select case searchType
	case "百度"
		reg.Pattern = "(wd=)(.*?)(&|$)" 
	case "google"
		reg.Pattern = "(q=)(.*?)(&|$)" 
	case "搜搜"
		reg.Pattern = "(w=)(.*?)(&|$)" 
	end select 
	
	Set oMatches = reg.Execute(word)   
	For Each oMatch In oMatches    
	  keyword= oMatch.SubMatches(1)  
		  's   =reg.Replace(temp,GetMyInfo(oMatch.SubMatches(1)))   
	Next   
End sub 
%>