<%   
set conn=server.CreateObject("adodb.connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& server.MapPath("data/Stat.mdb")
conn.open connstr


dim pageUrl,userIp,lastpage,searchName,keyword,visitDate,visitTime,IsSearch,OpSystem,browser,screen
pageUrl=request("l")		'��ȡhtml�ķ��ʵ�ַ
userIp=GetIp()				'��ȡip
lastpage=request("r")		'��ȡ���ӵ�ַ
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
	sql="delete from online where DateDiff('s','"&visitDate&" "&visitTime&"',now())>600"'ɾ��ʮ����֮ǰ������
	conn.execute(sql)
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from online where  userIp='"&userIp&"' and pageUrl='"&pageUrl&"'"	'��online����ܲ�����ص�IP��¼
	rs.Open sql,conn,1,1
	if not rs.eof then	'����м�¼
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
'��ȡ�ͻ���IP
sub GetSearchName(pageUrl)
	if Instr(1, pageUrl, "baidu.com")>0 then
		searchName="�ٶ�"
		IsSearch=0
		CheckExp searchName,pageUrl 
	elseif Instr(1, pageUrl, "google.com")>0 then
		searchName="google"
		IsSearch=0
		CheckExp searchName,pageUrl 
	elseif Instr(1, pageUrl, "soso.com")>0 then
		searchName="����"
		IsSearch=0
		CheckExp searchName,pageUrl 
	elseif pageUrl="" then
		searchName="������վ"
		IsSearch=0
	else
		searchName="ֱ��������ַ"
		IsSearch=1
	end if

end sub

'��ȡ�ͻ���IP
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
	case "�ٶ�"
		reg.Pattern = "(wd=)(.*?)(&|$)" 
	case "google"
		reg.Pattern = "(q=)(.*?)(&|$)" 
	case "����"
		reg.Pattern = "(w=)(.*?)(&|$)" 
	end select 
	
	Set oMatches = reg.Execute(word)   
	For Each oMatch In oMatches    
	  keyword= oMatch.SubMatches(1)  
		  's   =reg.Replace(temp,GetMyInfo(oMatch.SubMatches(1)))   
	Next   
End sub 
%>