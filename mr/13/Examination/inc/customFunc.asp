<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
'*****************************
'����:ת�������ı��еĻس��Ϳո�
'���÷���:convert(�����ַ���)
'����ֵ:�ַ�
'*****************************
function convert(str)
	if str<>"" then
		str=replace(str,chr(13),"<br>")
		str=replace(str,chr(32),"&nbsp;")
	end if
	convert=str
end function
'*****************************
'����:�رռ�¼��
'���÷���:closeRS(��¼������)
'*****************************
function closeRS(rsName)
	rsName.close
	set rsName=nothing
end function
'*****************************
'����:�ر����ݿ�����
'���÷���:closeConn(��������)
'*****************************
function closeConn(connName)
	connName.close
	set connName=nothing
end function
'*****************************
'����:������ʾ�Ի����ض�����ҳ
'���÷���:message "��ʾ����","�ض�����ҳ��url��ַ"
'*****************************
function message(content,url)
	response.Write("<script language='javascript'>alert('"&content&"');window.location.href='"&url&"';</script>")
	response.End()
end function
'*****************************
'����:��ָ����С���´��ڲ����ô��ھ�����ʾ
'���÷���:call openwin(url,width,height)
'˵�������w��ֵΪ0�����´��ڳߴ粻������
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
'���ܣ����������ַ����е�Σ�շ���
'���÷�����filter_Str("String")
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
