<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Conn/conn.asp" -->
<!--#include file="js/Function.asp" -->
       <%
		Set rssml = Server.CreateObject("ADODB.Recordset")		'������¼��
		rssml.ActiveConnection = conn							'����Command�����������Ϣ
		rssml.Source = "SELECT * from tb_talk where (fuserid = "&request.QueryString("UserID")&" and ToUserID = " &request.Cookies("UserID")&") or (fuserid = "&request.Cookies("UserID")&" and ToUserID = " &request.QueryString("UserID")&") order by id desc"
		'��ѯ���ݿ���Ϣ
		rssml.CursorType = 0									'�����α�����
		rssml.CursorLocation = 2								'�����α��������
		rssml.LockType = 1										'RecordSet����Ϊֻ�������������κ��޸�
		rssml.Open()											'�򿪼�¼��
		%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script>
function clearmessage()
{
	window.frames("frmrefrash").location = "clearmessage.asp?UserID=<%=request.QueryString("UserID")%>"
}
</script>
<style type="text/css">
<!--
.STYLE2 {
	font-size: 9pt;
	color: #000000;
}
-->
</style>
</head>

<body topmargin="0">
<div id="Layer1" style="position:absolute; width:200px; height:115px; z-index:1; visibility: hidden;">
<iframe name="frmrefrash" src=""></iframe>
</div>
<span class="STYLE2">[<a href="javascript:clearmessage()">������������¼</a>]</span> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="STYLE2">[ <a href="#" onClick="window.clipboardData.setData('Text',document.getElementById('content').value);">���Ƶ�ǰ�����¼</a> ]</span><br>
<span class="STYLE2">
<%
while not rssml.EOF								'��ѯ�Ƿ��м�¼��Ϣ
mes=rssml.Fields.Item("Message").value			'��ȡ�ֶ���Ϣ
%>	
<font color="#FF0000"><%=rssml("SendDate")%>&nbsp;<%=rssml("fusernickname")%> </font><br>
<%'Ӧ���������ɫ��ʾ������Ϣ���ں��û��ǳ�%>
<font color="#3333FF"><%=mes%> </font><br>
<%'Ӧ���������ɫ��ʾ�����¼��Ϣ%>
<%
content=content&rssml("SendDate")&"&nbsp;&nbsp;"&rssml("fusernickname")&chr(13)&mes&chr(13)&chr(13)
'Ϊcontent������ֵ
rssml.movenext()
'�����ƶ���¼ָ��
wend
rssml.close()
'�رռ�¼��
set rssml=nothing
'�ͷ��ڴ�ռ�
%>
</span>
<input type="hidden" name="content" id="content" value="<%=content%>">
</body>
</html>
