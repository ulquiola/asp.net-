<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
response.ContentType = "text/html;charset=gb2312"
A_listName=split(Application("UserName"),",")
A_head=split(Application("head"),",")
%>
<div  style="text-align:left; padding:12px;">目前共有<font color="#CC0000"><b><%=UBound(A_ListName)%></b></font>人在线</div>
<%
For i=0 To UBound(A_ListName)%>
<div style="text-align:left; text-indent: 25px;">
	<a href="bottom.asp?name=<%=A_ListName(i)%>" target="bottomFrame" onClick="return chkname('<%=session("UserName")%>','<%=A_ListName(i)%>')">
	<%HeadPath="headICO/"+A_head(i)%>
	<img src="<%=HeadPath%>" border="0" width="23" height="23">
	<%=A_ListName(i)%></a>
</div>
<%Next%>