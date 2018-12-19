<font color="#0000FF">在文件index02.asp中显示的文本:</font><P>
<%
	Response.Write "Session会话超时时间为:"&Session.Timeout&"<P>"
	Response.Write "ASP脚本超时时间为:"&Server.ScriptTimeout&"<P>"
%>
