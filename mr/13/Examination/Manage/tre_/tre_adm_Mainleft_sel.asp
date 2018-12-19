<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
getsel_profession = request("sel_profession")
getsel_lesson = request("sel_lesson")
response.Write(getsel_profession&"<p>")
response.Write(getsel_lesson&"<p>")
%>
