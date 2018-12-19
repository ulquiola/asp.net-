<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Application(Session("StuID"))=""
Session("StuID") = ""
Session("PassWord") = ""
Response.Write("<script language=JavaScript>window.open('index.asp','','toolbar,menubar,scrollbars,resizable,status,location,directories,copyhistory,height=400,width=600');window.close();</script>")
%>
