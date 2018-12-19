<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../../Conn/conn.asp"-->
<%
getChkBox = request("ChkBox")
getcondition = replace(trim(request("condition")),"'","''")
getkey = replace(trim(request("key")),"'","''")
getpageno = replace(trim(request("pageno")),"'","")
rssql = "delete from tb_stuResult where ID in ("&getChkBox&")"
conn.Execute(rssql)
response.Redirect("../adm_Mainright.asp?condition="&getcondition&"&key="&getkey)
%>
