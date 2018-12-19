<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../../../Conn/conn.asp"-->
<%
getChkBox = request("ChkBox")
getcondition = replace(trim(request("condition")),"'","''")
getkey = replace(trim(request("key")),"'","''")
getpageno = replace(trim(request("pageno")),"'","")
rssql = "delete from tb_Administrator where ID in ("&getChkBox&")"
conn.Execute(rssql)
Response.Redirect("../adm_Admin.asp?condition="&getcondition&"&key="&getkey)
%>