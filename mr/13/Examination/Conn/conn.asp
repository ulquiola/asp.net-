<%
'---------- 防止SQL注入 -----------
dim SQL_Injdata 
SQL_Injdata = "'|;|and|exec|insert|select|delete|update|count|*|%|chr|mid|master|truncate|char|declare" 
SQL_inj = split(SQL_Injdata,"|") 
If Request.QueryString<>"" Then 
  For Each SQL_Get In Request.QueryString 
    For SQL_Data=0 To Ubound(SQL_inj) 
      if instr(Request.QueryString(SQL_Get),Sql_Inj(Sql_Data))>0 Then 
        Response.Redirect("/index.asp")
      end if 
    next 
  Next 
End If
'---------- 连接数据库 ----------
set conn=server.CreateObject("adodb.connection")
path=server.MapPath("/database/db_Examination.mdb")
conn.open "provider=microsoft.jet.oledb.4.0; data source="&path
%>