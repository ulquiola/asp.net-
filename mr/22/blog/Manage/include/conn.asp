<%
'---------- 防止SQL注入 -----------
dim SQL_Injdata 
SQL_Injdata = "'|;|and|exec|insert|select|delete|update|count|*|%|chr|mid|master|truncate|char|declare" 
SQL_inj = split(SQL_Injdata,"|") 

If Request.QueryString<>"" Then 
For Each SQL_Get In Request.QueryString 
For SQL_Data=0 To Ubound(SQL_inj) 
if instr(Request.QueryString(SQL_Get),Sql_Inj(Sql_Data))>0 Then 
Response.end 
end if 
next 
Next 
End If
'---------- 连接数据库 ----------
Dim conn,connstr
Set conn=Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;User ID=admin;Password=;Data Source="&Server.MapPath("/DataBase/db_HomePage.mdb")&";"
conn.open connstr
%>
