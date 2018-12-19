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
Dim conn,connstr
Set conn=Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;User ID=admin;Password=;Data Source="&Server.MapPath("/DataBase/db_HomePage.mdb")&";"
conn.open connstr
'---------- 统计访问量 ----------
If Trim(Request.Cookies("visitor"))="" Then
'读取文件count.txt的数据
Set FSObject=Server.CreateObject("Scripting.FileSystemObject")
Set TextFile=FSObject.OpenTextFile(Server.MapPath("count.txt"))
If not TextFile.AtEndOfStream Then
num3=TextFile.ReadLine
End If
Set TextFile=Nothing
'向文件count.txt中追加数据
Set TextFile=FSObject.OpenTextFile(Server.MapPath("count.txt"),2,true)
TextFile.WriteLine num3+1
Set TextFile=Nothing
Set FSObject=Nothing
Response.Cookies("visitor")="visited"
Response.Cookies("visitor").Expires=DateAdd("d",1,now()) '指定Cookie的有效时间
End If
%>
