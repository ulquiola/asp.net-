<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Server.ScriptTimeOut=900
dim conn1,conn2
set conn1=CreateObject("ADODB.Connection")
set conn2=CreateObject("ADODB.Connection")
conn1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=;Data Source="&server.mappath("../../database/db_Examination.mdb")
conn2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=;Extended properties=Excel 5.0;Data Source="&server.mappath("../../database/db_Examination.xls")

sql="select * FROM [Sheet1$]" 
Set rsread=CreateObject("adodb.recordset")
rsread.open sql,conn2,1,1
Do while not rsread.eof
sql1="insert into tb_StuResult(级别,考生ID号,考号,姓名,性别,年级,考点代码,口语等级,听力等级,综合等级,作文等级,总评,口语得分,听力得分,综合得分,作文得分,总分,综合题Ⅰ,综合题Ⅱ,综合题Ⅲ,综合题Ⅳ,综合题Ⅴ) Values('"&rsRead.Fields("级别")&"','"&rsRead.Fields("考生ID号")&"','"&rsRead.Fields("考号")&"','"&rsRead.Fields("姓名")&"','"&rsRead.Fields("性别")&"','"&rsRead.Fields("年级")&"','"&rsRead.Fields("考点代码")&"','"&rsRead.Fields("口语等级")&"','"&rsRead.Fields("听力等级")&"','"&rsRead.Fields("综合等级")&"','"&rsRead.Fields("作文等级")&"','"&rsRead.Fields("总评")&"','"&rsRead.Fields("口语得分")&"','"&rsRead.Fields("听力得分")&"','"& rsRead.Fields("综合得分")&"','"&rsRead.Fields("作文得分")&"','"&rsRead.Fields("总分")&"','"&rsRead.Fields("综合题Ⅰ")&"','"&rsRead.Fields("综合题Ⅱ")&"','"&rsRead.Fields("综合题Ⅲ")&"','"&rsRead.Fields("综合题Ⅳ")&"','"&rsRead.Fields("综合题Ⅴ")&"')"
Conn1.Execute(sql1) 
rsread.movenext
Loop
conn1.close
set conn1 = nothing
conn2.close
set conn2 = nothing
response.write "<script>alert('导入成功');history.back();</script>"
%>