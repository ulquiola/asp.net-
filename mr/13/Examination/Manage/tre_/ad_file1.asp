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
sql1="insert into tb_StuResult(����,����ID��,����,����,�Ա�,�꼶,�������,����ȼ�,�����ȼ�,�ۺϵȼ�,���ĵȼ�,����,����÷�,�����÷�,�ۺϵ÷�,���ĵ÷�,�ܷ�,�ۺ����,�ۺ����,�ۺ����,�ۺ����,�ۺ����) Values('"&rsRead.Fields("����")&"','"&rsRead.Fields("����ID��")&"','"&rsRead.Fields("����")&"','"&rsRead.Fields("����")&"','"&rsRead.Fields("�Ա�")&"','"&rsRead.Fields("�꼶")&"','"&rsRead.Fields("�������")&"','"&rsRead.Fields("����ȼ�")&"','"&rsRead.Fields("�����ȼ�")&"','"&rsRead.Fields("�ۺϵȼ�")&"','"&rsRead.Fields("���ĵȼ�")&"','"&rsRead.Fields("����")&"','"&rsRead.Fields("����÷�")&"','"&rsRead.Fields("�����÷�")&"','"& rsRead.Fields("�ۺϵ÷�")&"','"&rsRead.Fields("���ĵ÷�")&"','"&rsRead.Fields("�ܷ�")&"','"&rsRead.Fields("�ۺ����")&"','"&rsRead.Fields("�ۺ����")&"','"&rsRead.Fields("�ۺ����")&"','"&rsRead.Fields("�ۺ����")&"','"&rsRead.Fields("�ۺ����")&"')"
Conn1.Execute(sql1) 
rsread.movenext
Loop
conn1.close
set conn1 = nothing
conn2.close
set conn2 = nothing
response.write "<script>alert('����ɹ�');history.back();</script>"
%>