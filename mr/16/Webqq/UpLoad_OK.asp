<!--#include file="adovbs.inc"-->
<%
Dim FormData,FormSize,DataStart,CLStr,DivStr
FormSize=Request.TotalBytes  'ȡ�ÿͻ�����Ӧ�����ݵ��ֽڴ�С
on error resume next 
FormData=Request.BinaryRead(FormSize)  '�Զ������뷽ʽ��ȡ�ͻ���POST����
CLStr=ChrB(13)&ChrB(10)
DataStart=InStrB(FormData,CLStr&CLStr)+4
'4�����Իس����з��ĳ���
DivStr=LeftB(FormData,InStrB(FormData,CLStr)-1)
DataSize=InStrB(DataStart+1,FormData,DivStr)-DataStart-2
FormData=MidB(FormData,DataStart,DataSize)
Application.Lock()  '����Application����
Application("Send_img")=FormData
Application.UnLock()  '����Application����
response.Redirect("send_index.asp")
%>

