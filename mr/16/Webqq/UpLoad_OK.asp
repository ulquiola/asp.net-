<!--#include file="adovbs.inc"-->
<%
Dim FormData,FormSize,DataStart,CLStr,DivStr
FormSize=Request.TotalBytes  '取得客户端响应的数据的字节大小
on error resume next 
FormData=Request.BinaryRead(FormSize)  '以二进制码方式读取客户端POST数据
CLStr=ChrB(13)&ChrB(10)
DataStart=InStrB(FormData,CLStr&CLStr)+4
'4是两对回车换行符的长度
DivStr=LeftB(FormData,InStrB(FormData,CLStr)-1)
DataSize=InStrB(DataStart+1,FormData,DivStr)-DataStart-2
FormData=MidB(FormData,DataStart,DataSize)
Application.Lock()  '锁定Application对象
Application("Send_img")=FormData
Application.UnLock()  '解锁Application对象
response.Redirect("send_index.asp")
%>

