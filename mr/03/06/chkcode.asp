<%
Const stringarr = "0123456789ABCDEF" 
dim randarr,strlength
strlength = Len(stringarr)
Randomize 
For i = 1 To 4 
'// ���� Mid Rnd ����ÿ�δ������ַ����������ȡһ���ַ� 
randarr = randarr&Mid(stringarr,Int(Rnd * strlength + 1),1) 
Next 
response.write(randarr)
%>