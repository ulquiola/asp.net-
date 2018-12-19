<%
Const stringarr = "0123456789ABCDEF" 
dim randarr,strlength
strlength = Len(stringarr)
Randomize 
For i = 1 To 4 
'// 利用 Mid Rnd 函数每次从种子字符串中随机抽取一个字符 
randarr = randarr&Mid(stringarr,Int(Rnd * strlength + 1),1) 
Next 
response.write(randarr)
%>