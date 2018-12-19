<%
'¼ìË÷·Ç·¨×Ö·û
Function  Switch(Str)
	Dim temp,num,feifa,ffchar,zu
	
	For i=1 To Len(Str)
		temp = Mid(Str,i,1)
		num = Asc(temp)
		If((num >= 32 and num <= 47) or (58 =< num and num <= 64) or (91 =< num and num <= 96) or (123 =< num and num <= 126))then
			if(InStr(feifa,temp) <= 0)then
				feifa = feifa&temp
			end if
		End If
	Next
	
	For j=1 To Len(feifa)
		ffchar = Mid(feifa,j,1)
		zu = Split(Str,ffchar)
		Str = "" 
		For k=0 To UBound(zu)		
			Str = Str&zu(k)
		Next		
	Next
	Switch = Str
End Function
%>
