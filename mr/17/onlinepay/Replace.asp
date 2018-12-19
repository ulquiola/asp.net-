<%
'检索非法字符
Function  Switch(Str)					'创建自定义函数
	Dim temp,num,feifa,ffchar,zu		'定义变量
	For i=1 To Len(Str)					
		temp = Mid(Str,i,1)				'从字符串中返回指定数目的字符
		num = Asc(temp)					'返回与字符串的第一个字母对应的 ANSI 字符代码
		If((num >= 32 and num <= 47) or (58 =< num and num <= 64) or (91 =< num and num <= 96) or (123 =< num and num <= 126))then
			if(InStr(feifa,temp) <= 0)then	'返回某字符串在另一字符串中第一次出现的位置
				feifa = feifa&temp			'为feifa变量赋值
			end if
		End If
	Next
	For j=1 To Len(feifa)
		ffchar = Mid(feifa,j,1)
		zu = Split(Str,ffchar)			'返回基于 0 的一维数组
		Str = "" 						'将Str变量赋值为空
		For k=0 To UBound(zu)			
			Str = Str&zu(k)				'为Str变量赋值
		Next		
	Next
	Switch = Str						'为Switch变量赋值
End Function
%>
