<%
'�����Ƿ��ַ�
Function  Switch(Str)					'�����Զ��庯��
	Dim temp,num,feifa,ffchar,zu		'�������
	For i=1 To Len(Str)					
		temp = Mid(Str,i,1)				'���ַ����з���ָ����Ŀ���ַ�
		num = Asc(temp)					'�������ַ����ĵ�һ����ĸ��Ӧ�� ANSI �ַ�����
		If((num >= 32 and num <= 47) or (58 =< num and num <= 64) or (91 =< num and num <= 96) or (123 =< num and num <= 126))then
			if(InStr(feifa,temp) <= 0)then	'����ĳ�ַ�������һ�ַ����е�һ�γ��ֵ�λ��
				feifa = feifa&temp			'Ϊfeifa������ֵ
			end if
		End If
	Next
	For j=1 To Len(feifa)
		ffchar = Mid(feifa,j,1)
		zu = Split(Str,ffchar)			'���ػ��� 0 ��һά����
		Str = "" 						'��Str������ֵΪ��
		For k=0 To UBound(zu)			
			Str = Str&zu(k)				'ΪStr������ֵ
		Next		
	Next
	Switch = Str						'ΪSwitch������ֵ
End Function
%>
