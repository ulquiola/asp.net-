<%
'�滻�ո񡢻���
'==========================================================================
Function HTMLEncode(Char)
	Char=Replace(Char,chr(13)," ")
	Char=Replace(Char,chr(10) & chr(10),"<p></p>")
	Char=Replace(Char,chr(10),"<br>")
	HTMLEncode=Char
End Function
'==========================================================================

'�ж��Ƿ�Ϊ������
'==========================================================================
Function SafeRequest(Para)
	If Not IsNumeric(Para) Then
		Response.Write ("<script>alert('���� " & Para & " ����Ϊ�����ͣ�����ȷ������');history.back();</script>")
		Response.End
	End If
End Function
'==========================================================================

'��ʱ�䡢���ڲ������ظ��ĺ���
'==========================================================================
Function GetOrderNo(dDate)
    GetOrderNo = RIGHT("0000"+Trim(Year(dDate)),4)+RIGHT("00"+Trim(Month(dDate)),2)+RIGHT("00"+Trim(Day(dDate)),2)+RIGHT("00" + Trim(Hour(dDate)),2)+RIGHT("00"+Trim(Minute(dDate)),2)+RIGHT("00"+Trim(Second(dDate)),2)
End Function
'==========================================================================

'==========================================================================
'���ܣ����������ַ����е�Σ�շ���
function filter_Str(InString)
	NewStr=Replace(InString,"'","''")
	NewStr=Replace(NewStr,"<","&lt")
	NewStr=Replace(NewStr,">","&gt")
	NewStr=Replace(NewStr,"chr(60)","&lt;")
	NewStr=Replace(NewStr,"chr(37)","&gt;")
	NewStr=Replace(NewStr,"""","&quot")
	NewStr=Replace(NewStr,";",";;")
	NewStr=Replace(NewStr,"--","-")
	NewStr=Replace(NewStr,"/*"," ")
	NewStr=Replace(NewStr,"%"," ")
	filter_Str=NewStr
end function
'==========================================================================
%>