<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
function cart(buyID,price,buynumber)			'�����Զ��庯��
	if not isarray(session("arr")) then			'�ж�session("arr")�����Ƿ�������
		dim arr(0,2)							'��������
		arr(0,0)=buyID							'Ϊ���鸳ֵ
		arr(0,1)=price
		arr(0,2)=buynumber
	else
		arr=session("arr")						'Ϊarr������ֵ
		UB=ubound(arr,1)+1
		redim arr(UB,2)							'���嶯̬����
		sessionarr=session("arr")
		'������ӵ���Ʒ��Ϣ����
		arr(UB,0)=buyID
		arr(UB,1)=price
		arr(UB,2)=buynumber
		'Ӧ��ѭ����������seesion("arr")�����е���Ʒ��Ϣ��ӵ�arr��
		For I = 0 To ubound(sessionarr,1)
		   For J = 0 To 2
			  arr(I, J)=sessionarr(I,J)
		   Next
		Next	
	end if
	session("arr")=arr   						'Ϊsession������ֵ
end function
arr=session("arr")								'Ϊarr������ֵ
session("arr")=""								'���session����
For I = 0 To ubound(arr,1)
	sl="num"&I									'Ϊsl������ֵ
	spsl=request.Form(sl)						'Ϊspsl������ֵ
	call cart(arr(i,0),arr(i,1),spsl)			'�����Զ��庯��
Next
response.Redirect("cart_see.asp")				'��ת��ָ����aspҳ��
%>
