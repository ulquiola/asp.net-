<%
function cart(buyID,price,buynumber)			'�����Զ��庯��
	if not isarray(session("arr")) then			'�ж�session("arr")�����Ƿ�������
		dim arr(0,2)							'��������
		arr(0,0)=buyID							'Ϊ���鸳ֵ
		arr(0,1)=price							'Ϊ���鸳ֵ
		arr(0,2)=buynumber
	else
		arr=session("arr")						'Ϊarr������ֵ
		UB=ubound(arr,1)+1						
		redim arr(UB,2)							'���嶯̬����
		sessionarr=session("arr")				'Ϊsessionarr������ֵ
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
	session("arr")=arr   	
end function
'���ù��ﳵ��������ʾ���ﳵ�е�����
'***********************************
'���ܣ��ӹ��ﳵ����ȥָ������Ʒ
'���÷�����call remove(Ҫɾ����Ʒ��ID��)
'***********************************
function remove(spid)
	if isarray(session("arr")) then
		arr=Session("arr")
		UB=Ubound(arr,1)
		if UB=0 then
			session("arr")=""
		else	
			flag=false '����Ƿ��Ѿ�ɾ����Ӧ����Ʒ��Ϣ
			redim arr_temp(UB-1,2)  '������ʱ�洢���ﳵ��Ϣ������
			for i=0 to UB
				if arr(i,0)<>spid then			
					for j=0 to 2
						if flag then  '��ɾ��һ����Ʒ��Ϣ��
							arr_temp(i-1,j)=arr(i,j)
						else
							arr_temp(i,j)=arr(i,j)
						end if
					next
				else
					flag=true
				end if
			next
		end if
	end if
	session("arr")=arr_temp
end function
%>
