<%
function remove(spid)					'�����Զ��庯��
	if isarray(session("arr")) then		'�ж�session("arr")�����Ƿ�Ϊ����
		arr=Session("arr")				'Ϊarr������ֵ
		UB=Ubound(arr,1)				'��ȡ������������±�
		if UB=0 then
			session("arr")=""			'���session����
		else	
			flag=false     				'����Ƿ��Ѿ�ɾ����Ӧ����Ʒ��Ϣ
			redim arr_temp(UB-1,2)      '������ʱ�洢���ﳵ��Ϣ������
			for i=0 to UB
				if arr(i,0)<>spid then			
					for j=0 to 2
						if flag then     				'��ɾ��һ����Ʒ��Ϣ��
							arr_temp(i-1,j)=arr(i,j)
						else
							arr_temp(i,j)=arr(i,j)		'Ϊarr_temp(i,j)������ֵ
						end if
					next
				else
					flag=true							'Ϊflag������ֵ
				end if
			next
		end if
	end if
	session("arr")=arr_temp								'Ϊsession("arr")������ֵ
end function
'ɾ��ָ���ļ�¼
if request.QueryString("ID")<>"" then			'�ж�idֵ�Ƿ�Ϊ��
	call remove(request.QueryString("ID"))		'�����Զ��庯��
end if
response.Redirect("cart_see.asp")				'��ת��ָ����ASP��̬ҳ��
%>