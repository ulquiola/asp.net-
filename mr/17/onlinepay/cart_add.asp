<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="conn/conn.asp"--><!--���ݿ������ļ�-->
<%
function cart(buyID,price,buynumber)		'�����Զ��庯��
	if not isarray(session("arr")) then		'�ж�session("arr")�����Ƿ�Ϊ����
		dim arr(0,2)						'����һ����ά����
		arr(0,0)=buyID						'Ϊarr(0,0)���鸳ֵ
		arr(0,1)=price
		arr(0,2)=buynumber
	else
		arr=session("arr")					'Ϊarr������ֵ
		UB=ubound(arr,1)+1					'��ȡ�������±�
		redim arr(UB,2)						'������̬����
		sessionarr=session("arr")			'Ϊsessionarr������ֵ
		'������ӵ���Ʒ��Ϣ����
		arr(UB,0)=buyID						'Ϊ���鸳ֵ
		arr(UB,1)=price
		arr(UB,2)=buynumber
		'Ӧ��ѭ����������seesion("arr")�����е���Ʒ��Ϣ��ӵ�arr��
		For I = 0 To ubound(sessionarr,1)
		   For J = 0 To 2
			  arr(I, J)=sessionarr(I,J)		
		   Next
		Next	
	end if
	session("arr")=arr   					'Ϊsession("arr")������ֵ
end function
'���浽���ﳵ
Q_spid=request.QueryString("goodsID")		'ΪQ_spid������ֵ
Q_DJ=1										'ΪQ_DJ������ֵ
Q_SL=1										'ΪQ_SL������ֵ
if Q_spid<>"" and Q_DJ>0 and Q_SL>0 then 	
	set rs=server.CreateObject("ADODB.RecordSet")	'������¼��
	sql="select * from tb_goods where ID="&Q_spid	'��ѯ����
	rs.open sql,conn,1,3							'�򿪼�¼��
	Q_DJ=rs("nowPrice")								'ΪQ_DJ������ֵ
	if isarray(session("arr")) then					'�ж�session("arr")�����Ƿ�Ϊ����
		arr=session("arr")							'Ϊarr������ֵ
		For I = 0 To ubound(arr,1)					'��ȡ������������±�
			flag=false  '�������ӵ���Ʒ��Ϣ�Ƿ����
			if Q_spid=arr(I,0) then  '����Ʒ��Ϣ����ظ�ʱ
				arr(I,2)=cint(arr(I,2))+cint(Q_SL)  '�ۼ���Ʒ����
				session("arr")=arr 					'Ϊsession("arr")������ֵ
				flag=true	
				exit for							'�˳�ѭ��
			end if
		Next
		if not flag then
			call cart(Q_spid,Q_DJ,Q_SL)				 '����Ʒ��Ϣ��������ﳵ
		end if
	else
		call cart(Q_spid,Q_DJ,Q_SL) 				'����Ʒ��Ϣ��������ﳵ
	end if
	response.Redirect("cart_see.asp")				'��ת��ָ����ASPҳ��
else
%>
	<script language="javascript">
	alert("���Ĳ�������!");							//������ʾ�Ի���
	window.location.href="index.asp";				//��ת��ָ����ASPҳ��
	</script>
<%end if%>
<title>��������ﳵ</title>