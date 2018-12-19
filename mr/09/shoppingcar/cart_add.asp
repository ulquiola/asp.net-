<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="conn/conn.asp"--><!--数据库连接文件-->
<%
function cart(buyID,price,buynumber)		'创建自定义函数
	if not isarray(session("arr")) then		'判断session("arr")变量是否为数组
		dim arr(0,2)						'定义一个二维数组
		arr(0,0)=buyID						'为arr(0,0)数组赋值
		arr(0,1)=price
		arr(0,2)=buynumber
	else
		arr=session("arr")					'为arr变量赋值
		UB=ubound(arr,1)+1					'获取最大可用下标
		redim arr(UB,2)						'声明动态数组
		sessionarr=session("arr")			'为sessionarr变量赋值
		'将新添加的商品信息保存
		arr(UB,0)=buyID						'为数组赋值
		arr(UB,1)=price
		arr(UB,2)=buynumber
		'应用循环将保存在seesion("arr")数组中的商品信息添加到arr中
		For I = 0 To ubound(sessionarr,1)
		   For J = 0 To 2
			  arr(I, J)=sessionarr(I,J)		
		   Next
		Next	
	end if
	session("arr")=arr   					'为session("arr")变量赋值
end function
'保存到购物车
Q_spid=request.QueryString("goodsID")		'为Q_spid变量赋值
Q_DJ=1										'为Q_DJ变量赋值
Q_SL=1										'为Q_SL变量赋值
if Q_spid<>"" and Q_DJ>0 and Q_SL>0 then 	
	set rs=server.CreateObject("ADODB.RecordSet")	'创建记录集
	sql="select * from tb_goods where ID="&Q_spid	'查询数据
	rs.open sql,conn,1,3							'打开记录集
	Q_DJ=rs("nowPrice")								'为Q_DJ变量赋值
	if isarray(session("arr")) then					'判断session("arr")变量是否为数组
		arr=session("arr")							'为arr变量赋值
		For I = 0 To ubound(arr,1)					'获取数组的最大可用下标
			flag=false  '标记新添加的商品信息是否存在
			if Q_spid=arr(I,0) then  '当商品信息添加重复时
				arr(I,2)=cint(arr(I,2))+cint(Q_SL)  '累加商品数量
				session("arr")=arr 					'为session("arr")变量赋值
				flag=true	
				exit for							'退出循环
			end if
		Next
		if not flag then
			call cart(Q_spid,Q_DJ,Q_SL)				 '将商品信息添加至购物车
		end if
	else
		call cart(Q_spid,Q_DJ,Q_SL) 				'将商品信息添加至购物车
	end if
	response.Redirect("cart_see.asp")				'跳转到指定的ASP页面
else
%>
	<script language="javascript">
	alert("您的操作有误!");							//弾出提示对话框
	window.location.href="index.asp";				//跳转到指定的ASP页面
	</script>
<%end if%>
<title>添加至购物车</title>