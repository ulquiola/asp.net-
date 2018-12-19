<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
function cart(buyID,price,buynumber)			'创建自定义函数
	if not isarray(session("arr")) then			'判断session("arr")变量是否是数组
		dim arr(0,2)							'定义数组
		arr(0,0)=buyID							'为数组赋值
		arr(0,1)=price
		arr(0,2)=buynumber
	else
		arr=session("arr")						'为arr变量赋值
		UB=ubound(arr,1)+1
		redim arr(UB,2)							'定义动态数组
		sessionarr=session("arr")
		'将新添加的商品信息保存
		arr(UB,0)=buyID
		arr(UB,1)=price
		arr(UB,2)=buynumber
		'应用循环将保存在seesion("arr")数组中的商品信息添加到arr中
		For I = 0 To ubound(sessionarr,1)
		   For J = 0 To 2
			  arr(I, J)=sessionarr(I,J)
		   Next
		Next	
	end if
	session("arr")=arr   						'为session变量赋值
end function
arr=session("arr")								'为arr变量赋值
session("arr")=""								'清空session变量
For I = 0 To ubound(arr,1)
	sl="num"&I									'为sl变量赋值
	spsl=request.Form(sl)						'为spsl变量赋值
	call cart(arr(i,0),arr(i,1),spsl)			'调用自定义函数
Next
response.Redirect("cart_see.asp")				'跳转到指定的asp页面
%>
